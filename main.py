# Grade de provas – CP‑SAT  +  CSV por curso (matriz completa)

import json
from pathlib import Path
from itertools import combinations
from ortools.sat.python import cp_model
import xlsxwriter

# ── entrada ──────────────────────────────────────────────
horarios  = json.loads(Path("Horarios.json").read_text(encoding="utf-8"))
recup_raw = json.loads(Path("AlunosEmRecuperacao.json").read_text(encoding="utf-8"))

AGENDA    = ["seg", "ter", "qua", "qui", "sex", "seg", "ter", "qua", "qui", "sex"]
SLOTS_DAY = len(next(iter(horarios.values()))["seg"])
T_SLOTS   = len(AGENDA) * SLOTS_DAY

lin   = lambda d,p: d*SLOTS_DAY + p
unlin = lambda k: divmod(k, SLOTS_DAY)

# ── util ──────────────────────────────────────────────
free = {c: [lin(d,p)
            for d,n in enumerate(AGENDA)
            for p,v in enumerate(horarios[c][n]) if v == 0]
        for c in horarios}

disc_by_course, disc_by_student = {c:set() for c in horarios}, {}
for c,alunos in recup_raw.items():
    for a,dl in alunos.items():
        s=set(dl); disc_by_course[c]|=s; disc_by_student[(c,a)]=s

courses_by_disc={}
for c,ds in disc_by_course.items():
    for d in ds: courses_by_disc.setdefault(d,[]).append(c)

day_rng=[range(d*SLOTS_DAY,(d+1)*SLOTS_DAY) for d in range(len(AGENDA))]

# ── modelo ──────────────────────────────────────────────
m,slot=cp_model.CpModel(),{}
for c in horarios:
    dom=cp_model.Domain.FromValues(sorted(free[c]))
    for d in disc_by_course[c]:
        slot[(c,d)]=m.NewIntVarFromDomain(dom,f'{c}_{d}')

for (c,a),ds in disc_by_student.items():
    for d1,d2 in combinations(ds,2):
        m.Add(slot[(c,d1)]!=slot[(c,d2)])

bvar={}
for (c,d),v in slot.items():
    for di,rng in enumerate(day_rng):
        b=m.NewBoolVar(f'b_{c}_{d}_{di}')
        m.AddAllowedAssignments([v,b],
            [(k,1) for k in rng]+[(k,0) for k in range(T_SLOTS) if k not in rng])
        bvar[(c,d,di)]=b

for (c,a),ds in disc_by_student.items():
    for di in range(len(AGENDA)):
        m.Add(sum(bvar[(c,d,di)] for d in ds) <= 3)

for d,cs in courses_by_disc.items():
    for c1,c2 in combinations(cs,2):
        if set(free[c1]) & set(free[c2]):
            m.Add(slot[(c1,d)] == slot[(c2,d)])

latest=m.NewIntVar(0,T_SLOTS-1,"latest")
m.AddMaxEquality(latest,[slot[k] for k in slot])
m.Minimize(latest)

# ── solver ──────────────────────────────────────────────
s=cp_model.CpSolver(); s.parameters.max_time_in_seconds=10
if s.Solve(m) not in (cp_model.OPTIMAL,cp_model.FEASIBLE):
    raise SystemExit("Sem solução viável")

# ── grade ──────────────────────────────────────────────
grade={c:[[] for _ in range(T_SLOTS)] for c in horarios}
for (c,d),v in slot.items():
    grade[c][s.Value(v)].append(d)

# ── impressão no console ──────────────────────────────────────────────
for c in horarios:
    print(f'\n{c}:')
    for di,n in enumerate(AGENDA):
        linha=[', '.join(cel) if cel else '-' 
               for cel in grade[c][di*SLOTS_DAY:(di+1)*SLOTS_DAY]]
        print(f'  {n:>3}: {linha}')

# ── salvar em planilha xlsx ─────────────────────────────────────
TIME_LABELS = [
    "07:00 – 07:55",
    "07:55 – 08:50",
    "09:10 – 10:05",
    "10:05 – 11:00",
    "13:00 – 13:55",
    "13:55 – 14:50",
    "15:10 – 16:05",
    "16:05 – 17:00",
]  # deve ter SLOTS_DAY entradas (necessario ajustar caso seja diferente)

for c in horarios:
    wb  = xlsxwriter.Workbook(f'planilhas/{c}.xlsx')
    ws  = wb.add_worksheet('grade')

    # ── formatos ──────────────────────────────────────────────────────
    hdr_day = wb.add_format({'bold': True, 'align': 'center',
                             'valign': 'vcenter', 'border': 1,
                             'bg_color': '#D9D9D9'})
    hdr_time = wb.add_format({'align': 'center', 'valign': 'vcenter',
                              'border': 1, 'bg_color': '#D9D9D9'})
    fmt_cell = wb.add_format({'align': 'center', 'valign': 'vcenter',
                              'border': 1, 'text_wrap': True})
    fmt_num  = wb.add_format({'align': 'center', 'valign': 'vcenter',
                              'border': 1, 'num_format': '0'})
    # ── larguras/alturas ──────────────────────────────────────────────
    ws.set_column(0, 0, 15)                   # coluna de horários
    ws.set_column(1, len(AGENDA), 22)         # uma por dia
    ws.set_row(0, 25)                         # cabeçalho

    # ── cabeçalho: dias na linha 0 ────────────────────────────────────
    ws.write(0, 0, '')                        # canto vazio
    for col, day in enumerate(AGENDA, start=1):
        ws.write(0, col, day.capitalize(), hdr_day)

    # ── linhas de horários ────────────────────────────────────────────
    for row, tlabel in enumerate(TIME_LABELS, start=1):
        ws.write(row, 0, tlabel, hdr_time)    # primeira coluna = horário

    # ── conteúdo das células ──────────────────────────────────────────
    for di, day in enumerate(AGENDA):         # coluna = dia
        for aula in range(SLOTS_DAY):         # linha   = horário
            row = aula + 1                    # +1 por causa do cabeçalho
            col = di + 1
            k = di * SLOTS_DAY + aula
            provas = grade[c][k]

            if provas:                        # existe prova
                ws.write(row, col, ' | '.join(provas), fmt_cell)
            else:                             # sem prova → 1 ocupado / 0 livre
                ocupado = horarios[c][day][aula]
                ws.write_number(row, col, ocupado, fmt_num)

    wb.close()