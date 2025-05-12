# Scheduler of make‑up exams – CP‑SAT  +  CSV por curso (matriz completa)

import json
import os
import xlsxwriter
from pathlib import Path
from itertools import combinations
from ortools.sat.python import cp_model

# ── entrada ──────────────────────────────────────────────
schedules      = json.loads(Path("Horarios.json").read_text(encoding="utf-8"))
recovery_raw   = json.loads(Path("AlunosEmRecuperacao.json").read_text(encoding="utf-8"))

DAYS           = ["seg", "ter", "qua", "qui", "sex", "seg", "ter", "qua", "qui", "sex"]
SLOTS_PER_DAY  = len(next(iter(schedules.values()))["seg"])
TOTAL_SLOTS    = len(DAYS) * SLOTS_PER_DAY

lin            = lambda d, p: d * SLOTS_PER_DAY + p
unlin          = lambda k: divmod(k, SLOTS_PER_DAY)

# ── util ──────────────────────────────────────────────
free_slots = {course: [lin(day_idx, period_idx)
                       for day_idx, day_name in enumerate(DAYS)
                       for period_idx, flag in enumerate(schedules[course][day_name]) if flag == 0]
              for course in schedules}

subjects_by_course, subjects_by_student = {c: set() for c in schedules}, {}
for course, students in recovery_raw.items():
    for student, disc_list in students.items():
        subj_set = set(disc_list)
        subjects_by_course[course] |= subj_set
        subjects_by_student[(course, student)] = subj_set

courses_by_subject = {}
for course, subj_set in subjects_by_course.items():
    for subj in subj_set:
        courses_by_subject.setdefault(subj, []).append(course)

daily_slot_ranges = [range(d * SLOTS_PER_DAY, (d + 1) * SLOTS_PER_DAY) for d in range(len(DAYS))]

# ── modelo ──────────────────────────────────────────────
model, exam_slot = cp_model.CpModel(), {}
for course in schedules:
    domain = cp_model.Domain.FromValues(sorted(free_slots[course]))
    for subj in subjects_by_course[course]:
        exam_slot[(course, subj)] = model.NewIntVarFromDomain(domain, f'{course}_{subj}')

for (course, student), subj_set in subjects_by_student.items():
    for subj1, subj2 in combinations(subj_set, 2):
        model.Add(exam_slot[(course, subj1)] != exam_slot[(course, subj2)])

bool_var = {}
for (course, subj), var in exam_slot.items():
    for day_idx, rng in enumerate(daily_slot_ranges):
        b = model.NewBoolVar(f'b_{course}_{subj}_{day_idx}')
        model.AddAllowedAssignments(
            [var, b],
            [(k, 1) for k in rng] + [(k, 0) for k in range(TOTAL_SLOTS) if k not in rng]
        )
        bool_var[(course, subj, day_idx)] = b

for (course, student), subj_set in subjects_by_student.items():
    for day_idx in range(len(DAYS)):
        model.Add(sum(bool_var[(course, subj, day_idx)] for subj in subj_set) <= 3)

for subj, course_list in courses_by_subject.items():
    for course1, course2 in combinations(course_list, 2):
        if set(free_slots[course1]) & set(free_slots[course2]):
            model.Add(exam_slot[(course1, subj)] == exam_slot[(course2, subj)])

latest_slot = model.NewIntVar(0, TOTAL_SLOTS - 1, "latest_slot")
model.AddMaxEquality(latest_slot, [exam_slot[k] for k in exam_slot])
model.Minimize(latest_slot)

# ── solver ──────────────────────────────────────────────
solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = 10
if solver.Solve(model) not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    raise SystemExit("Sem solução viável")

# ── grade ──────────────────────────────────────────────
exam_schedule = {c: [[] for _ in range(TOTAL_SLOTS)] for c in schedules}
for (course, subj), var in exam_slot.items():
    exam_schedule[course][solver.Value(var)].append(subj)

# ── impressão no console ──────────────────────────────────────────────
for course in schedules:
    print(f'\n{course}:')
    for day_idx, day_name in enumerate(DAYS):
        row_cells = [', '.join(cell) if cell else '-'
                     for cell in exam_schedule[course][day_idx * SLOTS_PER_DAY:(day_idx + 1) * SLOTS_PER_DAY]]
        print(f'  {day_name:>3}: {row_cells}')

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
]  # deve ter SLOTS_PER_DAY entradas (necessário ajustar caso seja diferente)

os.makedirs("planilhas", exist_ok=True)

for course in schedules:
    wb = xlsxwriter.Workbook(f'planilhas/{course}.xlsx')
    ws = wb.add_worksheet('grade')

    # ── formatos ──────────────────────────────────────────────────────
    hdr_day = wb.add_format({'bold': True, 'align': 'center',
                             'valign': 'vcenter', 'border': 1,
                             'bg_color': '#D9D9D9'})
    hdr_time = wb.add_format({'align': 'center', 'valign': 'vcenter',
                              'border': 1, 'bg_color': '#D9D9D9'})
    fmt_cell = wb.add_format({'align': 'center', 'valign': 'vcenter',
                              'border': 1, 'text_wrap': True})
    fmt_num = wb.add_format({'align': 'center', 'valign': 'vcenter',
                             'border': 1, 'num_format': '0'})
    # ── larguras/alturas ──────────────────────────────────────────────
    ws.set_column(0, 0, 15)                       # coluna de horários
    ws.set_column(1, len(DAYS), 22)               # uma por dia
    ws.set_row(0, 25)                             # cabeçalho

    # ── cabeçalho: dias na linha 0 ────────────────────────────────────
    ws.write(0, 0, '')                            # canto vazio
    for col, day in enumerate(DAYS, start=1):
        ws.write(0, col, day.capitalize(), hdr_day)

    # ── linhas de horários ────────────────────────────────────────────
    for row, label in enumerate(TIME_LABELS, start=1):
        ws.write(row, 0, label, hdr_time)         # primeira coluna = horário

    # ── conteúdo das células ──────────────────────────────────────────
    for day_idx, day in enumerate(DAYS):          # coluna = dia
        for period_idx in range(SLOTS_PER_DAY):   # linha   = horário
            row = period_idx + 1                  # +1 por causa do cabeçalho
            col = day_idx + 1
            k = day_idx * SLOTS_PER_DAY + period_idx
            exams_here = exam_schedule[course][k]

            if exams_here:                        # existe prova
                ws.write(row, col, ' | '.join(exams_here), fmt_cell)
            else:                                 # sem prova → 1 ocupado / 0 livre
                busy_flag = schedules[course][day][period_idx]
                ws.write_number(row, col, busy_flag, fmt_num)

    wb.close()