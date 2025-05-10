# Grade Automática de Provas de Recuperação  
TCC – IF Sudeste MG • Campus Rio Pomba

Sistema em **Python** que, com **OR‑Tools / CP‑SAT**, monta automaticamente a
grade de provas de recuperação dos cursos técnicos, garantindo:

* uso apenas dos horários livres de cada curso;  
* no máximo **1 prova por horário** e **3 provas por dia** para cada aluno;  
* aplicação conjunta de uma mesma disciplina em cursos diferentes quando
  existir slot livre comum;  
* grade compacta — minimiza o horário mais tarde e o número total de slots
  utilizados.

A grade final é exportada em arquivos **.xlsx** já no layout oficial do campus
(dias nas colunas, faixas horárias nas linhas, células centralizadas; provas
separadas por `|`; 0 = livre sem prova, 1 = ocupado por aula).

---

## Estrutura de dados

### `Horarios.json`
```json
{
  "Agropecuaria": {
    "seg": [1,1,1,1,0,0,0,0],
    "ter": [1,1,1,1,1,1,1,1],
    "qua": [1,1,1,1,0,0,0,0],
    "qui": [1,1,1,1,1,1,1,0],
    "sex": [1,1,1,1,1,1,1,1]
  },
  "Informatica": { ... }
}
```
`0` = horário livre; `1` = aula normal.

### `AlunosEmRecuperacao.json`
```json
{
  "Agropecuaria": { 
    "0": ["fisica","matematica","historia"],
    "1": ["filosofia"], 
    "2": ["sociologia"]
  },
  "Informatica": { ... }
}
```
Contem a lista cursos, alunos e as disciplinas em que cada um pegou recuperação

---

## Execução

```bash
# 1. ambiente virtual
python -m venv venv
# Windows
venv\Scripts\activate
# Linux / macOS
source venv/bin/activate

# 2. dependências
pip install -r requirements.txt   # ortools, xlsxwriter

# 3. gerar grades
python main.py
```

Arquivos gerados:

```text
planilhas/Agropecuaria.xlsx
planilhas/Informatica.xlsx
...
```

---

## Principais arquivos

| Arquivo                     | Função                                            |
|-----------------------------|---------------------------------------------------|
| `main.py`                  | modelagem CP‑SAT e exportação XLSX                |
| `Horarios.json`            | matriz de horários livres/ocupados                |
| `AlunosEmRecuperacao.json` | disciplinas em recuperação por aluno              |
| `requirements.txt`         | dependências Python                               |
| `LICENSE`                  | licença MIT                                       |

---

## Dependências

* Python ≥ 3.8  
* **OR‑Tools** (Google)  
* **XlsxWriter**

---

## Licença

Distribuído sob a **MIT License** — uso livre, exigindo apenas manutenção do
aviso de copyright.

---

## Autor

Desenvolvido por **Eliam Xavier David** como Trabalho de Conclusão de Curso no **IF Sudeste MG – Campus Rio Pomba** (2025).
