# cea-excel-test

A teaching cost-effectiveness model built as an Excel workbook: screening for a hypothetical
cancer with a yearly screening test, compared as no screening, screening from 50 to 74, and
screening from 45 to 74. The cancer, the test and the treatments are generic, and all values are
illustrative.

The workbook, [cea_model.xlsx](cea_model.xlsx), is the model. It works on its own in Excel or
LibreOffice, without macros. A learner can open it, read the *Start here* and *How it works*
sheets, change the blue inputs and watch every result recalculate. The same file can also be run
from code.

## The model

A four-state Markov cohort model with yearly cycles, from age 40 for 60 years:

- **Healthy** → **Early cancer** (undiagnosed): the onset probability depends on age.
- **Early cancer** → **Healthy**: the cancer is detected, by symptoms or screening, and cured.
- **Early cancer** → **Advanced cancer**: the cancer is missed and progresses, or treatment fails.
- **Advanced cancer** → **Dead**: death from the cancer.
- Every state → **Dead**: death from other causes, which also depends on age.

Screening raises the yearly probability of detecting an early cancer, which is
`1 - (1 - p_detect_symptoms) * (1 - screened * sensitivity)`. It also costs tests and
follow-up diagnostic tests, including the follow-up tests after false positives.

The differences from the model in `cea-model-test`:

- It has an undiagnosed early stage, which is what screening acts on, and an advanced stage.
- The background mortality and cancer onset tables depend on age.
- Screening applies only between the ages each strategy sets.
- Treating an early cancer has a one-off cost.
- Costs and health are discounted at separate rates.
- Strategies are compared with a full incremental analysis (dominance, extended dominance and
  net monetary benefit).

## The workbook

| Sheet | Contents |
|---|---|
| Start here | What the model is and how to use it, colour code, state diagram, glossary, 8 exercises |
| How it works | Every transition probability, cost and outcome in words and formulas, and the model's shortcuts |
| Inputs | All the parameters (blue cells), with units, base-case values, the names used in formulas and explanations |
| Trace 1–3 | One sheet per strategy, one row per year. Each row holds that year's probabilities, its transition matrix flattened row by row, the cohort (computed from the previous row), events, costs, QALYs and discounting. Column headers have comments explaining their formulas |
| Transition matrix | The 4×4 matrix of any strategy and age, and a check that cohort × matrix gives the trace's next row |
| Results | Per-person costs, QALYs and net monetary benefit; ICERs against strategy 1; the efficiency frontier; cost breakdown; cancer outcomes; diagnosed incidence by age group |
| Charts | Cost-effectiveness plane with the willingness-to-pay line, incidence by age, advanced cancer and cancer deaths over time, Markov trace |
| Interface | For code, not learners: which cells are the inputs and the results (see below) |

The formulas use named ranges (for example `=(1-D9)*E9`, `=F9*screen_participation`) so that they
read like the model's equations. Only functions from Excel 2007 are used, for compatibility with
LibreOffice and older versions of Excel. The file stores the result of every formula, so it shows
values even in Excel's Protected View and in file previews.

## Running it from code

[xlsx_model.py](xlsx_model.py) runs a workbook with the
[`formulas`](https://github.com/vinci1it2000/formulas) package, a Python implementation of the
Excel calculation engine. It writes the inputs into the input cells, recalculates and reads the
output cells. The model logic exists only in the xlsx, and the file on disk is never modified.

Which cells are inputs and outputs is declared in the workbook's **Interface** sheet, one row per
input or output:

| Column | Meaning |
|---|---|
| Name | What code calls it |
| Role | `input` or `output` |
| Cell | A cell, a range or a defined name. Blank: the defined name called Name |
| Label, Unit, Group | Text that describes it (optional; can be formulas pointing to the workbook's own labels) |
| Min, Max | Bounds of an input. Blank: those of the cell's data validation |
| Labels | A range or defined name with one label per row of a table, such as age groups (optional) |

The engine knows nothing else about the model, so any workbook can be run the same way: add an
Interface sheet (or, to leave the file untouched, write the same table as a CSV file) and

```python
from xlsx_model import XlsxModel

model = XlsxModel('other_model.xlsx')                 # or XlsxModel('other_model.xlsx', 'interface.csv')
model.inputs['p_cure'].base, model.inputs['p_cure'].max   # base value and bounds
model.run({'p_cure': 0.8})                           # {output name: value, list or list of rows}
```

[cea_workbook.py](cea_workbook.py) is the part specific to this workbook: it fills the strategy
slots and arranges the outputs into one row per strategy.

```python
from cea_workbook import WorkbookModel

model = WorkbookModel()          # compiles the workbook (about 15 s)
res = model.run({'screen_participation': 0.8,
                 'p_cancer_onset': [0.0003] * 12})       # about 3 s per run
res = model.run(strategies=[{'name': 'Screen 50-60', 'screens': 'Yes',
                             'first_age': 50, 'last_age': 60}])
res['summary']    # cost, qalys, life_years, nmb, icer_vs_first, cancer deaths... per strategy
res['frontier']   # strategies sorted by cost, with status and ICER
res['incidence']  # diagnosed cancers per 100,000 person-years, by age group
```

An age table can also be given as a dict keyed by age group, such as `{'40-44': 0.0003, ...}`.

Each run recalculates the whole workbook (about 16,000 cells), which takes about 3 s. That is fine
for scenarios and one-way sensitivity analysis. It is slow for calibration or a large
probabilistic sensitivity analysis, where a persistent LibreOffice process would be faster.

## In THALASSA

[thalassa_interface.R](thalassa_interface.R) describes the model to the app and calls
`cea_workbook.py` through `reticulate`, as `cea-python-test` does. The parameters (with their
labels, groups and bounds), base values, strategies and age-group strata are all read from the
workbook. The age-dependent onset and other-cause mortality are stratified parameters. The
`standard` calibration scheme fits `p.cancer.onset` to diagnosed incidence by age group.
[environment.yml](environment.yml) is the conda environment the app builds for the model.

```r
Sys.setenv(RETICULATE_PYTHON = "/path/to/python")   # a Python with the packages of environment.yml
source("thalassa_interface.R")                      # compiles the workbook, about 15 s
run.simulation(c("no_screening", "screening_50_74"), pars)$summary
```

## Files

Which files are needed depends on how the model is used:

| File | Excel only | From Python | In THALASSA | Purpose |
|---|:-:|:-:|:-:|---|
| `cea_model.xlsx` | ✓ | ✓ | ✓ | The model |
| `xlsx_model.py` | | ✓ | ✓ | Runs any workbook with an Interface sheet from code |
| `cea_workbook.py` | | ✓ | ✓ | Arranges the inputs and results of this workbook |
| `thalassa_interface.R` | | | ✓ | Describes the model to THALASSA and runs it |
| `overview.md` | | | ✓ | Text of the Overview tab |
| `environment.yml` | | | ✓ | Conda environment of the model (R, reticulate, Python, formulas) |

The other files are only for developing the model:

| File | Purpose |
|---|---|
| `build_workbook.py` | Writes the workbook with openpyxl, draws the state diagram and embeds the computed values |
| `reference_model.py` | An independent NumPy implementation of the same model. It is the source of the base values and is used to check the workbook |
| `test_workbook.py` | Compares the workbook against the reference: base case, random parameters, random strategies, dominance logic and input checks |
| `test_xlsx_model.py` | Tests of the generic engine on a small workbook it builds |
| `tests/run_tests.R` | Tests of the THALASSA interface |
| `assets/model_diagram.png` | State diagram embedded in the workbook |
| `requirements.txt` | Python packages for building and testing |

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python build_workbook.py               # regenerate cea_model.xlsx
.venv/bin/python -m pytest -q test_xlsx_model.py test_workbook.py   # about 2 min

conda env create -p ./env -f environment.yml
RETICULATE_PYTHON=./env/bin/python ./env/bin/Rscript tests/run_tests.R   # about 1 min
```

Edits made to the workbook by hand are overwritten when `build_workbook.py` runs. Make lasting
changes in the script.
