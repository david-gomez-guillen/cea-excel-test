# Cancer screening (Excel model)

A teaching cost-effectiveness model of screening for a hypothetical cancer with a yearly screening
test, built entirely as an Excel workbook, `cea_model.xlsx`. The cancer, the test and the
treatments are generic: they stand for any cancer that screening can find early, when it is
easier to cure. The app runs the workbook itself:
every result is one of its formulas. Open the file in Excel or LibreOffice to follow any number
step by step. Its *Start here* sheet explains how to use it, and it includes exercises for
learners. All values are illustrative.

## Model structure

A Markov cohort model with yearly cycles. A cohort of 100,000 people is followed from age 40 for
60 years, through four states:

- **Healthy**: no cancer. The whole cohort starts here.
- **Early cancer**: a small cancer with few symptoms that has not been diagnosed yet. Each year
  it can be detected, either because of symptoms or, in the years a strategy screens, by the
  test. A detected cancer is treated and cured with probability `p.cure.early`, which sends the
  person back to *Healthy*. An undetected cancer can progress.
- **Advanced cancer**: diagnosed, either after progressing or after a failed treatment of an
  early cancer. It is treated every year and carries a yearly risk of death from the cancer.
- **Dead**: from the cancer or from other causes. Absorbing.

The probability of developing cancer and the probability of dying from other causes depend on
the age group. Within a year, death from other causes comes first, and the other events happen
to those who survive it. The probability that an early cancer is detected in a year is

`1 - (1 - p.detect.symptoms) × (1 - screened × screen.sensitivity)`

where `screened` is `screen.participation` in the years a strategy screens and 0 otherwise. That
is the only way screening changes the course of the disease. It also adds costs: a test for
each person screened, and a follow-up diagnostic test for each positive screen, including the
false positives among healthy people (`1 - screen.specificity`).

Costs and QALYs of a year are counted for the people in each state at the start of the year,
without a half-cycle correction. Treating an early cancer is a one-off cost in the year it is
diagnosed. Costs and health effects are discounted at separate rates.

## Strategies

- **No screening**: cancers are only diagnosed because of symptoms.
- **Screening 50-74**: everyone aged 50 to 74 is offered a screening test every year.
- **Screening 45-74**: the same, from age 45.

The strategies are defined in the workbook (Inputs sheet, section 8).

## Parameters

<!-- PARAMETERS -->

The probabilities of death from other causes and of cancer onset have one value per age group.
The other parameters are one cell of the workbook each, so they cannot differ between age
groups.

## Strata

Five-year age groups, from 40-44 to 95-99.

## Outputs

- `summary`: one row per strategy with the cost (`C`, the discounted cost per person over the
  lifetime) and the effect (`E`, the discounted QALYs per person).
- `outcomes`: life years, both discounted and undiscounted, and cancer deaths and cancers
  diagnosed early and at an advanced stage per 100,000 people.
- `incidence`: cancers diagnosed per 100,000 person-years in each age group. It rises when
  screening starts, because the test finds cancers that were already there.

## Calibration

The `standard` scheme calibrates `p.cancer.onset`, one value per age group, against the
diagnosed incidence that a cancer registry would report for a population without screening,
using the `no_screening` strategy. The targets are hypothetical and 10-25% above the incidence
the base values give. The error is the sum of the squared relative differences between simulated
and target incidence, so the young age groups, with few cancers, weigh as much as the old ones.
Parameter sets that the workbook rejects get an infinite error.

The strata cannot be fitted one at a time, for two reasons. Incidence counts diagnoses, and an
early cancer is usually diagnosed a few years after it starts, so the onset in one age group
shows up partly in the next. And the people who develop cancer in one age group are no longer
healthy in the following ones.

Each simulation recalculates the whole workbook, which takes about 3 seconds.
