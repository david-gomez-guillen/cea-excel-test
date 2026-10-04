"""Run cea_model.xlsx from code, using the workbook itself as the model.

xlsx_model.py does the work: it reads the Interface sheet of the workbook, which
lists its input and output cells, compiles the workbook and recalculates it with
new inputs. This file only knows what this workbook's inputs and outputs mean:
it fills the strategy slots and arranges the results by strategy.

    from cea_workbook import WorkbookModel

    model = WorkbookModel()
    results = model.run({'screen_participation': 0.8, 'wtp': 30000})
    for row in results['summary']:
        print(row['strategy'], row['cost'], row['qalys'])

Parameters are the input names of the workbook (column "Name in formulas" of the
Inputs sheet): a number for a single input, or a list of 12 numbers (or a dict
keyed by age group) for the age tables `p_background_death` and
`p_cancer_onset`. Strategies can be replaced with a list of up to three dicts
with keys name, screens ('Yes' or 'No'), first_age and last_age.
"""

from pathlib import Path

from xlsx_model import XlsxModel

DEFAULT_PATH = Path(__file__).with_name('cea_model.xlsx')

# The inputs that describe the strategies, one cell per strategy slot.
STRATEGY_FIELDS = {'name': 'strategy_name', 'screens': 'strategy_screens',
                   'first_age': 'strategy_first_age', 'last_age': 'strategy_last_age'}
# The input whose labels are the age groups.
AGE_TABLE = 'p_cancer_onset'


class WorkbookModel:

    def __init__(self, path=DEFAULT_PATH):
        self.engine = XlsxModel(path)
        inputs = self.engine.inputs
        self.base_parameters = {nm: var.base for nm, var in inputs.items()}
        self.parameter_names = [nm for nm in inputs if nm not in STRATEGY_FIELDS.values()]
        self.parameter_info = {nm: {'label': inputs[nm].label, 'unit': inputs[nm].unit,
                                    'section': inputs[nm].group,
                                    'min': inputs[nm].min, 'max': inputs[nm].max}
                               for nm in self.parameter_names}
        self.age_groups = inputs[AGE_TABLE].labels
        self.n_slots = len(self.base_parameters['strategy_name'])

    @property
    def base_strategies(self):
        b = self.base_parameters
        return [{field: b[nm][i] for field, nm in STRATEGY_FIELDS.items()} for i in range(self.n_slots)]

    def run(self, parameters=None, strategies=None):
        """Recalculates the workbook with the given inputs and returns its results.

        Returns a dict with:
          summary:   one dict per strategy: cost, qalys, life_years and
                     life_years_undiscounted (per person, discounted unless said
                     otherwise), nmb, icer_vs_first, cancer_deaths_per_100k,
                     diagnosed_early_per_100k, diagnosed_advanced_per_100k.
          frontier:  the strategies sorted by cost, with status and ICER.
          incidence: per strategy, cancers diagnosed per 100,000 person-years
                     for each age group (None for groups the cohort never reaches).
          age_groups, best (the conclusion sentence) and check ('All OK' if the
          internal checks of the workbook pass).
        """
        values = dict(parameters or {})
        for nm in values:
            if nm in STRATEGY_FIELDS.values():
                raise KeyError(f'Unknown parameter: {nm} (pass strategies instead)')
        if strategies is not None:
            if not 1 <= len(strategies) <= self.n_slots:
                raise ValueError(f'The workbook compares 1 to {self.n_slots} strategies')
            # Unused slots repeat the first strategy; their results are dropped.
            padded = list(strategies) + [strategies[0]] * (self.n_slots - len(strategies))
            for field, nm in STRATEGY_FIELDS.items():
                values[nm] = [s[field] for s in padded]
        out = self.engine.run(values)

        n = len(strategies) if strategies is not None else self.n_slots
        names = out['out_strategy'][:n]
        columns = {'cost': 'out_cost', 'qalys': 'out_qalys', 'life_years': 'out_life_years',
                   'life_years_undiscounted': 'out_life_years_undiscounted', 'nmb': 'out_nmb',
                   'icer_vs_first': 'out_icer_vs_first', 'cancer_deaths_per_100k': 'out_cancer_deaths'}
        summary = []
        for i, s in enumerate(names):
            row = {'strategy': s}
            row.update({k: out[nm][i] for k, nm in columns.items()})
            # out_diagnoses has one row per strategy: early, advanced.
            row['diagnosed_early_per_100k'], row['diagnosed_advanced_per_100k'] = out['out_diagnoses'][i]
            summary.append(row)

        # out_incidence has one row per age group and one column per strategy.
        incidence = {s: [row[i] for row in out['out_incidence']] for i, s in enumerate(names)}

        frontier = []
        if n == self.n_slots:
            icers = [None] + out['out_frontier_icer'][1:]
            frontier = [{'strategy': s, 'status': st, 'icer': None if ic == '—' else ic}
                        for s, st, ic in zip(out['out_frontier_strategy'], out['out_frontier_status'], icers)]
        return {'summary': summary, 'frontier': frontier, 'incidence': incidence,
                'age_groups': self.age_groups, 'best': out['out_best'], 'check': out['out_check']}


if __name__ == '__main__':
    import time
    t = time.time()
    model = WorkbookModel()
    print(f'Compiled in {time.time() - t:.1f} s')
    t = time.time()
    res = model.run()
    print(f'Ran in {time.time() - t:.2f} s')
    print(res['check'], '|', res['best'])
    for row in res['summary']:
        print(row)
    for row in res['frontier']:
        print(row)
