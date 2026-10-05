# THALASSA interface of the cancer screening teaching model.
#
# The model is the Excel workbook cea_model.xlsx: every calculation is a formula
# in it. This file only describes the model to the app and runs the workbook
# through cea_workbook.py and xlsx_model.py, which recalculate it with the Python
# package formulas (an implementation of Excel's calculation engine), called from
# R with reticulate. Parameter names, labels, bounds, base values, strategies and
# age groups are all read from the workbook (its Interface sheet lists them), so
# the workbook is the only place where they are set.
#
# reticulate uses the Python of the RETICULATE_PYTHON environment variable, which
# the app sets to the Python of the model's conda environment (environment.yml).

library(reticulate)

workbook <- import_from_path('cea_workbook', path='.')

# Compiling the workbook takes about 15 seconds, once per model load. Each run
# then recalculates it in about 3 seconds.
MODEL <- workbook$WorkbookModel('cea_model.xlsx')

# Python names use underscores and R names dots: p_cancer_onset in the workbook
# is p.cancer.onset here.
python.name <- function(name) gsub('.', '_', name, fixed=TRUE)
r.name <- function(name) gsub('_', '.', name, fixed=TRUE)

# Inputs of the workbook that are not offered as parameters. The cohort size only
# scales the counts (results are per person), the start age is fixed so that the
# strata always cover the same ages, and the willingness to pay only feeds the
# workbook's own conclusion, since the app applies its own.
NOT.PARAMETERS <- c('cohort_size', 'start_age', 'wtp')

# The age-dependent inputs, one value per age group (stratum).
STRATIFIED <- c('p_background_death', 'p_cancer_onset')

get.overview <- function() {
  # Markdown shown in the Overview tab. The text lives in overview.md, and the
  # table of parameters is built from the workbook, where the <!-- PARAMETERS -->
  # line is.
  params <- get.parameters()
  seen <- c()
  rows <- c('| Parameter | Base value | Group | Meaning |', '|---|---|---|---|')
  for (p in params) {
    if (p$name %in% seen) next
    seen <- c(seen, p$name)
    value <- if (p$name %in% r.name(STRATIFIED)) {
      'by age group'
    } else {
      format(p$base.value, big.mark=',', scientific=FALSE)
    }
    rows <- c(rows, sprintf('| `%s` | %s | %s | %s |', p$name, value, p$class, p$display.name))
  }
  text <- readLines('overview.md')
  text <- unlist(lapply(text, function(line) if (trimws(line) == '<!-- PARAMETERS -->') rows else line))
  return(paste(text, collapse='\n'))
}

get.model.settings <- function() {
  return(list(cost.unit='€', effectiveness.unit='QALY', run.initial.guess=TRUE))
}

# ---- Strategies and strata ------------------------------------------------------

# The strategies are the rows of section 8 of the Inputs sheet. They are run
# under a name made from the workbook's label: "Screening 50-74" is
# screening_50_74.
strategy.id <- function(label) gsub('^_|_$', '', gsub('[^a-z0-9]+', '_', tolower(label)))
WORKBOOK.STRATEGIES <- MODEL$base_strategies
names(WORKBOOK.STRATEGIES) <- sapply(WORKBOOK.STRATEGIES, function(s) strategy.id(s$name))

# The description of a strategy is the Description column of the strategy table.
# Its attributes are made from the columns the workbook runs it with (screens,
# first and last age screened), so they always say what the workbook does with
# it. A strategy that does not screen ignores the ages, and has none.
strategy.screens <- function(s) tolower(trimws(s$screens)) == 'yes'

get.strategies <- function() {
  return(lapply(names(WORKBOOK.STRATEGIES), function(id) {
    s <- WORKBOOK.STRATEGIES[[id]]
    attributes <- if (strategy.screens(s)) {
      list(screening='Yes', first.age=as.integer(s$first_age), last.age=as.integer(s$last_age))
    } else {
      list(screening='No')
    }
    list(name=id, display.name=s$name, description=s$description, attributes=attributes)
  }))
}

get.strategy.attributes <- function() {
  # What describes the strategies, shown as columns of the Strategies tab.
  # Whether a strategy screens is drawn as the shape of its point on the base
  # case plot, and the first age screened as its fill, which is what tells the
  # screening strategies apart.
  return(list(
    screening=list(label='Screening', plot='shape', values=c(No='diamond', Yes='circle')),
    first.age=list(label='First age screened', plot='fill', ordered=TRUE),
    last.age='Last age screened'
  ))
}

get.strata <- function() {
  # The age groups of table 7 of the Inputs sheet.
  return(unlist(MODEL$age_groups))
}

# ---- Parameters -------------------------------------------------------------------

# Constraints: each is called as f(par.name, params), with the name of the
# parameter declaring it and the whole parameter list, and returns TRUE when the
# values are acceptable or the message shown for them when they are not. The
# ranges of the Interface sheet need none: the app checks them itself, from the
# min.value and max.value of each parameter. A constraint that ties several
# parameters together is declared on all of them, so that the app highlights
# whichever the user changed. A parameter split into strata arrives as a list,
# hence the unlist().
utilities.in.order <- function(par.name, params) {
  u <- lapply(c(healthy='u.healthy', early='u.early', advanced='u.advanced'),
              function(name) unlist(params[[name]]))
  if (any(u$advanced > u$early) || any(u$early > u$healthy))
    sprintf(paste('Quality of life must not rise as the cancer advances',
                  '(healthy %s, early cancer %s, advanced cancer %s)'),
            format(u$healthy), format(u$early), format(u$advanced))
  else TRUE
}

# The distribution the PSA draws each input from. The workbook has no notion of
# one, so they are declared here. The discount rates declare none, which keeps
# them out of the PSA.
DISTRIBUTIONS <- c(
  p_background_death='beta', p_cancer_onset='beta',
  p_detect_symptoms='beta', p_progression='beta', p_cure_early='beta', p_death_advanced='beta',
  screen_participation='beta', screen_sensitivity='beta', screen_specificity='beta',
  cost_screen_test='gamma', cost_followup_test='gamma', cost_early_treatment='gamma',
  cost_advanced_year='gamma',
  u_healthy='beta', u_early='beta', u_advanced='beta'
)

parameter.descriptors <- function(py.name) {
  info <- MODEL$parameter_info[[py.name]]
  name <- r.name(py.name)
  base <- MODEL$base_parameters[[py.name]]
  descriptor <- list(name=name, display.name=info$label, class=info$section)
  if (py.name %in% names(DISTRIBUTIONS)) descriptor$distribution <- DISTRIBUTIONS[[py.name]]
  if (!is.null(info$min)) descriptor$min.value <- info$min
  if (!is.null(info$max)) descriptor$max.value <- info$max
  if (name %in% c('u.healthy', 'u.early', 'u.advanced')) {
    descriptor$constraints <- utilities.in.order
  }
  if (!py.name %in% STRATIFIED) return(list(c(descriptor, list(base.value=base))))
  # One descriptor per age group, each with its own base value.
  return(lapply(seq_along(get.strata()), function(i) {
    c(descriptor, list(base.value=base[[i]], stratum=get.strata()[i]))
  }))
}

MODEL.PARAMETERS <- do.call(c, lapply(setdiff(unlist(MODEL$parameter_names), NOT.PARAMETERS),
                                      parameter.descriptors))

get.parameters <- function() {
  return(MODEL.PARAMETERS)
}

# ---- Overview tab ------------------------------------------------------------------

get.model.states <- function() {
  # State diagram shown in the Overview tab. It describes the transition matrix of
  # the Trace sheets of the workbook and must be kept in sync with it.
  return(list(
    title='Model states',
    description='The four states of the cohort and the yearly transitions between them. Hover a state or a transition for what drives it.',
    nodes=data.frame(
      id=c('healthy', 'early', 'advanced', 'dead'),
      label=c('Healthy', 'Early cancer\n(undiagnosed)', 'Advanced cancer', 'Dead'),
      description=c(
        'No cancer. The whole cohort starts here, at age 40.',
        'A small cancer with few symptoms that has not been diagnosed. Screening looks for it.',
        'A cancer that has grown: diagnosed, treated every year, with a high risk of death.',
        'From any cause. Absorbing state.'
      )
    ),
    edges=data.frame(
      from=c('healthy', 'early', 'early', 'healthy', 'early', 'advanced'),
      to=c('early', 'healthy', 'advanced', 'dead', 'dead', 'dead'),
      label=c('Onset', 'Detected and cured', 'Progression', 'Death', 'Death', 'Death'),
      description=c(
        'p.cancer.onset of the age group, among those who survive other causes.',
        'Detected by symptoms (p.detect.symptoms) or, in screening years, by the test (screen.participation x screen.sensitivity), and cured (p.cure.early).',
        'Missed and progresses (p.progression), or detected but not cured (1 - p.cure.early).',
        'p.background.death of the age group.',
        'p.background.death of the age group.',
        'p.death.advanced, or p.background.death of the age group.'
      )
    )
  ))
}

# ---- Simulation ----------------------------------------------------------------------

to.python <- function(pars) {
  # The app gives a parameter split into strata as a list keyed by stratum. The
  # age-dependent inputs take one value per age group, in order. The others are a
  # single cell of the workbook: they can only be split into strata if every
  # stratum has the same value.
  args <- list()
  for (name in names(pars)) {
    py.name <- python.name(name)
    value <- pars[[name]]
    if (is.list(value)) value <- unname(unlist(value[get.strata()]))
    if (py.name %in% STRATIFIED) {
      if (length(value) == 1) value <- rep(value, length(get.strata()))
      value <- as.list(value)
    } else if (length(value) > 1) {
      if (any(value != value[1])) stop(name, ' is the same for every age group in this model')
      value <- value[1]
    }
    args[[py.name]] <- value
  }
  return(args)
}

as.numbers <- function(values) {
  # A Python None (an age group the cohort never reaches) arrives as NULL.
  sapply(values, function(v) if (is.null(v)) NA_real_ else v)
}

run.simulation <- function(strategies, pars) {
  unknown <- setdiff(strategies, names(WORKBOOK.STRATEGIES))
  if (length(unknown) > 0) stop('Unknown strategy: ', paste(unknown, collapse=', '))
  results <- MODEL$run(to.python(pars), unname(WORKBOOK.STRATEGIES[strategies]))
  if (results$check != 'All OK') {
    stop('The workbook rejects these parameters: an input is outside its valid range, or a transition probability outside [0, 1]')
  }

  # Everything is converted to plain R values: the app copies results between
  # processes, and a pointer to a Python object cannot be copied.
  rows <- results$summary
  value <- function(field) sapply(rows, function(r) r[[field]])
  summary <- data.frame(strategy=strategies, C=value('cost'), E=value('qalys'))
  outcomes <- data.frame(
    strategy=strategies,
    life.years=value('life_years'),
    life.years.undiscounted=value('life_years_undiscounted'),
    cancer.deaths.per.100k=value('cancer_deaths_per_100k'),
    diagnosed.early.per.100k=value('diagnosed_early_per_100k'),
    diagnosed.advanced.per.100k=value('diagnosed_advanced_per_100k')
  )
  incidence <- lapply(seq_along(strategies), function(i) {
    setNames(as.numbers(results$incidence[[rows[[i]]$strategy]]), get.strata())
  })
  names(incidence) <- strategies
  return(list(summary=summary, outcomes=outcomes, incidence=incidence))
}

# ---- Calibration ------------------------------------------------------------------------

get.calibration.schemes <- function() {
  return(list(
    standard=list(
      description='Diagnosed cancer incidence without screening',
      parameters='p.cancer.onset',
      # Every stratum is calibrated, so the vector the app builds from the base
      # values has one value per age group, in the order of get.strata().
      strata=get.strata(),
      target=list(
        # Cancers diagnosed per 100,000 person-years in a population without
        # screening, as a (hypothetical) cancer registry would report them. They
        # are 10-25% above what the base values give.
        `Diagnosed incidence`=list(
          `40-44`=15, `45-49`=30, `50-54`=50, `55-59`=75, `60-64`=105, `65-69`=140,
          `70-74`=175, `75-79`=205, `80-84`=225, `85-89`=230, `90-94`=220, `95-99`=200
        )
      ),
      initial_guess=unlist(MODEL$base_parameters$p_cancer_onset),
      error_function=calibration.error,
      latent_space_training_set=generate.training.dataset,
      other.plots=NULL
    )))
}

calibration.error <- function(pars, target) {
  calibration.strategy <- 'no_screening'
  # The target is a named list with one entry per stratum, so it is flattened into
  # a named numeric vector to match the simulated values by stratum name.
  target.inc <- unlist(target$`Diagnosed incidence`)
  result <- tryCatch({
    results <- run.simulation(calibration.strategy, pars)
    incidence <- results$incidence[[calibration.strategy]]
    # Relative differences, so that the young age groups, with few cancers, count
    # as much as the old ones. Only the strata present in the target contribute.
    list(error=sum(((incidence[names(target.inc)] - target.inc) / target.inc)^2),
         output=list(`Diagnosed incidence`=incidence))
  }, error=function(e) {
    # Parameter sets the workbook rejects get an infinite error, which the app
    # records as a failed attempt.
    incidence <- setNames(rep(NA, length(get.strata())), get.strata())
    list(error=Inf, output=list(`Diagnosed incidence`=incidence))
  })
  return(result)
}

generate.training.dataset <- function(initial_guess, n, ...) {
  # Samples for the latent space methods: the initial guess scaled by a uniform
  # random factor per parameter, kept between 0 and 1 so that every sample is a
  # valid probability.
  f.pars <- list(...)
  variation <- f.pars$variation

  n_params <- length(initial_guess)
  dataset <- matrix(NA, nrow=n, ncol=n_params)
  for (i in 1:n) {
    factors <- runif(n_params, min=1-variation, max=1+variation)
    dataset[i, ] <- pmin(1, pmax(0, initial_guess * factors))
  }
  dataset <- dataset[sample(nrow(dataset)), ]
  return(dataset)
}
