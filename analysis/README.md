# POUR analysis reproducibility code

This directory contains the analysis code accompanying the manuscript:

**Dual-time-point prediction of early postoperative urinary retention after colpocleisis for advanced pelvic organ prolapse: development and internal validation**

## Script

`POUR_analysis_reproducible_SciRep_v1.R`

The script implements the locked analysis workflow used for the manuscript, including:

- input and data-quality checks
- descriptive and univariable analyses
- multiple imputation by chained equations
- candidate-predictor screening
- multivariable logistic model development and Rubin pooling
- bootstrap internal validation
- ROC, calibration and decision-curve analyses
- nomogram construction
- ridge, elastic-net and LASSO sensitivity analyses

## Data

Patient-level clinical data are **not included** because they contain sensitive health information and cannot be deposited publicly under the study's ethical/privacy constraints.

For an authorized de-identified dataset, place:

`analysis_dataset.xlsx`

in:

`data/`

with worksheet:

`Sheet1`

The required variable names and coding are defined in the script.

## Online calculator

https://wangqi20190516.github.io/POUR-risk-calculator/

## Public-version changes

The public script differs from the locked research working script only in privacy protection, environment-independent file paths, dependency checks, publication-facing file labels, and reproducibility metadata. Statistical algorithms, predictor definitions, model-selection rules, random seeds and validation procedures were not intentionally changed.

## Privacy warning

Do **not** commit the real `data/` directory or the generated `outputs/` directory. Generated RDS, Excel and patient-level prediction files may contain sensitive clinical information.
