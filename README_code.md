# Reproducibility Code

These scripts generated the public-data validation tables, manuscript figures and submission package.

- run_real_workflow.py
- run_spatial_axis_deep_validation.py
- run_independent_patient_and_spatial_quant.py
- build_final_submission_package.py
- escc_splice_workflow/ helper package
- project_config.yaml
- requirements.txt

Run order: deep validation, independent patient/spatial quantification, then final package build.
The workflow uses public data only and records executor/reviewer gates in output tables.
