# S2 Code

This archive contains the reproducible scripts and configuration files used for the public-data ESCC spatial-to-bulk validation workflow and package audit.

## Environment

- Python 3.12.13 was used for package rebuilding and audit in the local environment.
- Rscript 4.5.3 was available, but the reconstruction scripts included here are Python-based.

## Main scripts

- `scripts/run_spatial_axis_deep_validation.py`
- `scripts/run_independent_patient_and_spatial_quant.py`
- `scripts/build_final_submission_package.py`
- `scripts/audit_clean_plos_upload.py`

## Data

All primary datasets are public: TCGA/GDC/UCSC Xena, GEO GSE47404, GEO GSE53625, HRA003627, HRA008846 and GDSC2. This code archive does not contain large source datasets.
- `scripts/run_transferability_supplement.py`
- `scripts/assemble_plos_one_checked_submission.py`
