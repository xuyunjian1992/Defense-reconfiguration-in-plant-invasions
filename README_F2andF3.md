# Figures 2, 3 and supplementary Figure S5

## Input and execution

`F2andF3.R` analyzes invasive versus native plants using `data/Table S2.xlsx`. Keep the script at the repository root and run the complete file:

```sh
Rscript F2andF3.R
```

It can also be sourced by its full path in R or RStudio. No helper file or RStudio project is required. Packages are `readxl`, `dplyr`, `stringr`, `ggplot2`, `writexl`, `tibble`, `scales`, `patchwork`, and `metafor`. The script checks dependencies without installing them. Optional `input_file_override`, `project_root_override`, and `output_root_override` settings are near the beginning of the script.

The first worksheet is read. Headers are detected automatically within the first twelve rows; the distributed Table S2 has headers in row 2. Header/trait aliases are normalized in memory and the workbook remains unchanged.

## Panel definitions

| Panel/file prefix | Predictor | Response | Paired-data treatment | Model |
| --- | --- | --- | --- | --- |
| F2_A | Plant carbon | Physical defense | RAW | Quadratic |
| F2_B | Plant carbon | Chemical defense | IQR_removed | Quadratic |
| F2_C | Plant nitrogen | Chemical defense | RAW | Quadratic |
| F2_D | Plant phosphorus | Chemical defense | IQR_removed | Quadratic |
| F3_A | Plant carbon | Phenolics | RAW | Linear |
| F3_B | Plant carbon | Lignin | RAW | Linear |
| F3_C | Plant nitrogen | Phenolics | IQR_removed | Linear |
| F3_D | Plant nitrogen | Tannins | RAW | Quadratic |

RAW means eligible paired observations without additional IQR exclusion. For IQR_removed panels, inclusive 1.5-IQR bounds are applied to both lnRR axes after pairing; this is a single, non-iterative filter. The numerical tables preserve IQR membership. Only the defined figure version is exported for each panel. Both composites use A, B / C, D order. The original F3 model-choice check is retained.

## Statistical definitions

Eligibility uses `ComparisonType = invasive_vs_native`, `UseForRegression = main`, `PairingLevel` in `exact/baseline`, the accepted data levels (`species_pair`, `treatment_level`, `population_pair`), leaf tissue and the numerical validity checks in the script. Lignin is treated as physical defense.

Both axes are `lnRR = log(Mi/Mc)`, with i/c denoting invasive/native plants. Nutrient effects within a SystemID are aggregated by inverse sampling variance. F2 pairs on SystemID and StudyID; F3 pairs on SystemID and retains the defense-record StudyID, using the nutrient StudyID as fallback. The code rejects a SystemID shared by different studies. Multiple defense observations within a matched system can produce multiple fitted rows; the fitted row count is not a study count.

Weighted least squares uses inverse response sampling variance, with the original numerical floor and no weight cap. `zX` and `zY` here denote unstandardized lnRR, not z scores. Linear models include an intercept and X; quadratic models also include X². Weighted R² is calculated from weighted residual and total sums of squares and retains the original lower bound of zero. Exported prediction tables contain full confidence limits for the fitted mean; figure-only ribbon clipping does not change the model or these tables.

## Companion forests and funnels

F3 left-hand forests pool original trait lnRR records underlying the pairs retained for that panel. In F3_C, the 62 eligible pairs become 58 after IQR filtering, and the source records underlying those 58 pairs supply the forests and funnels. There is no second IQR filter on those source lnRR records. Forest pooling retains the original REML, single-effect handling and fixed-effect fallback.

Forest stars refer to the pooled-effect test against zero (`* P < 0.05`, `** P < 0.01`, `*** P < 0.001`). Right-hand regressions display weighted R² only. Regression coefficient P values remain in the result tables and are separate from forest P values.

Each F3 panel has two funnels, nutrient first and defense trait second. The funnels use the corresponding forest source records with valid lnRR and positive sampling variance. Repeated nutrient or defense names across panels describe different paired subsets. SF5 arranges these eight funnels in row-major order in a 3 × 3 grid, with the ninth cell blank. It assembles separate analyses without pooling their data.

Funnel diagnostics use a univariate REML model, with the original fixed-effect fallback. With at least ten effect-size rows, the script calculates Egger regression, Begg rank correlation and trim-and-fill. Only observed effects are drawn; the estimated missing-effect count `k0` is reported without replacing forest estimates. Funnel P values are asymmetry diagnostics. In annotations, `n` counts effects and `k` counts studies. Begg warnings about tied ranks are recorded with the computed approximate P values.

The contour fills are white for P ≥ 0.10, light gray for 0.05 ≤ P < 0.10 and gray for P < 0.05, based on two-sided normal-theory tests about lnRR = 0. The 90%, 95% and 99% boundaries are retained. Single and composite funnels retain the supplementary-figure titles, colors and layout.

## Outputs and completion

| Folder under `F2_F3_results/` | Contents |
| --- | --- |
| `F2/` | Composite `F2`, four panel files `F2_A`–`F2_D`, `F2_results.xlsx`, model-summary CSV |
| `F3/` | Composite `F3`, four forest/regression panel files `F3_A`–`F3_D`, `F3_results.xlsx`, model-summary CSV |
| `F3_funnel/` | Eight individual funnels, composite `SF5`, diagnostic workbook/CSV/text |

The 19 figure designs are exported as PDF and 300-dpi PNG. The output root also contains a combined model-summary CSV, `panel_manifest.csv`, and `run_info.txt`. Numerical workbooks include coefficients, fitted records, predictions, filtering audits, forest source rows and model-choice records. Source Excel rows are `RowID + 2` for this input.

Use the output set only after `run_info.txt` reports `COMPLETE`. A failed or interrupted run can leave mixed old/new files. Close locked files, resolve the reported error and rerun successfully. A model or source-row self-check failure requires investigation; it should not be bypassed. The script reads/writes local files and records input checksums, warnings and package versions. Use a fresh output directory for the final release.
