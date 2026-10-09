# Figures 4, 5 and supplementary Figure S6

## Input and execution

`F4andF5.R` analyzes introduced-range versus native-range populations using `data/Table S2.xlsx`. Keep the script at the repository root and run the complete file:

```sh
Rscript F4andF5.R
```

Alternatively, source the full script path in R or RStudio. No helper file or RStudio project is required. Optional input, project-root and output-root overrides are defined near the beginning of the script.

Required packages are `readxl`, `dplyr`, `tidyr`, `stringr`, `ggplot2`, `writexl`, `tibble`, `purrr`, `patchwork`, `clubSandwich`, `scales`, and `metafor`. The script checks availability without installing packages. It reads the first worksheet, detects the header row and normalizes recognized aliases in memory. The distributed Table S2 has 768 records, with headers in row 2; `SourceExcelRow = RowID + 2`. The workbook is not modified.

## Panel definitions

All displayed panels use M2-RAW: the signed-square-root and RMS-scaled differences described below, using eligible paired observations without IQR removal. IQR results are numerical sensitivity analyses.

| Panel | Predictor | Response | Model |
| --- | --- | --- | --- |
| F4A | Plant carbon | Physical defense | Linear |
| F4B | Plant carbon | Chemical defense | Quadratic |
| F4C | Plant nitrogen | Chemical defense | Quadratic |
| F4D | Plant phosphorus | Chemical defense | Linear |
| F5A | Plant carbon | Cellulose | Linear |
| F5B | Plant carbon | Leaf mass per area (LMA) | Linear |
| F5C | Plant carbon | Terpenoids | Quadratic |
| F5D | Plant nitrogen | Alkaloids | Quadratic |
| F5E | Plant nitrogen | Flavonoids | Quadratic |
| F5F | Plant nitrogen | Terpenoids | Quadratic |

The F5 specification defines panel order and forms. The result tables also record agreement with the implemented automatic model-form rule.

## Analysis definitions

**Direction and eligibility.** Mi is the introduced-range mean and Mc the native-range mean. The program retains `introduced_vs_native_range`, leaf tissue, `UseForRegression = main`, `PairingLevel` in `exact/baseline`, and the accepted data levels (`species_pair`, `treatment_level`, `population_pair`). Numerical checks require finite means/SDs/sample sizes, positive sample sizes and nonnegative SDs. Lignin is treated as physical defense. A SystemID spanning studies or exact duplicated analytical records stops execution. Magnitude-based QC flags are recorded without automatic deletion.

**Pairing and Cha.** Nutrient and defense observations are paired within StudyID and SystemID. Repeated nutrient measurements are combined by inverse-variance weighting; separate defense records remain separate. The mean difference is `d = Mi − Mc`, with variance `v_d = Sdi²/Ni + SDc²/Nc`. The transformation is `z = sign(d)√|d|`, with delta-method variance `v_d/[4 max(|d|, 10⁻⁸)]`. M2 divides transformed nutrient differences by their positive RMS over unique paired StudyID/SystemID combinations and transformed defense differences by their positive RMS within each defense trait in the paired data. Variances are divided by the squared corresponding scales. No mean is subtracted. These scaled differences are the Cha values used for regression.

**Regression and model form.** Linear and quadratic weighted least-squares models use an intercept and inverse transformed/scaled response-variance weights. Predictor variances are retained but are not fitted in an errors-in-variables model. Weighted R² is calculated from weighted residual and total sums of squares. The gray ribbon is a 95% confidence interval for the fitted WLS mean based on model covariance and residual-degrees-of-freedom t quantiles; it is not an observation prediction interval or CR2 confidence band.

The automatic rule retains the quadratic form only with at least ten rows and five SystemIDs, an AICc improvement of at least 2, a weighted R² improvement of at least 0.03, ordinary quadratic-coefficient P ≤ 0.05 and a vertex within the observed X 5th–95th percentile range. Otherwise, the linear form is retained, subject to fitting requirements. The fixed F5 forms are checked against this rule.

**Inference and sensitivity.** Ordinary coefficient tests and StudyID-clustered CR2/Satterthwaite tests are reported separately. F5 BH adjustment is applied across fitted candidate relationships within each method/data treatment, separately for ordinary and CR2 tests. CR2 and leave-one-study-out values remain unavailable when their fitting requirements fail. Leave-one-study-out analyses retain the original transformation/scales and model form. IQR sensitivity uses 1.5-IQR bounds on both transformed axes without rescaling. Weight sensitivities include capped weights and equal total study weights; F4 additionally reports defense-trait fixed effects and equal total trait weights. These numerical analyses do not replace the displayed RAW curves.

## Companion forests and funnels

F5 forests use the original observations underlying each paired regression. They pool `lnRR = log(Mi/Mc)`, with sampling variance `Sdi²/(Ni Mi²) + SDc²/(Nc Mc²)` and positive-mean requirements. Pooling uses REML with study and nested-system random intercepts, with the original univariate fallback. Forest stars test the pooled lnRR against zero (`* P < 0.05`, `** P < 0.01`, `*** P < 0.001`). The right regression displays weighted R² only. Total tau² is the sum of random-intercept variance components; total/component I² use total tau² plus typical sampling variance in their denominator. Table S3 also reports the heterogeneity Q statistics.

Each funnel uses the corresponding source records with valid lnRR and strictly positive sampling variance. The funnel model is univariate REML, with the original fixed-effect fallback; its Egger, Begg and trim-and-fill diagnostics differ from multilevel forest inference and CR2 regression tests. Funnels require at least three effects and the bias diagnostics at least ten. Begg uses the asymptotic calculation. Only observed effects are plotted; `k0` is reported without replacing primary forest estimates. `NA` denotes an unavailable result, not zero. In annotations, `n` counts effects and `k` counts studies.

With the current input, the F5F nutrient (N) funnel's Egger and trim-and-fill calculations do not converge and are reported as `NA`. These entries do not establish absence of asymmetry or a zero missing-effect count. They do not replace or invalidate the separately fitted WLS relationship or companion forest; their limitation should be retained when reporting the funnel diagnostics.

The contour fills are white for P ≥ 0.10, light gray for 0.05 ≤ P < 0.10 and gray for P < 0.05, based on two-sided normal-theory tests about lnRR = 0. The 90%, 95% and 99% boundaries are retained. Single, paired and combined funnels retain the supplementary-figure presentation. SF6 uses a 4 × 3 grid in panel order A–F, with the nutrient then defense funnel for each relationship. It assembles twelve separate analyses without re-pooling effects.

## Outputs and completion

| Folder under `F4_F5_results/` | Contents |
| --- | --- |
| `F4/` | Composite `F4`, panels `F4A`–`F4D`, `F4_results.xlsx`, model-summary CSV |
| `F5/` | Composite `F5`, panels `F5A`–`F5F`, `F5_results.xlsx`, six-row `F5_model_summary.csv`, twelve-row companion-forest `Table_S3.csv` |
| `F5_funnel/` | Twelve individual funnels, six paired figures, composite `SF6`, diagnostic workbook and CSV |

All 31 designs are exported as vector PDF and 600-dpi TIFF. Funnel filenames identify both panel and trait, for example `F5A_C_funnel.pdf` and `F5A_Cellulose_funnel.pdf`; paired figures use `F5A_funnels.pdf`. The output root contains `panel_manifest.csv` and `run_info.txt`. Numerical workbooks retain model alternatives, paired/transformed records, source identifiers, sensitivity analyses, diagnostic checks and software information.

Use results only after `run_info.txt` reports `COMPLETE` and the internal checks have passed. A self-check failure can indicate missing or inconsistent panel/table output and must be investigated. A warning that a single-level random-effect component is fixed to zero records a model limitation and is not by itself a failed run. Reruns replace matching filenames but do not clear obsolete outputs; use a fresh directory for the final release and close file-locking viewers before execution.
