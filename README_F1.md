# Figure 1 and trait-specific funnel plots

## Input and execution

Run `F1.R` from the repository root, with `data/F1_Data.xlsx` in place:

```sh
Rscript F1.R
```

Alternatively, source the complete script by its full path in R or RStudio. All analysis functions are contained in `F1.R`. Required packages are `readxl`, `dplyr`, `purrr`, `metafor`, `writexl`, `ggplot2`, and `scales`. Missing packages are installed from CRAN; this requires a suitable network connection or prior offline installation. PDF export requires Cairo support.

| Worksheet | Comparison | Defense category | Input rows |
| --- | --- | --- | ---: |
| F1_A | Invasive versus native plants | Physical defense | 504 |
| F1_B | Invasive versus native plants | Chemical defense | 896 |
| F1_C | Introduced-range versus native-range populations | Physical defense | 670 |
| F1_D | Introduced-range versus native-range populations | Chemical defense | 1670 |

Each sheet contains `StudyID`, `Variable`, `Mc`, `SDc`, `Nc`, `Mi`, `Sdi`, and `Ni`. Overall-defense and individual-trait blocks contain the same underlying observations grouped for different analyses; the full worksheet row count is not an independent-effect count. Each `Variable` is fitted separately. See the data dictionary for group directions and the preparation of this workbook from Table S1.

## Models and interpretation

The effect is `lnRR = log(Mi/Mc)`. Its SD-based sampling variance is `Sdi^2/[Ni*(Mi^2 + 1e-12)] + SDc^2/[Nc*(Mc^2 + 1e-12)]`. The original `1/Nc + 1/Ni` fallback remains enabled if the SD-based variance is unavailable or nonpositive; the row-level output identifies the formula used.

The primary forest model first attempts a StudyID random intercept under REML when there is more than one study and repeated effects within a study; otherwise, or if that fit fails, it uses univariate random-effects meta-analysis. Bias diagnostics use a separate univariate REML model. If Egger or Begg P < 0.05 and trim-and-fill succeeds, its estimate is used in the final forest; otherwise, the base estimate is retained. When the default trim-and-fill result has `k0 = 0`, the forest module also tries both sides and retains the largest estimated `k0`, preserving the specified analysis rule. Model tables report the base estimate, adjustment and final source.

Only the `forest_primary` version is exported. It is the same result definition previously named `primary_and_trimfill`; the shorter filename does not remove the adjustment rule. No separate original-estimate forest is exported.

The funnel module retains its own univariate REML estimates. It requires at least three valid effects for a funnel and at least ten effects for its bias tests and conditional trim-and-fill. It does not apply the forest module's forced left/right trial. Consequently, forest and funnel pooled estimates or adjustment outcomes need not coincide. Bias tests diagnose asymmetry; their P values are distinct from the forest pooled-effect test.

## Outputs

Results are written to `results/F1_Data/<sheet>/`. Names use the worksheet prefix, for example:

| Output | Contents |
| --- | --- |
| `F1_A_forest_primary.pdf/png` | Final forest plot |
| `F1_A_forest_results.xlsx` | Plotted estimates and analysis details |
| `F1_A_forest_primary_results.csv` | Final forest estimates |
| `F1_A_forest_primary_statistics.xlsx/csv` | Figure-ready estimates, confidence intervals and P values |
| `F1_A_forest_model_details.csv` | Base, bias-test, adjustment and final-result fields |
| `F1_A_forest_row_level_effects.xlsx` | Row-level effects and eligibility |
| `forest_details/` | Trait-specific model reports |
| `funnel/` | Bias summaries, source rows and trait-specific original/conditional trim-and-fill funnels |
| `F1_A_session_info.txt` | Input checksum, settings and software environment |

There is no `funnel_all.pdf`. The same output conventions apply to the other three sheets. `F1_Data_run_summary.csv` records worksheet status and errors. The larger F1_D analyses can take several minutes; wait for `DONE. Selected worksheets completed:` and check the summary.

All four sheets run by default. `sheets_to_run` can restrict a rerun, although all four valid input sheets must remain present. `output_root_override` can specify a shorter writable destination. Close generated workbooks and file-locking viewers before rerunning. Write-access checks and diagnostic files report path or file-lock problems. Repeated runs overwrite matching files without deleting obsolete output; use an empty output directory for the final release.
