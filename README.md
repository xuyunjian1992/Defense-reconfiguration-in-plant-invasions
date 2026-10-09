# Data and code for the plant defense meta-analysis

This repository contains the analysis inputs and three R programs for Figures 1–5 and the associated funnel plots. The programs run independently; no earlier script, external helper file or RStudio project file is required. The scientific methods and output definitions are documented separately for each program.

## Files and figure correspondence

| Program | Input | Analyses and figures | Detailed instructions |
| --- | --- | --- | --- |
| `F1.R` | `data/F1_Data.xlsx` | Four Figure 1 forest plots and their trait-specific funnels | [F1 instructions](docs/README_F1.md) |
| `F2andF3.R` | `data/Table S2.xlsx` | Invasive versus native plants: Figures 2 and 3, companion forests and supplementary Figure S5 | [F2/F3 instructions](docs/README_F2andF3.md) |
| `F4andF5.R` | `data/Table S2.xlsx` | Introduced-range versus native-range populations: Figures 4 and 5, companion forests and supplementary Figure S6 | [F4/F5 instructions](docs/README_F4andF5.md) |

Keep the three `.R` files at the repository root, the two workbooks in `data/`, and the documentation in `docs/`. The space in `Table S2.xlsx` is part of its filename. Extract the downloaded archive before running the programs.

## Run the analyses

Install R and the required packages. A single installation command covering all three programs is:

```r
install.packages(c(
  "readxl", "dplyr", "purrr", "metafor", "writexl", "ggplot2",
  "scales", "stringr", "tibble", "patchwork", "tidyr", "clubSandwich"
))
```

Package installation requires a configured software repository or an offline package source. `F1.R` also attempts to install missing dependencies from CRAN; the other two scripts check availability and stop if packages are missing. Once dependencies are available, the analysis scripts read and write local files. They contain no data-upload operation. F1 PDF export requires Cairo graphics support.

Run each program in a separate R process from the repository root:

```sh
Rscript F1.R
Rscript F2andF3.R
Rscript F4andF5.R
```

The order is optional. In R or RStudio, the complete script can instead be sourced by its full path, for example `source("/path/to/repository/F2andF3.R", encoding = "UTF-8")`. Sourcing the complete file allows the script to resolve paths from its location. Running isolated lines requires the correct working directory or the documented path override.

Outputs are created under `results/F1_Data/`, `F2_F3_results/`, and `F4_F5_results/`, respectively. F1 records each worksheet's status in `F1_Data_run_summary.csv`; the other programs record successful completion as `COMPLETE` in `run_info.txt`. Use a fresh output directory for a release and review completion records before using the output set. Reruns replace matching filenames; obsolete files from earlier versions are not automatically removed. Runtime logs record input checksums and the actual R/package environment.

## Data provenance and interpretation

The source observations were compiled in Table S1 (TS1) from the published studies included in the meta-analysis. `F1_Data.xlsx` reorganizes those observations into four worksheets by comparison type and defense category. `Table S2.xlsx` is the analysis-ready input for the paired nutrient–defense analyses. These exact workbooks are supplied so that reproduction starts from the documented analysis inputs. The R programs do not reconstruct either workbook from Table S1.

Source Table S1 location: `{{TABLE_S1_LOCATION}}`. Replace this placeholder with the stable location of the source supplementary dataset before release.

`NHshujuBG.xlsx` was the author's local filename for the Table S2 analysis input. The public programs read `Table S2.xlsx`; a second workbook with the local filename is unnecessary. Workbook presentation and descriptive metadata can differ between the local and supplementary versions, so the distributed `data/Table S2.xlsx` and its recorded checksum identify the release input.

Definitions, group directions, units, coding values and derived quantities are given in the [English data dictionary](docs/Data_Dictionary_EN.md) and [Chinese data dictionary](docs/Data_Dictionary_ZH.md). Instructions in Chinese are available in [README_zh.md](README_zh.md). Numerical sensitivity analyses are retained in the result tables; exported main figures follow the panel definitions documented for each program.

## License and archived citation

Release version: `{{RELEASE_VERSION}}`.

Code license: `{{CODE_LICENSE}}`. Data license: `{{DATA_LICENSE}}`. The authors must replace these placeholders with the chosen licenses and include their applicable license texts before publication.

Archived release DOI: `{{VERSION_DOI}}`. Replace this placeholder with the DOI of the published, version-specific code-and-data deposit and provide its complete citation. The DOI should identify the same inputs and code used for the reported results. A DOI reservation or an unversioned repository URL does not document a completed archive.

## Environment and file integrity

[Software environments](docs/Software_Environment.md) records the observed Windows and Linux versions and the scope of the reproduction checks. [Input version notes](docs/Data_Revision_Notes.md) identifies the distributed workbooks. Verify files against [FILE_CHECKSUMS_SHA256.txt](FILE_CHECKSUMS_SHA256.txt) when transferring the release.
