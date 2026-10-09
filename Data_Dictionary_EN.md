# Data dictionary and analysis-input provenance

## 1. Scope and provenance

This document describes `data/F1_Data.xlsx` and `data/Table S2.xlsx`, the inputs to `F1.R`, `F2andF3.R`, and `F4andF5.R`. All source observations originate from the published-study records compiled in Table S1 (TS1). The two workbooks organize those records for different analyses; they are not separate experimental datasets.

For Figure 1, the records were organized first by comparison (invasive versus native plants; introduced-range versus native-range populations), then by physical or chemical defense. The relevant group means, standard deviations, sample sizes and StudyID were placed in four worksheets. Within each worksheet, records are provided once under their overall defense category and again under their specific trait. Each `Variable` is analyzed separately. The overall and trait blocks therefore must not be concatenated as independent observations in one model.

Table S2 contains the study/system identifiers, eligibility codes, nutrient and defense measurements, and source metadata required for the paired analyses. F2/F3 use its invasive-versus-native records; F4/F5 use its introduced-versus-native-range records. `NHshujuBG.xlsx` is the author's local filename for this analysis input. The distributed input is `Table S2.xlsx`; its filename and checksum in the run record identify the version used. Differences in workbook layout or descriptive source labels mean that these filenames should not be interpreted as a claim of byte-identical files.

The repository provides the exact analysis-ready inputs. The scripts calculate effects, perform pairing and filtering, fit models, and export results from these files. They do not regenerate these workbooks from Table S1. StudyID and, in Table S2, study titles and figure/table references support source lookup; neither workbook contains a complete row-by-row Table S1 crosswalk.

## 2. Workbook structure and group directions

| File/sheet | Comparison and contents | Data rows | Studies |
| --- | --- | ---: | ---: |
| F1_Data.xlsx / F1_A | Invasive versus native plants; physical defense | 504 | 40 |
| F1_Data.xlsx / F1_B | Invasive versus native plants; chemical defense | 896 | 37 |
| F1_Data.xlsx / F1_C | Introduced versus native range; physical defense | 670 | 39 |
| F1_Data.xlsx / F1_D | Introduced versus native range; chemical defense | 1670 | 71 |
| Table S2.xlsx / Table S2 | Both comparisons; nutrients and defense traits | 768 | 31 |

F1 headers are in row 1. Table S2 has its title in row 1, headers in row 2, and observations in rows 3–770. It has 24 physical columns: 23 named columns and an unnamed column B containing study titles. Table S2 contains 311 `invasive_vs_native` rows and 457 `introduced_vs_native_range` rows across 180 SystemIDs. These are input-record counts; they are not regression sample sizes or numbers of independent studies. The study counts across files/sheets overlap and must not be added.

| Comparison | Focal group, suffix i | Reference group, suffix c | Positive effect/difference |
| --- | --- | --- | --- |
| invasive_vs_native | Invasive plants | Native comparator plants | Higher value in invasive plants |
| introduced_vs_native_range | Introduced-range populations | Native-range populations of the focal species | Higher value in introduced-range populations |

## 3. F1 fields and trait groups

| Field | Type | Definition |
| --- | --- | --- |
| StudyID | Text identifier | Source study; repeated IDs denote observations from the same study |
| Variable | Text category | Overall defense category or specific trait analyzed in a separate model |
| Mc | Numeric | Reference-group mean |
| SDc | Numeric | Reference-group standard deviation |
| Nc | Numeric count | Reference-group sample size underlying the mean |
| Mi | Numeric | Focal-group mean |
| Sdi | Numeric | Focal-group standard deviation; equivalent to `SDi` in imported analysis objects |
| Ni | Numeric count | Focal-group sample size underlying the mean |

SDc and Sdi are standard deviations, not standard errors. Nc and Ni are sample sizes, not study counts. Means and SDs in one row must refer to the same measurement unit and basis. F1 has no explicit unit, species, SystemID or Table S1 row-ID field; the original source records provide those details. All eight fields are populated in the distributed workbook. Its group means and sample sizes are positive and SDs nonnegative.

| Sheet | Overall-category rows | Specific-trait rows, with labels used in the workbook |
| --- | ---: | --- |
| F1_A | Physical Defense: 252 | Cellulose 42; Trichome 6; Leaf Thickness 25; Leaf Toughness 57; Resin 16; LMA 35; Lignin 71 |
| F1_B | Chemical Defense: 448 | Terpenoids 37; Flavonoids 89; Phenolic Compounds 216; Tannins 82; Alkaloids 14; Glycosides 10 |
| F1_C | Physical Defense: 335 | Cellulose 86; Hemicellulose 33; Trichome 87; Leaf Thickness 16; Leaf Toughness 54; LMA 45; Lignin 14 |
| F1_D | Chemical Defense: 835 | Terpenoids 132; Flavonoids 101; Phenolic Compounds 190; Tannins 85; Alkaloids 199; Proteinaceous Toxins 14; Glucosinolates 46; Glycosides 68 |

LMA means leaf mass per area. `Trichome` denotes the workbook's trichome trait. Leaf Thickness labels contain a leading nonbreaking space in the input; the table above omits that invisible prefix for readability. The distributed workbook preserves the original labels. The trait and overall blocks contain the same underlying study/numerical records in different groupings.

## 4. Table S2 fields and coding

The six measurement fields `Mc`, `SDc`, `Nc`, `Mi`, `Sdi`, and `Ni` have the definitions above, with group directions set by ComparisonType. The remaining fields are:

| Input field | Definition and observed coding |
| --- | --- |
| Study ID | Source-study identifier; normalized to `StudyID` in R |
| Column B, no header | Source-publication title. This document calls it the study-title column; it is descriptive metadata and not a model variable |
| System ID | Identifier for the matched biological comparison within a study; normalized to `SystemID`. Nutrient and defense records with the relevant shared identifiers are paired |
| ComparisonType | `invasive_vs_native` (311); `introduced_vs_native_range` (457) |
| InvasiveSpecies | Scientific name of the focal invasive species, including for the range comparison |
| ComparatorSpecies | Native comparator species, or `native-range populations` for the within-species range comparison |
| Trait | Measured nutrient or defense trait; observed labels are listed below |
| TraitCategory | Author's category code: `nutrient` (342), `physical_defense` (158), `chemical_defense` (268). The scripts assign Lignin to physical defense in memory |
| Treatment | Source-specific treatment/condition label; free text, not a universal numerical treatment scale |
| TraitForm | Author's trait-form code: `material` (703), `non_material` (65). These are record-level codes, not units or a unique mapping from trait name. F3 normalizes specified nutrient, structural-material and chemical-trait names to `material` in memory |
| Tissue | Recorded tissue code: `leaf` (767) and one literal `physical_defense`. Both paired-analysis scripts retain only `leaf`; the latter record does not pass that filter |
| ExperimentType | `Field` (318), `Greenhouse` (410), `Common_garden` (28), `Laboratory` (12) |
| PairDesign | Author's design code: `co_occurring_pair` (50), `related_pair` (261), `home_away_population` (457). The label itself does not encode a numerical measure of relatedness |
| DataLevel | Observation-level code: `species_pair` (311), `population_pair` (457). The scripts also accept `treatment_level`, which is absent from this input |
| PairingLevel | Author's pairing code: `exact` (741), `baseline` (27). Both are retained. The encoded SystemID determines the actual pairing; this file does not provide a separate operational decision rule for assigning these two labels |
| UseForRegression | Inclusion-designation field; all 768 records contain `main`. Other numerical, tissue and pairing requirements still apply |
| Unit | Source measurement-unit label; observed labels and interpretation are listed below |
| FigureTable | Source figure/table locator as recorded. Free-text references assist source lookup; they are not DOIs or model predictors |

Field counts describe the supplied workbook before filtering or in-memory reclassification. Text case, whitespace, column-name aliases and recognized trait aliases are normalized in R. Species and treatment strings are retained as metadata rather than converted to analytical grouping factors automatically.

Observed Trait labels are Plant carbon (160), Plant nitrogen (161), Plant phosphorus (21), Leaf toughness (30), Leaf thickness (4), Leaf mass per area (39), Trichomes (19), Cellulose (43), Lignin (24), Phenolics (68), Tannins (23), Flavonoids (35), Terpenoids (65), Alkaloids (30), and Glycosides (46). The scripts recognize Plant carbon/nitrogen/phosphorus as C/N/P and Leaf mass per area as LMA.

| Unit label | Rows | Meaning and interpretation |
| --- | ---: | --- |
| mg/g | 676 | Milligrams per gram. The label alone does not specify fresh- or dry-mass basis |
| g/m² | 39 | Grams per square metre |
| mm | 3 | Millimetres |
| N | 31 | Conventionally newtons; the literal input labels are retained, including records where the trait–unit pairing requires source-specific interpretation |
| n | 19 | Literal lower-case n|

A recorded label should not be used to infer a uniform unit for every row of a trait. The scripts do not correct or convert measurement units automatically. Within each row, the two groups and their SDs must be comparable; cross-record unit interpretation should follow the original study. All Table S2 measurement and metadata cells are populated; the unnamed header in column B is a structural feature, not a missing study title. Missing/unestimable derived results are exported as `NA`, rather than zero.

## 5. Derived variables and analytical use

| Derived quantity | Definition and use |
| --- | --- |
| lnRR | `log(Mi/Mc)`, requiring positive group means. Dimensionless effect used in F1, F2/F3, and F5 companion forests/funnels |
| Sampling variance of lnRR | Based on `Sdi²/(Ni Mi²) + SDc²/(Nc Mc²)`. F1 and F2/F3 retain their documented small denominator stabilizers; F1 additionally retains its explicitly flagged 1/N fallback |
| d / raw mean difference | `Mi − Mc`; variance `Sdi²/Ni + SDc²/Nc`, subject to the script's numerical stabilizers where applicable |
| Cha, M2 | For F4/F5 regression, signed-square-root transformed differences, divided by positive RMS scales without subtracting a mean. Nutrient scale uses unique paired study/system combinations; defense scale is trait-specific within the paired data. This is not a z score |
| zX, zY | F2/F3: predictor/response lnRR. F4/F5 M2: scaled signed-square-root differences. Their meaning depends on the program |
| RAW | Eligible, valid paired records before additional IQR removal; it does not mean every spreadsheet row is included |
| IQR_removed | Records within the inclusive 1.5-IQR bounds on both axes. F2B, F2D and F3C use this treatment for the displayed regression; F4/F5 use it only for numerical sensitivity analyses |
| Weighted R² | One minus the ratio of weighted residual to weighted total sums of squares, following each script's implementation. Describes the fitted relationship; not a P value |
| Forest P | Test of the pooled lnRR against zero. This determines stars in the left F3/F5 forests |
| Regression P | Coefficient test from the specified model. F3/F5 right-hand panels display R²; coefficient P values are reported in tables |
| Egger/Begg P, k0 | Funnel asymmetry-test P values and estimated missing-effect count from trim-and-fill. They do not have the meaning of forest or regression P values |

F3 forests and funnels use original trait lnRR records underlying the pairs retained for each panel. F3C retains 58 of 62 eligible paired rows after IQR filtering; its forests/funnels use the source records underlying those 58 rows, with no second IQR exclusion. F5 forests/funnels similarly use original source observations underlying each M2-RAW paired relationship. A nutrient repeated across panels can therefore have different source subsets and pooled estimates.

For the current input, the F5F nutrient (N) funnel's Egger and trim-and-fill calculations do not converge and are reported as `NA`. These entries must not be interpreted as zero, absence of asymmetry, or a zero missing-effect count. They describe unavailable funnel diagnostics, separately from the fitted regression and companion forest.

## 6. Record traceability and release identification

F1 exports the input-row index `.rowid` within each worksheet. F2/F3 and F4/F5 create RowID in original data-row order, after the detected header. For the distributed Table S2, the Excel row is `RowID + 2`. Paired-data outputs retain nutrient-source row lists and defense-source row identifiers (including `X_RowIDs` and `Y_RowID` where applicable); forest/funnel workbooks retain their original source rows.

Row indices refer to this workbook version and are not permanent identifiers across reordering or later revisions. StudyID, SystemID, trait, treatment and the source locator provide additional context. Input checksums, panel manifests, completion records and package versions identify each run. The archived version-specific DOI should be cited together with the final release version once publication and licensing are complete.
