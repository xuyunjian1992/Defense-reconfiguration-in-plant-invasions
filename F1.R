############################################################
# F1 four-worksheet batch: forest plots and matching individual funnel plots
# Updated 2026-10-08: protected file output and original-error preservation
#
# Required columns in each worksheet:
# StudyID, Variable, Mc, SDc, Nc, Mi, Sdi, Ni
# yi = log(Mi / Mc)
# vi = Sdi^2 / (Ni * (Mi^2 + 1e-12)) + SDc^2 / (Nc * (Mc^2 + 1e-12))
# The supplied fallback rule and all statistical settings are retained.
#
# Only the forest based on the original final_* values is exported.
# Its new name is <sheet_name>_forest_primary.pdf/png.
# This naming change does not change the model or remove trim-and-fill calculations.
# Forest and funnel modules preserve their own original bias-test settings.
# One F1_Data.xlsx workbook contains F1_A, F1_B, F1_C and F1_D.
# Each worksheet is analyzed separately in a single batch run.
############################################################

# 1) INPUT AND OUTPUT PATHS — portable project version ----
# Open F1_project.Rproj and source this file, or run: Rscript F1.R
# Put the complete F1_Data.xlsx workbook in data/. Run once for all four panels.
# Absolute paths from the author's computer are not required.
find_project_root <- function() {
  source_files <- vapply(sys.frames(), function(frame) {
    if (!is.null(frame$ofile)) as.character(frame$ofile)[1L] else ""
  }, character(1))
  source_files <- source_files[nzchar(source_files)]
  if (length(source_files) > 0L) {
    return(dirname(normalizePath(tail(source_files, 1L), mustWork = TRUE)))
  }
  script_arg <- grep("^--file=", commandArgs(trailingOnly = FALSE), value = TRUE)
  if (length(script_arg) > 0L) {
    script_path <- sub("^--file=", "", script_arg[1L])
    # Rscript can encode spaces as ~+~ in its internal --file argument.
    script_path <- gsub("~+~", " ", script_path, fixed = TRUE)
    return(dirname(normalizePath(script_path, mustWork = TRUE)))
  }
  # Running the whole script with RStudio's Run button uses the project working directory.
  normalizePath(getwd(), mustWork = TRUE)
}
project_dir <- find_project_root()
file_name <- "F1_Data.xlsx"
data_dir <- file.path(project_dir, "data")
infile <- file.path(data_dir, file_name)
workbook_prefix <- tools::file_path_sans_ext(basename(file_name))
output_root <- file.path(project_dir, "results", workbook_prefix)

# Optional short output path, e.g. "D:/F1_results" on Windows.
# NULL retains this version's default output folder.
output_root_override <- NULL
if (!is.null(output_root_override)) output_root <- path.expand(output_root_override)

# Default: run all panels. To rerun only the failed panel, use "F1_D".
sheets_to_run <- c("F1_A", "F1_B", "F1_C", "F1_D")

# Read the four named sheets in a fixed order, independently of workbook tab order.
sheet_names <- c("F1_A", "F1_B", "F1_C", "F1_D")
# NULL automatically identifies the overall-defense title within each worksheet.
# Optional example: group_titles$F1_A <- "Physical Defense"
group_titles <- setNames(vector("list", length(sheet_names)), sheet_names)
if (length(sheets_to_run) == 0L || anyNA(sheets_to_run) ||
    anyDuplicated(sheets_to_run) || !all(sheets_to_run %in% sheet_names)) {
  stop("sheets_to_run must contain unique names selected from F1_A, F1_B, F1_C and F1_D.")
}
sheets_to_run <- sheet_names[sheet_names %in% sheets_to_run]

if (!file.exists(infile)) {
  stop("Input Excel file not found: ", infile,
       "\nSet file_name/data_dir in section 1, or place the file in the project's data folder.")
}
if (tolower(tools::file_ext(infile)) != "xlsx") stop("The input must be an .xlsx file.")

# 2) PACKAGES ----
pkgs <- c("readxl", "dplyr", "purrr", "metafor", "writexl", "ggplot2", "scales")
for (pkg in pkgs) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org")
  }
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}

# File output helpers. Statistical calculations are not changed.
ensure_output_dir <- function(path) {
  if (!dir.exists(path)) suppressWarnings(dir.create(path, recursive = TRUE, showWarnings = FALSE))
  if (!dir.exists(path)) {
    stop("Cannot create output directory: ", path,
         "\nCheck the folder permissions and path length, or set output_root_override to a short writable path.",
         call. = FALSE)
  }
  invisible(path)
}

with_output_file <- function(path, writer) {
  ensure_output_dir(dirname(path))
  path <- file.path(normalizePath(dirname(path), winslash = "/", mustWork = TRUE), basename(path))
  if (dir.exists(path)) stop("Output file path is a directory: ", path, call. = FALSE)
  extension <- tools::file_ext(path)
  temp_path <- tempfile(pattern = ".f1_", tmpdir = dirname(path),
                        fileext = if (nzchar(extension)) paste0(".", extension) else "")
  on.exit(unlink(temp_path), add = TRUE)
  io_warnings <- character()
  tryCatch(withCallingHandlers({
    writer(temp_path)
    size <- file.info(temp_path)$size
    if (is.na(size) || size <= 0) stop("The temporary export is empty or missing.")
    if (tolower(tools::file_ext(path)) == "png") {
      con <- file(temp_path, open = "rb")
      tail_bytes <- tryCatch({
        seek(con, where = max(0, size - 12), origin = "start")
        readBin(con, what = "raw", n = 12L)
      }, finally = close(con))
      expected_end <- as.raw(c(0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130))
      if (!identical(tail_bytes, expected_end)) stop("The PNG export is incomplete (missing IEND).")
    }
    if (!file.copy(temp_path, path, overwrite = TRUE, copy.mode = FALSE)) {
      stop("Cannot replace the destination file. It may be open in another program or not writable.")
    }
    if (is.na(file.info(path)$size) || file.info(path)$size != size) {
      stop("The destination file size does not match the completed export.")
    }
    invisible(path)
  }, warning = function(w) {
    io_warnings <<- c(io_warnings, conditionMessage(w))
    # Preserve warnings; do not suppress statistical or graphics diagnostics.
  }), error = function(e) {
    stop("Cannot write output file:\n", path,
         "\nOriginal error: ", conditionMessage(e),
         if (length(io_warnings)) paste0("\nRelated warning: ", tail(io_warnings, 1L)) else "",
         "\nClose this file in Excel or other viewers; check write permission and path length.",
         call. = FALSE)
  })
}

safe_write_csv <- function(x, file, ...) {
  with_output_file(file, function(path) utils::write.csv(x, file = path, ...))
}
safe_write_xlsx <- function(x, path, ...) {
  with_output_file(path, function(temp) writexl::write_xlsx(x, path = temp, ...))
}
safe_write_lines <- function(text, con, ...) {
  with_output_file(con, function(path) base::writeLines(text, con = path, ...))
}
safe_text_report <- function(path, expr) {
  code <- substitute(expr)
  env <- parent.frame()
  lines <- capture.output(eval(code, envir = env))
  safe_write_lines(lines, path)
}
safe_ggsave <- function(filename, ...) {
  with_output_file(filename, function(path) ggplot2::ggsave(filename = path, ...))
}
safe_plot_device <- function(path, open_device, draw) {
  with_output_file(path, function(temp) {
    open_device(temp)
    device_id <- dev.cur()
    on.exit({
      if (device_id %in% dev.list()) dev.off(device_id)
    }, add = TRUE)
    draw()
    dev.off(device_id)
  })
}

# Exercise a path as long as the longest planned filename before fitting models.
probe_output_dir <- function(path, filename_length) {
  ensure_output_dir(path)
  probe <- tempfile(pattern = ".f1check_", tmpdir = path)
  probe <- paste0(probe, strrep("x", max(0, filename_length - nchar(basename(probe)))))
  on.exit(unlink(probe), add = TRUE)
  safe_write_lines("F1 output write check", probe)
  invisible(TRUE)
}

check_existing_output <- function(path) {
  if (dir.exists(path)) stop("An expected output filename is a directory: ", path, call. = FALSE)
  if (!file.exists(path)) return(invisible(TRUE))
  # Open in append mode without writing any bytes, then close immediately.
  # This detects many Windows sharing locks before a long model fit.
  tryCatch({
    con <- file(path, open = "ab")
    close(con)
    invisible(TRUE)
  }, error = function(e) {
    stop("Existing output cannot be opened for writing:\n", path,
         "\nOriginal error: ", conditionMessage(e),
         "\nClose the file in Excel or other viewers, or choose a different output_root_override.",
         call. = FALSE)
  })
}

preflight_output_paths <- function(output_root, input_sheets, sheets_to_run) {
  probe_output_dir(output_root, 55L)
  safe_trait <- function(x) {
    x <- gsub("[/\\\\:*?\"<>|]+", "_", as.character(x))
    x <- gsub("\\s+", "_", x)
    gsub("__+", "_", x)
  }
  for (sheet in sheets_to_run) {
    base <- file.path(output_root, sheet)
    probe_output_dir(base, nchar(paste0(sheet, "_forest_primary_statistics.xlsx")))
    root_suffixes <- c("forest_primary.pdf", "forest_primary.png", "forest_results.xlsx",
      "forest_primary_results.csv", "forest_primary_statistics.xlsx", "forest_primary_statistics.csv",
      "forest_model_details.csv", "forest_row_level_effects.xlsx", "session_info.txt")
    for (suffix in root_suffixes) check_existing_output(file.path(base, paste0(sheet, "_", suffix)))
    funnel_base <- file.path(base, "funnel")
    for (suffix in c("funnel_publication_bias_summary.xlsx", "funnel_publication_bias_summary.csv",
                     "funnel_row_level_effects.xlsx")) {
      check_existing_output(file.path(funnel_base, paste0(sheet, "_", suffix)))
    }
    traits <- unique(c(as.character(input_sheets[[sheet]]$Variable),
                       trimws(as.character(input_sheets[[sheet]]$Variable))))
    traits <- traits[!is.na(traits) & nzchar(traits)]
    for (trait in traits) {
      key <- safe_trait(trait)
      probe_output_dir(file.path(base, "forest_details", paste0("Variable_", key)),
                       nchar(paste0(sheet, "_forest_", key, "_model_details.txt")))
      probe_output_dir(file.path(base, "funnel", paste0("Variable_", key)),
                       nchar(paste0(sheet, "_funnel_", key, "_trimfill_details.txt")))
      check_existing_output(file.path(base, "forest_details", paste0("Variable_", key),
                            paste0(sheet, "_forest_", key, "_model_details.txt")))
      trait_funnel <- file.path(funnel_base, paste0("Variable_", key))
      for (suffix in c(paste0("funnel_original_", key, c(".pdf", ".png")),
                       paste0("funnel_trimfill_", key, c(".pdf", ".png")),
                       paste0("funnel_", key, c("_model_tests.txt", "_trimfill_details.txt")))) {
        check_existing_output(file.path(trait_funnel, paste0(sheet, "_", suffix)))
      }
    }
  }
  invisible(TRUE)
}

# 3) ANALYSIS MODULES ----
# Keep separate function environments: model objects and options must not overwrite each other.
run_forest <- function(input_data, out_root, dataset_prefix, group_title = NULL) {
  out_file <- function(suffix, directory = out_root) {
    file.path(directory, paste0(dataset_prefix, "_", suffix))
  }
  ensure_output_dir(out_root)
  # Forest x-axis label. The panel title is supplied by the common entry point.
  x_axis_label <- "Effect Size"

  # Overall rows forced to the top if present in Variable.
  overall_names <- c("Physical Defense", "Chemical Defense")

  # Model switch.
  # FALSE = row-level random-effects model rma().
  # TRUE  = try rma.mv(..., random = ~1|StudyID) for the primary forest estimates only.
  #         Note: trim-and-fill itself is still based on the standard rma object.
  use_study_random_effect <- TRUE

  # Variance handling.
  use_fallback_1overN <- TRUE

  # Publication-bias / trim-and-fill options.
  alpha_egger <- 0.05
  use_begg_ranktest <- TRUE
  alpha_begg <- 0.05
  run_trimfill_if_bias <- TRUE

  # This reproduces the old script's behavior: if default trimfill gives k0=0,
  # try side="left" and side="right" and keep the option with larger k0.
  # For a more conservative sensitivity analysis, set FALSE.
  force_try_both_sides_if_k0_0 <- TRUE

  # Figure display options.
  show_numeric_effect_text <- FALSE  # reviewer suggested removing estimate/CI text from Fig. 1
  show_significance_stars <- TRUE
  save_png <- TRUE

  # Preserve the supplied forest appearance. Adjustment status is recorded in the result tables.
  mark_trimfill_adjusted_points <- FALSE

  # Optional fixed x-axis limits for the actual forest data region.
  # Use NULL to calculate automatically from BOTH primary and adjusted CIs.
  # Example: fixed_xlim <- c(-2, 2)
  fixed_xlim <- NULL

  # Multi-panel alignment options.
  # To align Fig. 1 panels with different trait counts, use the same target row number,
  # same width, and same height for all subplots. For your current Fig. 1, 9 is useful
  # because the largest panel has 9 displayed rows.
  pad_to_n_rows <- 9
  height_mode <- "fixed"       # "fixed" or "by_rows"
  fixed_plot_height <- 4.15
  row_height <- 0.48
  base_height <- 1.10
  plot_width <- 6.55

  # Layout tuning.
  # The x-axis is manually drawn only across the effect-size region.
  # Trait labels are right-aligned immediately to the left of that region, so the
  # zero line stays visually centered and the gap between labels and CIs is reduced.
  label_column_width_ratio <- 0.32
  right_star_width_ratio   <- 0.10
  min_label_column_width   <- 0.80
  min_right_star_width     <- 0.22
  label_to_axis_gap_ratio  <- 0.035
  star_to_axis_gap_ratio   <- 0.080
  label_size <- 4.35
  star_size  <- 5.75
  point_size <- 2.72   # slightly smaller than v7 to reduce crowding around short CIs
  ci_line_width <- 0.82
  ci_cap_height <- 0.155

  # Padded blank rows. For clean multi-panel alignment, using the same pale-yellow fill
  # avoids a visually abrupt white block beneath short panels.
  pad_blank_fill <- "specific"   # "specific" or "white"

  # Forest colors.
  col_bg_overall  <- "#ECEAF8"  # soft lavender
  col_bg_specific <- "#FFF4D6"  # soft pale yellow
  col_sig_text    <- "#1F5FBF"  # blue for significant labels
  col_nonsig_text <- "black"
  col_star        <- "#D7191C"  # red
  col_point       <- "black"
  col_zero        <- "grey40"

  # 2) helper functions ----
  safe_name <- function(x) {
    x <- as.character(x)
    x <- gsub("[/\\\\:*?\"<>|]+", "_", x)
    x <- gsub("\\s+", "_", x)
    x <- gsub("__+", "_", x)
    x
  }

  fmt_p <- function(p) {
    if (is.na(p) || !is.finite(p)) return("NA")
    if (p < 1e-4) return(formatC(p, format = "e", digits = 2))
    formatC(p, format = "f", digits = 4)
  }

  stars <- function(p) {
    if (is.na(p) || !is.finite(p)) return("")
    if (p < 0.001) return("***")
    if (p < 0.01)  return("**")
    if (p < 0.05)  return("*")
    ""
  }

  clean_key <- function(x) {
    tolower(gsub("[^A-Za-z0-9]+", "", x))
  }

  standardize_meta_columns <- function(dat) {
    aliases <- list(
      StudyID  = c("studyid", "study", "id", "number", "refid", "referenceid"),
      Variable = c("variable", "trait", "traits", "index", "category"),
      Mc       = c("mc", "meanc", "meancontrol", "controlmean", "native_mean", "nativemean", "home_mean", "homemean"),
      SDc      = c("sdc", "sdcontrol", "controlsd", "native_sd", "nativesd", "home_sd", "homesd"),
      Nc       = c("nc", "ncontrol", "controln", "native_n", "nativen", "home_n", "homen"),
      Mi       = c("mi", "meani", "meaninvasive", "invasivemean", "introduced_mean", "introducedmean", "away_mean", "awaymean"),
      Sdi      = c("sdi", "sdinvasive", "invasivesd", "introduced_sd", "introducedsd", "away_sd", "awaysd"),
      Ni       = c("ni", "ninvasive", "invasiven", "introduced_n", "introducedn", "away_n", "awayn")
    )
    
    original_names <- names(dat)
    keys <- clean_key(original_names)
    
    for (canonical in names(aliases)) {
      hit <- which(keys %in% aliases[[canonical]])
      if (length(hit) == 0) {
        stop(
          "Missing required column for ", canonical, ".\n",
          "Current column names are: ", paste(original_names, collapse = ", ")
        )
      }
      names(dat)[hit[1]] <- canonical
    }
    dat
  }

  eps <- 1e-12
  compute_row <- function(Mc, SDc, Nc, Mi, Sdi, Ni) {
    yi <- NA_real_
    vi <- NA_real_
    method <- NA_character_
    
    if (!is.na(Mc) && !is.na(Mi) && Mc > 0 && Mi > 0) {
      yi <- log(Mi / Mc)
      
      if (!is.na(SDc) && !is.na(Sdi) && is.finite(SDc) && is.finite(Sdi) &&
          !is.na(Nc) && !is.na(Ni) && Nc > 0 && Ni > 0) {
        vi <- (Sdi^2) / (Ni * (Mi^2 + eps)) + (SDc^2) / (Nc * (Mc^2 + eps))
        method <- "SD_formula"
        if (!is.finite(vi) || vi <= 0) vi <- NA_real_
      }
      
      if (is.na(vi) && isTRUE(use_fallback_1overN) &&
          !is.na(Nc) && !is.na(Ni) && Nc > 0 && Ni > 0) {
        vi <- 1 / Nc + 1 / Ni
        method <- "1/N_fallback"
        if (!is.finite(vi) || vi <= 0) vi <- NA_real_
      }
    } else {
      method <- "invalid_mean"
    }
    
    list(
      yi = yi,
      vi = vi,
      sei = ifelse(is.na(vi), NA_real_, sqrt(vi)),
      method = method
    )
  }

  fit_primary_model <- function(dv) {
    if (isTRUE(use_study_random_effect) &&
        length(unique(dv$StudyID)) > 1 &&
        any(duplicated(dv$StudyID))) {
      
      fit_mv <- tryCatch(
        metafor::rma.mv(yi = yi, V = vi, random = ~ 1 | StudyID,
                        data = dv, method = "REML"),
        error = function(e) NULL
      )
      
      if (!is.null(fit_mv)) {
        return(list(fit = fit_mv, model_used = "rma.mv_REML_random_StudyID"))
      }
    }
    
    fit_uni <- metafor::rma(yi = dv$yi, vi = dv$vi, method = "REML")
    list(fit = fit_uni, model_used = "rma_REML")
  }

  extract_summary <- function(fit_obj) {
    fit <- fit_obj$fit
    data.frame(
      estimate = as.numeric(fit$b[1]),
      ci.lb = as.numeric(fit$ci.lb[1]),
      ci.ub = as.numeric(fit$ci.ub[1]),
      pval = as.numeric(fit$pval[1]),
      model_used = fit_obj$model_used,
      stringsAsFactors = FALSE
    )
  }

  # 3) read data & compute row-level yi/vi ----
  raw <- input_data
  raw <- standardize_meta_columns(raw)
  if (is.null(group_title)) {
    matched_overall <- overall_names[overall_names %in% trimws(as.character(raw$Variable))]
    group_title <- if (length(matched_overall) == 1L) matched_overall else dataset_prefix
  }

  df0 <- raw %>%
    mutate(
      StudyID  = as.character(StudyID),
      Variable = trimws(as.character(Variable)),
      Mc  = as.numeric(Mc),
      SDc = as.numeric(SDc),
      Nc  = as.numeric(Nc),
      Mi  = as.numeric(Mi),
      Sdi = as.numeric(Sdi),
      Ni  = as.numeric(Ni),
      .rowid = dplyr::row_number()
    )

  tmp <- purrr::pmap(
    df0 %>% dplyr::select(Mc, SDc, Nc, Mi, Sdi, Ni),
    compute_row
  )

  df_all <- df0 %>%
    mutate(
      yi = purrr::map_dbl(tmp, "yi"),
      vi = purrr::map_dbl(tmp, "vi"),
      sei = purrr::map_dbl(tmp, "sei"),
      Var_method = purrr::map_chr(tmp, "method")
    )

  df <- df_all %>%
    filter(
      !is.na(StudyID), !is.na(Variable),
      !is.na(yi), !is.na(vi),
      is.finite(yi), is.finite(vi), vi > 0
    )

  invalid_rows <- df_all %>%
    filter(
      is.na(StudyID) | is.na(Variable) |
        is.na(yi) | is.na(vi) |
        !is.finite(yi) | !is.finite(vi) | vi <= 0
    )

  if (nrow(df) == 0) stop("No valid effect sizes after filtering. Please check the input data.")

  safe_write_xlsx(
    list(valid_row_level_effects = df, invalid_or_excluded_rows = invalid_rows),
    out_file("forest_row_level_effects.xlsx")
  )

  # 4) run meta-analysis by trait + Egger/Begg + trimfill ----
  trait_order_in_file <- unique(df$Variable)
  summary_list <- list()

  for (v in trait_order_in_file) {
    dv <- df %>% filter(Variable == v)
    n_es <- nrow(dv)
    k_st <- length(unique(dv$StudyID))
    if (n_es < 1) next
    
    v_dir <- file.path(out_root, "forest_details", paste0("Variable_", safe_name(v)))
    ensure_output_dir(v_dir)
    
    primary_obj <- tryCatch(fit_primary_model(dv), error = function(e) NULL)
    if (is.null(primary_obj)) next
    primary_sm <- extract_summary(primary_obj)
    
    res_rma <- tryCatch(metafor::rma(yi = dv$yi, vi = dv$vi, method = "REML"),
                        error = function(e) NULL)
    
    egger_p <- begg_p <- NA_real_
    eg_obj <- rk_obj <- NULL
    trigger <- FALSE
    trf <- NULL
    trf_used <- FALSE
    k0_best <- NA_integer_
    trimfill_side_used <- NA_character_
    tf_est <- tf_lb <- tf_ub <- tf_p <- NA_real_
    egger_p_tf <- NA_real_
    
    if (!is.null(res_rma)) {
      eg_obj <- tryCatch(metafor::regtest(res_rma, model = "rma", predictor = "sei"),
                         error = function(e) NULL)
      if (!is.null(eg_obj)) egger_p <- tryCatch(as.numeric(eg_obj$pval), error = function(e) NA_real_)
      
      rk_obj <- tryCatch(metafor::ranktest(res_rma), error = function(e) NULL)
      if (!is.null(rk_obj)) begg_p <- tryCatch(as.numeric(rk_obj$pval), error = function(e) NA_real_)
      
      if (is.finite(egger_p) && !is.na(egger_p) && egger_p < alpha_egger) trigger <- TRUE
      if (isTRUE(use_begg_ranktest) && is.finite(begg_p) && !is.na(begg_p) && begg_p < alpha_begg) trigger <- TRUE
      
      if (isTRUE(run_trimfill_if_bias) && isTRUE(trigger)) {
        trf0 <- tryCatch(metafor::trimfill(res_rma), error = function(e) NULL)
        trimfill_side_used <- "default"
        
        if (isTRUE(force_try_both_sides_if_k0_0) && !is.null(trf0)) {
          k0_0 <- tryCatch(as.integer(trf0$k0), error = function(e) NA_integer_)
          if (!is.na(k0_0) && k0_0 == 0) {
            trfL <- tryCatch(metafor::trimfill(res_rma, side = "left"),  error = function(e) NULL)
            trfR <- tryCatch(metafor::trimfill(res_rma, side = "right"), error = function(e) NULL)
            k0L <- if (!is.null(trfL)) tryCatch(as.integer(trfL$k0), error = function(e) NA_integer_) else NA_integer_
            k0R <- if (!is.null(trfR)) tryCatch(as.integer(trfR$k0), error = function(e) NA_integer_) else NA_integer_
            candidates <- list(default = trf0, left = trfL, right = trfR)
            k0s <- c(default = k0_0, left = k0L, right = k0R)
            best_name <- names(which.max(ifelse(is.na(k0s), -Inf, k0s)))[1]
            trf <- candidates[[best_name]]
            trimfill_side_used <- best_name
          } else {
            trf <- trf0
          }
        } else {
          trf <- trf0
        }
        
        if (!is.null(trf)) {
          trf_used <- TRUE
          k0_best <- tryCatch(as.integer(trf$k0), error = function(e) NA_integer_)
          tf_est <- tryCatch(as.numeric(trf$b[1]), error = function(e) NA_real_)
          tf_lb  <- tryCatch(as.numeric(trf$ci.lb[1]), error = function(e) NA_real_)
          tf_ub  <- tryCatch(as.numeric(trf$ci.ub[1]), error = function(e) NA_real_)
          tf_p   <- tryCatch(as.numeric(trf$pval[1]), error = function(e) NA_real_)
          
          res_tf <- tryCatch(metafor::rma(yi = trf$yi, vi = trf$vi, method = "REML"),
                             error = function(e) NULL)
          if (!is.null(res_tf)) {
            eg_tf_obj <- tryCatch(metafor::regtest(res_tf, model = "rma", predictor = "sei"),
                                  error = function(e) NULL)
            if (!is.null(eg_tf_obj)) egger_p_tf <- tryCatch(as.numeric(eg_tf_obj$pval), error = function(e) NA_real_)
          }
        }
      }
      
      safe_text_report(out_file(paste0("forest_", safe_name(v), "_model_details.txt"), v_dir), {
        cat("Variable:", v, "\n")
        cat("n(effect sizes) =", n_es, "; k(studies) =", k_st, "\n\n")
        cat("Primary forest model:\n")
        print(primary_obj$fit)
        cat("\nStandard rma model used for Egger/Begg/trimfill:\n")
        print(res_rma)
        cat("\nEgger regtest:\n")
        if (is.null(eg_obj)) cat("Egger regtest failed.\n") else print(eg_obj)
        cat("\nBegg ranktest:\n")
        if (is.null(rk_obj)) cat("Begg ranktest failed.\n") else print(rk_obj)
        cat("\nTrimfill triggered:", trigger, "\n")
        cat("Trimfill side used:", trimfill_side_used, "\n")
        cat("Trimfill k0:", ifelse(is.na(k0_best), "NA", k0_best), "\n")
        cat("Egger p after trimfill:", fmt_p(egger_p_tf), "\n\n")
        if (is.null(trf)) cat("trimfill not used or failed.\n") else print(trf)
      })
    }
    
    use_tf_for_final <- (isTRUE(trigger) && isTRUE(trf_used) && is.finite(tf_est))
    
    final_est <- if (use_tf_for_final) tf_est else primary_sm$estimate
    final_lb  <- if (use_tf_for_final) tf_lb  else primary_sm$ci.lb
    final_ub  <- if (use_tf_for_final) tf_ub  else primary_sm$ci.ub
    final_p   <- if (use_tf_for_final) tf_p   else primary_sm$pval
    
    summary_list[[v]] <- data.frame(
      Variable = v,
      n_es = n_es,
      k_studies = k_st,
      primary_model_used = primary_sm$model_used,
      primary_est = primary_sm$estimate,
      primary_lb = primary_sm$ci.lb,
      primary_ub = primary_sm$ci.ub,
      primary_p = primary_sm$pval,
      primary_sig = stars(primary_sm$pval),
      egger_p = egger_p,
      begg_p = begg_p,
      trigger_trimfill = trigger,
      trimfill_used = trf_used,
      trimfill_side_used = trimfill_side_used,
      trimfill_k0 = k0_best,
      egger_p_after_trimfill = egger_p_tf,
      tf_est = tf_est,
      tf_lb = tf_lb,
      tf_ub = tf_ub,
      tf_p = tf_p,
      tf_sig = stars(tf_p),
      use_trimfill_final = use_tf_for_final,
      final_est = final_est,
      final_lb = final_lb,
      final_ub = final_ub,
      final_p = final_p,
      final_sig = stars(final_p),
      final_source = ifelse(use_tf_for_final, "trimfill_adjusted", "primary_original"),
      stringsAsFactors = FALSE
    )
  }

  summ <- dplyr::bind_rows(summary_list) %>%
    mutate(
      egger_sig_primary = ifelse(is.na(egger_p), NA, egger_p < alpha_egger),
      egger_sig_after_trimfill = ifelse(is.na(egger_p_after_trimfill), NA, egger_p_after_trimfill < alpha_egger)
    )

  if (nrow(summ) == 0) stop("No model could be fitted.")

  # Preserve plot order: overall rows first, then order in file.
  present_overall <- overall_names[overall_names %in% summ$Variable]
  other_traits <- trait_order_in_file[trait_order_in_file %in% summ$Variable &
                                        !(trait_order_in_file %in% present_overall)]
  plot_order <- c(present_overall, other_traits)

  summ <- summ %>%
    mutate(Variable = factor(Variable, levels = plot_order)) %>%
    arrange(Variable) %>%
    mutate(Variable = as.character(Variable))

  summary_trimfill_final <- summ %>%
    transmute(
      Variable, n_es, k_studies,
      estimate = final_est,
      ci.lb = final_lb,
      ci.ub = final_ub,
      pval = final_p,
      significance = final_sig,
      final_source,
      primary_est, primary_lb, primary_ub, primary_p,
      tf_est, tf_lb, tf_ub, tf_p,
      egger_p, begg_p,
      trigger_trimfill, trimfill_used, trimfill_side_used, trimfill_k0,
      egger_p_after_trimfill
    )

  # The first sheet is exactly the set of estimates shown in the retained forest.
  # Model details retain both base and adjusted fields for provenance, not a second figure.
  safe_write_xlsx(
    list(
      forest_primary_results = summary_trimfill_final,
      model_details = summ,
      valid_row_level_effects = df,
      invalid_or_excluded_rows = invalid_rows
    ),
    out_file("forest_results.xlsx")
  )
  safe_write_csv(summ, out_file("forest_model_details.csv"), row.names = FALSE)
  safe_write_csv(summary_trimfill_final, out_file("forest_primary_results.csv"), row.names = FALSE)

  trimfill_display_table <- summary_trimfill_final %>%
    mutate(
      ci_text = sprintf("%.3f (%.3f, %.3f)", estimate, ci.lb, ci.ub),
      p_text = vapply(pval, fmt_p, character(1))
    ) %>%
    select(Variable, n_es, k_studies, estimate, ci.lb, ci.ub, pval, p_text, significance, ci_text,
           final_source, primary_est, primary_lb, primary_ub, primary_p,
           tf_est, tf_lb, tf_ub, tf_p,
           egger_p, begg_p, trigger_trimfill, trimfill_used, trimfill_side_used, trimfill_k0,
           egger_p_after_trimfill)

  safe_write_xlsx(list(forest_primary = trimfill_display_table),
                      out_file("forest_primary_statistics.xlsx"))
  safe_write_csv(trimfill_display_table, out_file("forest_primary_statistics.csv"), row.names = FALSE)

  # 5) forest plotting ----
  # v6 layout principle:
  #   The statistical x values are transformed into a fixed-width forest panel
  #   within a 0--1 plotting coordinate system. This prevents an overly wide
  #   x-axis from squeezing/cropping labels or stars, while keeping the tick
  #   labels in the original lnRR/effect-size units.
  make_plot_table <- function(tab, est_col, lb_col, ub_col, p_col) {
    tab %>%
      transmute(
        Variable = Variable,
        n_es = n_es,
        estimate = .data[[est_col]],
        ci.lb = .data[[lb_col]],
        ci.ub = .data[[ub_col]],
        pval = .data[[p_col]]
      )
  }

  primary_plot_tab <- make_plot_table(
    summ,
    est_col = "primary_est", lb_col = "primary_lb", ub_col = "primary_ub", p_col = "primary_p"
  )

  trimfill_plot_tab <- make_plot_table(
    summ,
    est_col = "final_est", lb_col = "final_lb", ub_col = "final_ub", p_col = "final_p"
  )

  # Keep the original joint CI range so the retained forest has the same axis as before.
  all_ci_for_axis <- c(primary_plot_tab$ci.lb, primary_plot_tab$ci.ub,
                       trimfill_plot_tab$ci.lb, trimfill_plot_tab$ci.ub)
  if (is.null(fixed_xlim)) {
    raw_range <- range(all_ci_for_axis, na.rm = TRUE)
    if (!is.finite(raw_range[1]) || !is.finite(raw_range[2]) || raw_range[1] == raw_range[2]) {
      raw_range <- c(-1, 1)
    }
    # Add modest padding but avoid the previous overly-expanded pretty() range.
    # Always include 0 because it is the null-effect reference line.
    raw_range <- range(c(raw_range, 0), na.rm = TRUE)
    raw_width <- diff(raw_range)
    if (!is.finite(raw_width) || raw_width <= 0) raw_width <- 1
    x_data_xlim_common <- raw_range + c(-0.08, 0.08) * raw_width
  } else {
    x_data_xlim_common <- fixed_xlim
  }

  format_x_tick <- function(x) {
    ifelse(abs(x - round(x)) < 1e-8,
           as.character(round(x)),
           formatC(x, format = "f", digits = 1))
  }

  make_x_breaks <- function(xlim, n = 5) {
    br <- pretty(xlim, n = n)
    br <- br[br >= xlim[1] - 1e-9 & br <= xlim[2] + 1e-9]
    lab <- format_x_tick(br)
    br[!duplicated(lab)]
  }

  assign_y_positions <- function(tab, plot_order, overall_names, pad_to_n_rows) {
    tab <- tab %>% mutate(Variable = as.character(Variable))
    
    real_order <- plot_order[plot_order %in% tab$Variable]
    total_display_rows <- length(real_order)
    if (!is.null(pad_to_n_rows) && is.finite(pad_to_n_rows)) {
      total_display_rows <- max(total_display_rows, pad_to_n_rows)
    }
    
    overall_vars <- real_order[real_order %in% overall_names]
    specific_vars <- real_order[!(real_order %in% overall_names)]
    
    pos_df <- data.frame(Variable = character(), y_num = numeric(), is_overall = logical())
    
    if (length(overall_vars) > 0) {
      pos_df <- rbind(
        pos_df,
        data.frame(
          Variable = overall_vars,
          y_num = seq(total_display_rows, total_display_rows - length(overall_vars) + 1),
          is_overall = TRUE,
          stringsAsFactors = FALSE
        )
      )
      specific_block_top <- total_display_rows - length(overall_vars)
    } else {
      specific_block_top <- total_display_rows
    }
    
    # Center specific rows inside the yellow block when there are padded rows.
    if (length(specific_vars) > 0) {
      available <- specific_block_top
      n_spec <- length(specific_vars)
      if (available <= n_spec) {
        y_spec <- seq(available, 1, length.out = n_spec)
      } else {
        gap <- (available - n_spec) / 2
        y_spec <- seq(available - gap, 1 + gap, length.out = n_spec)
      }
      pos_df <- rbind(
        pos_df,
        data.frame(
          Variable = specific_vars,
          y_num = y_spec,
          is_overall = FALSE,
          stringsAsFactors = FALSE
        )
      )
    }
    
    list(pos = pos_df, total_display_rows = total_display_rows,
         n_overall = length(overall_vars), specific_block_top = specific_block_top)
  }

  plot_forest_refstyle <- function(plot_tab, title, outfile_pdf, outfile_png = NULL,
                                   plot_order, overall_names, x_data_xlim_common) {
    
    pos_info <- assign_y_positions(plot_tab, plot_order, overall_names, pad_to_n_rows)
    pos_df <- pos_info$pos
    total_display_rows <- pos_info$total_display_rows
    n_overall <- pos_info$n_overall
    specific_block_top <- pos_info$specific_block_top
    
    
    # Fixed layout coordinates in a 0--1 plotting space.
    # v6 separates the label, forest, and star areas to avoid overflow.
    # Overall labels are left-aligned inside the lavender band; specific trait labels
    # are slightly indented inside the pale-yellow band.
    x_bg_min <- 0.010
    x_bg_max <- 0.990
    # Keep all trait labels left-aligned at the same x position. This avoids
    # the visually uneven indentation seen when long labels such as "Leaf Thickness"
    # are mixed with shorter labels.
    x_label_overall <- 0.030
    x_label_specific <- 0.030
    # Slightly shift the forest region left so that the forest axis title and the
    # panel title are visually more balanced within the whole figure panel.
    x_forest_min <- 0.390
    x_forest_max <- 0.735
    x_title_center <- mean(c(x_forest_min, x_forest_max))
    x_star <- 0.890
    xlim_use <- c(0, 1)
    
    plot_tab <- plot_tab %>%
      mutate(Variable_chr = as.character(Variable)) %>%
      left_join(pos_df, by = c("Variable_chr" = "Variable")) %>%
      filter(!is.na(y_num)) %>%
      mutate(
        label_text = paste0(Variable_chr, " (", n_es, ")"),
        star_label = if (isTRUE(show_significance_stars)) vapply(pval, stars, character(1)) else rep("", dplyr::n()),
        label_color = ifelse(!is.na(pval) & pval < 0.05, col_sig_text, col_nonsig_text),
        label_x = ifelse(is_overall, x_label_overall, x_label_specific)
      )
    
    # Data-value to fixed-panel-coordinate transform.
    x_min <- x_data_xlim_common[1]
    x_max <- x_data_xlim_common[2]
    if (!is.finite(x_min) || !is.finite(x_max) || x_min == x_max) {
      x_min <- -1; x_max <- 1
    }
    transform_x <- function(x) {
      x_forest_min + (x - x_min) / (x_max - x_min) * (x_forest_max - x_forest_min)
    }
    
    # Clip only CI endpoints to plotting axis limits, so extremely long intervals do not
    # push labels/stars out of the panel. Values are still kept unchanged in the result tables.
    plot_tab <- plot_tab %>%
      mutate(
        ci.lb_clip = pmax(ci.lb, x_min),
        ci.ub_clip = pmin(ci.ub, x_max),
        x_est = transform_x(estimate),
        x_lb = transform_x(ci.lb_clip),
        x_ub = transform_x(ci.ub_clip)
      )
    
    bg_df <- data.frame(
      xmin = x_bg_min, xmax = x_bg_max,
      ymin = c(if (n_overall > 0) specific_block_top + 0.5 else NA_real_, 0.5),
      ymax = c(if (n_overall > 0) total_display_rows + 0.5 else NA_real_, specific_block_top + 0.5),
      fill = c(if (n_overall > 0) col_bg_overall else NA_character_, col_bg_specific),
      stringsAsFactors = FALSE
    ) %>% filter(!is.na(ymin), !is.na(ymax), !is.na(fill))
    
    cap_df <- plot_tab %>%
      transmute(
        x_left = x_lb, x_right = x_ub,
        y = y_num,
        y_low = y_num - ci_cap_height,
        y_high = y_num + ci_cap_height
      )
    
    x_breaks <- make_x_breaks(x_data_xlim_common, n = 5)
    axis_df <- data.frame(
      x_value = x_breaks,
      x = transform_x(x_breaks),
      label = format_x_tick(x_breaks)
    )
    x_zero <- transform_x(0)
    
    # Manual x-axis is drawn exactly along the lower boundary of the yellow band,
    # so there is no white gap between the plot region and the axis line.
    y_axis <- 0.5
    y_tick_bottom <- 0.36
    y_tick_text <- 0.16
    # Move the x-axis title downward and the panel title upward so neither touches the color bands.
    y_axis_title <- -0.42
    y_panel_title <- total_display_rows + 1.16
    
    p <- ggplot() +
      annotate(
        "text",
        x = x_title_center, y = y_panel_title,
        label = title,
        size = 6.35, fontface = "bold", color = "black"
      ) +
      geom_rect(
        data = bg_df,
        aes(xmin = xmin, xmax = xmax, ymin = ymin, ymax = ymax, fill = fill),
        color = NA, inherit.aes = FALSE
      ) +
      scale_fill_identity() +
      # Restore the dashed zero-effect line. Use annotate() for fixed coordinates;
      # geom_segment(x=..., y=...) without aes/data can be ignored by ggplot2.
      annotate(
        "segment",
        x = x_zero, xend = x_zero,
        y = 0.5, yend = total_display_rows + 0.5,
        linetype = "dashed", linewidth = 0.72, color = col_zero
      ) +
      geom_segment(
        data = plot_tab,
        aes(x = x_lb, xend = x_ub, y = y_num, yend = y_num),
        linewidth = ci_line_width, color = "black", inherit.aes = FALSE
      ) +
      geom_segment(
        data = cap_df,
        aes(x = x_left, xend = x_left, y = y_low, yend = y_high),
        linewidth = ci_line_width, color = "black", inherit.aes = FALSE
      ) +
      geom_segment(
        data = cap_df,
        aes(x = x_right, xend = x_right, y = y_low, yend = y_high),
        linewidth = ci_line_width, color = "black", inherit.aes = FALSE
      ) +
      geom_point(
        data = plot_tab,
        aes(x = x_est, y = y_num),
        size = point_size, color = col_point, inherit.aes = FALSE
      ) +
      geom_text(
        data = plot_tab,
        aes(x = label_x, y = y_num, label = label_text, color = label_color),
        inherit.aes = FALSE,
        hjust = 0, vjust = 0.5, size = label_size, fontface = "bold"
      ) +
      scale_color_identity() +
      geom_text(
        data = plot_tab,
        aes(x = x_star, y = y_num, label = star_label),
        inherit.aes = FALSE,
        hjust = 0.5, vjust = 0.5, size = star_size, color = col_star, fontface = "bold"
      ) +
      # Manual axis only spans the fixed forest region, not the label/star columns.
      annotate(
        "segment",
        x = x_forest_min, xend = x_forest_max,
        y = y_axis, yend = y_axis,
        linewidth = 0.58, color = "black"
      ) +
      geom_segment(data = axis_df,
                   aes(x = x, xend = x, y = y_axis, yend = y_tick_bottom),
                   linewidth = 0.40, color = "black", inherit.aes = FALSE) +
      geom_text(data = axis_df,
                aes(x = x, y = y_tick_text, label = label),
                size = 4.12, color = "black", inherit.aes = FALSE) +
      annotate(
        "text",
        x = x_title_center, y = y_axis_title,
        label = x_axis_label,
        size = 4.70, fontface = "bold", color = "black"
      ) +
      coord_cartesian(xlim = xlim_use,
                      ylim = c(-0.82, y_panel_title + 0.14),
                      clip = "off") +
      labs(x = NULL, y = NULL) +
      theme_void(base_size = 13, base_family = "sans") +
      theme(
        plot.background = element_rect(fill = "white", color = NA),
        panel.background = element_rect(fill = "white", color = NA),
        plot.margin = margin(1, 3, 4, 3)
      )
    
    if (isTRUE(show_numeric_effect_text)) {
      plot_tab$effect_text <- sprintf("%.3f (%.3f, %.3f)",
                                      plot_tab$estimate, plot_tab$ci.lb, plot_tab$ci.ub)
      p <- p +
        geom_text(
          data = plot_tab,
          aes(x = 0.915, y = y_num, label = effect_text),
          inherit.aes = FALSE,
          hjust = 0, vjust = 0.5, size = 3.0, color = "black"
        )
    }
    
    n_rows_for_height <- if (!is.null(pad_to_n_rows) && is.finite(pad_to_n_rows)) {
      max(total_display_rows, pad_to_n_rows)
    } else {
      total_display_rows
    }
    
    height_pdf <- if (identical(height_mode, "fixed")) {
      fixed_plot_height
    } else {
      max(3.4, base_height + row_height * n_rows_for_height)
    }
    
    safe_ggsave(outfile_pdf, p, width = plot_width, height = height_pdf, dpi = 300, device = cairo_pdf)
    if (isTRUE(save_png) && !is.null(outfile_png)) {
      safe_ggsave(outfile_png, p, width = plot_width, height = height_pdf, dpi = 600)
    }
  }

  # 6) export only the retained forest (the original script's final_* estimates) ----
  plot_forest_refstyle(
    trimfill_plot_tab,
    title = group_title,
    outfile_pdf = out_file("forest_primary.pdf"),
    outfile_png = out_file("forest_primary.png"),
    plot_order = plot_order,
    overall_names = overall_names,
    x_data_xlim_common = x_data_xlim_common
  )
  cat("Forest complete: ", out_file("forest_primary.pdf"), "\n", sep = "")
  invisible(list(summary = summ, plotted_results = summary_trimfill_final))
}

run_funnel <- function(input_data, out_dir, dataset_prefix) {
  out_file <- function(suffix, directory = out_dir) {
    file.path(directory, paste0(dataset_prefix, "_", suffix))
  }
  ensure_output_dir(out_dir)
  x_axis_label <- "Effect size"

  # Bias-test settings.
  min_effects_for_bias_tests <- 10
  run_begg_ranktest <- TRUE
  run_trimfill <- TRUE
  trimfill_trigger <- "egger_or_begg"  # "egger", "begg", or "egger_or_begg"
  alpha_egger <- 0.05
  alpha_begg  <- 0.05

  # More conservative default: do not force left/right if default trimfill gives k0=0.
  # Set TRUE only for explicitly reported sensitivity analysis.
  try_both_sides_if_k0_0 <- FALSE

  # Funnel appearance.
  funnel_level <- c(90, 95, 99)
  funnel_shade <- c("white", "grey90", "grey80")
  show_contour_legend <- TRUE
  show_nk_on_plot <- TRUE

  # Variance handling.
  use_fallback_1overN <- TRUE

  # 2) helper functions ----
  safe_name <- function(x) {
    x <- as.character(x)
    x <- gsub("[/\\\\:*?\"<>|]+", "_", x)
    x <- gsub("\\s+", "_", x)
    x <- gsub("__+", "_", x)
    x
  }

  fmt_p <- function(p) {
    if (is.na(p) || !is.finite(p)) return("NA")
    if (p < 1e-4) return(formatC(p, format = "e", digits = 2))
    formatC(p, format = "f", digits = 4)
  }

  clean_key <- function(x) {
    tolower(gsub("[^A-Za-z0-9]+", "", x))
  }

  standardize_meta_columns <- function(dat) {
    aliases <- list(
      StudyID  = c("studyid", "study", "id", "number", "refid", "referenceid"),
      Variable = c("variable", "trait", "traits", "index", "category"),
      Mc       = c("mc", "meanc", "meancontrol", "controlmean", "native_mean", "nativemean", "home_mean", "homemean"),
      SDc      = c("sdc", "sdcontrol", "controlsd", "native_sd", "nativesd", "home_sd", "homesd"),
      Nc       = c("nc", "ncontrol", "controln", "native_n", "nativen", "home_n", "homen"),
      Mi       = c("mi", "meani", "meaninvasive", "invasivemean", "introduced_mean", "introducedmean", "away_mean", "awaymean"),
      Sdi      = c("sdi", "sdi", "sdinvasive", "invasivesd", "introduced_sd", "introducedsd", "away_sd", "awaysd"),
      Ni       = c("ni", "ninvasive", "invasiven", "introduced_n", "introducedn", "away_n", "awayn")
    )
    
    original_names <- names(dat)
    keys <- clean_key(original_names)
    
    for (canonical in names(aliases)) {
      hit <- which(keys %in% aliases[[canonical]])
      if (length(hit) == 0) {
        stop(
          "Missing required column for ", canonical, ".\n",
          "Current column names are: ", paste(original_names, collapse = ", ")
        )
      }
      names(dat)[hit[1]] <- canonical
    }
    
    dat
  }

  eps <- 1e-12

  compute_row <- function(Mc, SDc, Nc, Mi, Sdi, Ni) {
    yi <- NA_real_
    vi <- NA_real_
    method <- NA_character_
    
    if (!is.na(Mc) && !is.na(Mi) && Mc > 0 && Mi > 0) {
      yi <- log(Mi / Mc)
      
      if (!is.na(SDc) && !is.na(Sdi) && is.finite(SDc) && is.finite(Sdi) &&
          !is.na(Nc) && !is.na(Ni) && Nc > 0 && Ni > 0) {
        vi <- (Sdi^2) / (Ni * (Mi^2 + eps)) + (SDc^2) / (Nc * (Mc^2 + eps))
        method <- "SD_formula"
        if (!is.finite(vi) || vi <= 0) vi <- NA_real_
      }
      
      if (is.na(vi) && isTRUE(use_fallback_1overN) &&
          !is.na(Nc) && !is.na(Ni) && Nc > 0 && Ni > 0) {
        vi <- 1 / Nc + 1 / Ni
        method <- "1/N_fallback"
        if (!is.finite(vi) || vi <= 0) vi <- NA_real_
      }
    } else {
      method <- "invalid_mean"
    }
    
    list(
      yi = yi,
      vi = vi,
      sei = ifelse(is.na(vi), NA_real_, sqrt(vi)),
      method = method
    )
  }

  should_trimfill <- function(egger_p, begg_p) {
    egger_sig <- is.finite(egger_p) && !is.na(egger_p) && egger_p < alpha_egger
    begg_sig  <- is.finite(begg_p)  && !is.na(begg_p)  && begg_p  < alpha_begg
    
    if (trimfill_trigger == "egger") return(egger_sig)
    if (trimfill_trigger == "begg") return(begg_sig)
    egger_sig || begg_sig
  }

  draw_funnel <- function(model_obj, var_title, egger_p, begg_p,
                          n_effect_sizes, k_studies, trimfill_k0 = NA,
                          plot_note = NULL) {
    metafor::funnel(
      model_obj,
      refline = 0,
      level = funnel_level,
      shade = funnel_shade,
      xlab = x_axis_label,
      ylab = "Standard Error",
      cex = 0.85
    )
    
    title(main = var_title, cex.main = 1.05, font.main = 2)
    
    if (show_contour_legend) {
      legend(
        x = "topleft",
        cex = 0.75,
        legend = c("0.1 > p > 0.05", "0.05 > p > 0.01", "p < 0.01"),
        fill = funnel_shade,
        bty = "n"
      )
    }
    
    usr <- par("usr")
    info_txt <- paste0("Egger P = ", fmt_p(egger_p))
    if (run_begg_ranktest) {
      info_txt <- paste0(info_txt, "\nBegg P = ", fmt_p(begg_p))
    }
    if (!is.na(trimfill_k0)) {
      info_txt <- paste0(info_txt, "\nTrim-and-fill k0 = ", trimfill_k0)
    }
    if (show_nk_on_plot) {
      info_txt <- paste0(info_txt, "\nn = ", n_effect_sizes, ", k = ", k_studies)
    }
    if (!is.null(plot_note)) {
      info_txt <- paste0(info_txt, "\n", plot_note)
    }
    
    text(
      x = usr[2] - 0.02 * (usr[2] - usr[1]),
      y = usr[4] - 0.05 * (usr[4] - usr[3]),
      labels = info_txt,
      adj = c(1, 1),
      cex = 0.78
    )
  }

  save_funnel_pdf_png <- function(model_obj, var_title, egger_p, begg_p,
                                  n_effect_sizes, k_studies, trimfill_k0,
                                  plot_note, out_pdf, out_png) {
    draw_one <- function() {
      op <- par(mar = c(4.4, 4.4, 2.4, 1.2))
      draw_funnel(model_obj, var_title, egger_p, begg_p,
                  n_effect_sizes, k_studies, trimfill_k0, plot_note)
      par(op)
    }
    safe_plot_device(out_pdf,
      function(path) pdf(path, width = 6.4, height = 5.0), draw_one)
    safe_plot_device(out_png,
      function(path) png(path, width = 1900, height = 1500, res = 300), draw_one)
  }

  # 3) read data and compute row-level effect sizes ----
  raw <- input_data
  raw <- standardize_meta_columns(raw)

  df0 <- raw %>%
    mutate(
      StudyID  = as.character(StudyID),
      Variable = as.character(Variable),
      Mc  = as.numeric(Mc),
      SDc = as.numeric(SDc),
      Nc  = as.numeric(Nc),
      Mi  = as.numeric(Mi),
      Sdi = as.numeric(Sdi),
      Ni  = as.numeric(Ni),
      .rowid = dplyr::row_number()
    )

  tmp <- purrr::pmap(
    df0 %>% dplyr::select(Mc, SDc, Nc, Mi, Sdi, Ni),
    compute_row
  )

  df_all <- df0 %>%
    mutate(
      yi = purrr::map_dbl(tmp, "yi"),
      vi = purrr::map_dbl(tmp, "vi"),
      sei = purrr::map_dbl(tmp, "sei"),
      Var_method = purrr::map_chr(tmp, "method")
    )

  df <- df_all %>%
    filter(
      !is.na(StudyID), !is.na(Variable),
      !is.na(yi), !is.na(vi),
      is.finite(yi), is.finite(vi), vi > 0
    )

  invalid_rows <- df_all %>%
    filter(
      is.na(StudyID) | is.na(Variable) |
        is.na(yi) | is.na(vi) |
        !is.finite(yi) | !is.finite(vi) | vi <= 0
    )

  if (nrow(df) == 0) stop("No valid effect sizes after filtering. Please check the input data.")

  safe_write_xlsx(
    list(valid_row_level_effects = df, invalid_or_excluded_rows = invalid_rows),
    out_file("funnel_row_level_effects.xlsx")
  )

  # 4) per-variable funnels and bias tests ----
  trait_order_in_file <- unique(df$Variable)
  summary_list <- list()

  for (v in trait_order_in_file) {
    dv <- df %>% filter(Variable == v)
    n_es <- nrow(dv)
    k_st <- length(unique(dv$StudyID))
    
    v_dir <- file.path(out_dir, paste0("Variable_", safe_name(v)))
    ensure_output_dir(v_dir)
    
    if (n_es < 3) {
      summary_list[[v]] <- data.frame(
        Variable = v,
        n_effect_sizes = n_es,
        k_studies = k_st,
        estimate = NA_real_,
        ci.lb = NA_real_,
        ci.ub = NA_real_,
        pval = NA_real_,
        egger_p = NA_real_,
        begg_p = NA_real_,
        bias_test_note = "Too few effect sizes for funnel/meta-analysis",
        trimfill_triggered = FALSE,
        trimfill_k0 = NA_integer_,
        tf_estimate = NA_real_,
        tf_ci.lb = NA_real_,
        tf_ci.ub = NA_real_,
        tf_pval = NA_real_,
        stringsAsFactors = FALSE
      )
      next
    }
    
    res <- tryCatch(
      metafor::rma(yi = dv$yi, vi = dv$vi, method = "REML"),
      error = function(e) NULL
    )
    
    if (is.null(res)) {
      summary_list[[v]] <- data.frame(
        Variable = v,
        n_effect_sizes = n_es,
        k_studies = k_st,
        estimate = NA_real_,
        ci.lb = NA_real_,
        ci.ub = NA_real_,
        pval = NA_real_,
        egger_p = NA_real_,
        begg_p = NA_real_,
        bias_test_note = "rma() failed",
        trimfill_triggered = FALSE,
        trimfill_k0 = NA_integer_,
        tf_estimate = NA_real_,
        tf_ci.lb = NA_real_,
        tf_ci.ub = NA_real_,
        tf_pval = NA_real_,
        stringsAsFactors = FALSE
      )
      next
    }
    
    egger_p <- NA_real_
    begg_p <- NA_real_
    bias_note <- "Bias tests run"
    
    if (n_es >= min_effects_for_bias_tests) {
      eg_obj <- tryCatch(
        metafor::regtest(res, model = "rma", predictor = "sei"),
        error = function(e) NULL
      )
      if (!is.null(eg_obj)) egger_p <- tryCatch(as.numeric(eg_obj$pval), error = function(e) NA_real_)
      
      if (isTRUE(run_begg_ranktest)) {
        rk_obj <- tryCatch(metafor::ranktest(res), error = function(e) NULL)
        if (!is.null(rk_obj)) begg_p <- tryCatch(as.numeric(rk_obj$pval), error = function(e) NA_real_)
      }
    } else {
      bias_note <- paste0("Egger/Begg not run: n_effect_sizes < ", min_effects_for_bias_tests)
    }
    
    # Original funnel.
    out_pdf <- out_file(paste0("funnel_original_", safe_name(v), ".pdf"), v_dir)
    out_png <- out_file(paste0("funnel_original_", safe_name(v), ".png"), v_dir)
    
    plot_note <- if (n_es < min_effects_for_bias_tests) {
      paste0("Bias tests not run: n < ", min_effects_for_bias_tests)
    } else {
      NULL
    }
    
    save_funnel_pdf_png(
      model_obj = res,
      var_title = v,
      egger_p = egger_p,
      begg_p = begg_p,
      n_effect_sizes = n_es,
      k_studies = k_st,
      trimfill_k0 = NA,
      plot_note = plot_note,
      out_pdf = out_pdf,
      out_png = out_png
    )
    
    # Save text output of model and bias tests.
    safe_text_report(out_file(paste0("funnel_", safe_name(v), "_model_tests.txt"), v_dir), {
      cat("Variable:", v, "\n")
      cat("n(effect sizes) =", n_es, "; k(studies) =", k_st, "\n\n")
      cat("Random-effects model (rma REML; used for funnel/asymmetry diagnostics):\n")
      print(res)
      cat("\nEgger test:\n")
      if (n_es < min_effects_for_bias_tests) {
        cat("Not run because n_effect_sizes <", min_effects_for_bias_tests, "\n")
      } else {
        print(tryCatch(metafor::regtest(res, model = "rma", predictor = "sei"), error = function(e) e$message))
      }
      cat("\nBegg rank correlation test:\n")
      if (!isTRUE(run_begg_ranktest)) {
        cat("Not requested.\n")
      } else if (n_es < min_effects_for_bias_tests) {
        cat("Not run because n_effect_sizes <", min_effects_for_bias_tests, "\n")
      } else {
        print(tryCatch(metafor::ranktest(res), error = function(e) e$message))
      }
    })
    
    # Trim-and-fill sensitivity.
    tf <- NULL
    tf_trigger <- FALSE
    tf_k0 <- NA_integer_
    tf_est <- tf_lb <- tf_ub <- tf_p <- NA_real_
    
    if (isTRUE(run_trimfill) && n_es >= min_effects_for_bias_tests) {
      tf_trigger <- should_trimfill(egger_p, begg_p)
      
      if (tf_trigger) {
        tf <- tryCatch(metafor::trimfill(res), error = function(e) NULL)
        
        if (isTRUE(try_both_sides_if_k0_0) && !is.null(tf)) {
          k0_default <- tryCatch(as.integer(tf$k0), error = function(e) NA_integer_)
          if (!is.na(k0_default) && k0_default == 0) {
            tf_left  <- tryCatch(metafor::trimfill(res, side = "left"),  error = function(e) NULL)
            tf_right <- tryCatch(metafor::trimfill(res, side = "right"), error = function(e) NULL)
            
            k0_left  <- if (!is.null(tf_left))  tryCatch(as.integer(tf_left$k0),  error = function(e) NA_integer_) else NA_integer_
            k0_right <- if (!is.null(tf_right)) tryCatch(as.integer(tf_right$k0), error = function(e) NA_integer_) else NA_integer_
            
            k0s <- c(default = k0_default, left = k0_left, right = k0_right)
            candidates <- list(default = tf, left = tf_left, right = tf_right)
            best_name <- names(which.max(ifelse(is.na(k0s), -Inf, k0s)))[1]
            tf <- candidates[[best_name]]
          }
        }
        
        if (!is.null(tf)) {
          tf_k0 <- tryCatch(as.integer(tf$k0), error = function(e) NA_integer_)
          tf_est <- tryCatch(as.numeric(tf$b[1]), error = function(e) NA_real_)
          tf_lb  <- tryCatch(as.numeric(tf$ci.lb[1]), error = function(e) NA_real_)
          tf_ub  <- tryCatch(as.numeric(tf$ci.ub[1]), error = function(e) NA_real_)
          tf_p   <- tryCatch(as.numeric(tf$pval[1]), error = function(e) NA_real_)
          
          out_tf_pdf <- out_file(paste0("funnel_trimfill_", safe_name(v), ".pdf"), v_dir)
          out_tf_png <- out_file(paste0("funnel_trimfill_", safe_name(v), ".png"), v_dir)
          
          save_funnel_pdf_png(
            model_obj = tf,
            var_title = paste0(v, " (trim-and-fill sensitivity)"),
            egger_p = egger_p,
            begg_p = begg_p,
            n_effect_sizes = n_es,
            k_studies = k_st,
            trimfill_k0 = tf_k0,
            plot_note = NULL,
            out_pdf = out_tf_pdf,
            out_png = out_tf_png
          )
          
          safe_text_report(out_file(paste0("funnel_", safe_name(v), "_trimfill_details.txt"), v_dir), {
            cat("Variable:", v, "\n")
            cat("Trim-and-fill triggered by rule:", trimfill_trigger, "\n")
            cat("Egger P =", fmt_p(egger_p), "; Begg P =", fmt_p(begg_p), "\n")
            cat("k0 =", tf_k0, "\n\n")
            print(tf)
          })
        }
      }
    }
    
    summary_list[[v]] <- data.frame(
      Variable = v,
      n_effect_sizes = n_es,
      k_studies = k_st,
      estimate = as.numeric(res$b[1]),
      ci.lb = as.numeric(res$ci.lb[1]),
      ci.ub = as.numeric(res$ci.ub[1]),
      pval = as.numeric(res$pval[1]),
      egger_p = egger_p,
      begg_p = begg_p,
      bias_test_note = bias_note,
      trimfill_triggered = tf_trigger,
      trimfill_k0 = tf_k0,
      tf_estimate = tf_est,
      tf_ci.lb = tf_lb,
      tf_ci.ub = tf_ub,
      tf_pval = tf_p,
      stringsAsFactors = FALSE
    )
  }

  summ <- dplyr::bind_rows(summary_list) %>%
    mutate(
      egger_p_formatted = vapply(egger_p, fmt_p, character(1)),
      begg_p_formatted = vapply(begg_p, fmt_p, character(1)),
      egger_significant = ifelse(is.na(egger_p), NA, egger_p < alpha_egger),
      begg_significant = ifelse(is.na(begg_p), NA, begg_p < alpha_begg)
    )

  safe_write_xlsx(
    list(publication_bias_summary = summ),
    out_file("funnel_publication_bias_summary.xlsx")
  )
  safe_write_csv(summ, out_file("funnel_publication_bias_summary.csv"), row.names = FALSE)

  cat("Individual funnels complete: ", out_dir, "\n", sep = "")
  invisible(summ)
}

# 4) VALIDATE ALL FOUR WORKSHEETS BEFORE STARTING ANY ANALYSIS ----
available_sheets <- readxl::excel_sheets(infile)
missing_sheets <- setdiff(sheet_names, available_sheets)
if (length(missing_sheets) > 0L) {
  stop("Missing required worksheet(s): ", paste(missing_sheets, collapse = ", "),
       "\nUse the exact sheet names F1_A, F1_B, F1_C and F1_D.")
}
extra_sheets <- setdiff(available_sheets, sheet_names)
if (length(extra_sheets) > 0L) {
  message("Ignoring extra worksheet(s): ", paste(extra_sheets, collapse = ", "))
}
required_columns <- c("StudyID", "Variable", "Mc", "SDc", "Nc", "Mi", "Sdi", "Ni")
input_sheets <- setNames(lapply(sheet_names, function(sheet_name) {
  sheet_data <- readxl::read_excel(infile, sheet = sheet_name)
  missing_columns <- setdiff(required_columns, names(sheet_data))
  if (length(missing_columns) > 0L) {
    stop("Worksheet ", sheet_name, " is missing column(s): ",
         paste(missing_columns, collapse = ", "))
  }
  if (nrow(sheet_data) == 0L ||
      all(is.na(as.matrix(sheet_data[, required_columns])))) {
    stop("Worksheet ", sheet_name, " has no data. Fill all four worksheets before running.")
  }
  sheet_data
}), sheet_names)

# 5) RUN EACH WORKSHEET IN ITS OWN OUTPUT SUBFOLDER ----
# Never pool records across the four panels.
preflight_output_paths(output_root, input_sheets, sheets_to_run)
input_md5 <- unname(tools::md5sum(infile))
run_summary <- data.frame(
  workbook = rep(basename(infile), length(sheet_names)),
  input_md5 = rep(input_md5, length(sheet_names)),
  sheet = sheet_names,
  input_rows = vapply(input_sheets, nrow, integer(1)),
  output_subfolder = sheet_names,
  status = ifelse(sheet_names %in% sheets_to_run, "not_started", "not_requested"),
  error_message = rep("", length(sheet_names)),
  stringsAsFactors = FALSE,
  row.names = NULL
)
summary_file <- file.path(output_root, paste0(workbook_prefix, "_run_summary.csv"))
save_run_summary <- function() {
  primary_error <- tryCatch({
    safe_write_csv(run_summary, summary_file, row.names = FALSE)
    NULL
  }, error = function(e) e)
  if (is.null(primary_error)) return(invisible(TRUE))
  message("Run-summary save failed; analysis is not stopped by this status-file error.\n",
          conditionMessage(primary_error))
  fallback <- file.path(output_root, paste0(workbook_prefix, "_run_summary_",
                         format(Sys.time(), "%Y%m%d_%H%M%S"), ".csv"))
  if (identical(fallback, summary_file) || file.exists(fallback)) {
    fallback <- tempfile(pattern = paste0(workbook_prefix, "_run_summary_"),
                         tmpdir = output_root, fileext = ".csv")
  }
  tryCatch({
    safe_write_csv(run_summary, fallback, row.names = FALSE)
    summary_file <<- fallback
    message("Run summary now saved to: ", fallback)
    invisible(TRUE)
  }, error = function(e) {
    message("Run summary could not be saved; its data remain in run_summary in R.\n",
            conditionMessage(e))
    invisible(FALSE)
  })
}
save_run_summary()
all_results <- setNames(vector("list", length(sheet_names)), sheet_names)

for (run_index in seq_along(sheets_to_run)) {
  dataset_prefix <- sheets_to_run[run_index]
  i <- match(dataset_prefix, sheet_names)
  output_dir <- file.path(output_root, dataset_prefix)
  input_data <- input_sheets[[dataset_prefix]]
  message("\n[", run_index, "/", length(sheets_to_run), "] Worksheet: ", dataset_prefix,
          " (", nrow(input_data), " input rows)")
  run_summary$status[i] <- "running"
  save_run_summary()

  current_stage <- "forest"
  panel_warnings <- character()
  all_results[[dataset_prefix]] <- tryCatch(withCallingHandlers({
    forest_results <- run_forest(input_data, output_dir, dataset_prefix,
                                 group_titles[[dataset_prefix]])
    current_stage <- "funnel"
    funnel_results <- run_funnel(input_data, file.path(output_dir, "funnel"),
                                 dataset_prefix)

    current_stage <- "session record"
    if (length(panel_warnings)) {
      warning_file <- file.path(output_dir, paste0(dataset_prefix, "_warnings.txt"))
      tryCatch(safe_write_lines(panel_warnings, warning_file),
        error = function(warning_error) message("Warning log could not be saved: ", conditionMessage(warning_error)))
    }
    # Record both workbook and worksheet so every panel is traceable.
    run_info <- c(
      paste("Input file:", normalizePath(infile, winslash = "/", mustWork = TRUE)),
      paste("Input MD5:", input_md5),
      paste("Input worksheet:", dataset_prefix),
      paste("Input rows:", nrow(input_data)),
      paste("Dataset prefix:", dataset_prefix),
      paste("Run time (UTC):", format(Sys.time(), tz = "UTC", usetz = TRUE)),
      "Forest plotted values: final_* from the supplied forest script.",
      "Forest settings: StudyID random effect TRUE; fallback TRUE; Egger/Begg alpha 0.05;",
      "  bias-triggered trimfill TRUE; try both sides when default k0=0 TRUE.",
      "Funnel settings: rma REML; at least 10 effects for Egger/Begg; fallback TRUE;",
      "  Egger-or-Begg-triggered trimfill TRUE; try both sides when default k0=0 FALSE.",
      "Only individual trait funnel plots are exported; there is no combined funnel PDF.",
      "", capture.output(sessionInfo())
    )
    safe_write_lines(run_info, file.path(output_dir, paste0(dataset_prefix, "_session_info.txt")))
    list(forest = forest_results, funnel = funnel_results)
  }, warning = function(w) {
    panel_warnings <<- c(panel_warnings,
      paste0("Warning: ", conditionMessage(w), "\nCall: ",
             paste(deparse(conditionCall(w)), collapse = " ")))
    # Leave the warning visible in R as well as recording it here.
  }), error = function(e) {
    original_error <- paste0("Worksheet ", dataset_prefix, " failed during ",
                             current_stage, ": ", conditionMessage(e))
    run_summary$status[i] <<- "failed"
    run_summary$error_message[i] <<- original_error
    # Report the original failure BEFORE trying to write diagnostics.
    message(original_error)
    diagnostic_file <- file.path(output_root, paste0(dataset_prefix, "_error.txt"))
    tryCatch(safe_write_lines(c(original_error, "",
      "Warnings captured for this worksheet:",
      if (length(panel_warnings)) panel_warnings else "None captured."), diagnostic_file),
      error = function(log_error) message("Error log could not be saved: ", conditionMessage(log_error)))
    save_run_summary()
    stop(original_error, "\nEarlier completed panels remain in their subfolders.",
         call. = FALSE)
  })
  run_summary$status[i] <- "completed"
  save_run_summary()
}

cat("\nDONE. Selected worksheets completed: ", paste(sheets_to_run, collapse = ", "),
    "\nResults are in:\n",
    normalizePath(output_root, winslash = "/"), "\n")
cat("Each selected worksheet has its own subfolder and filename prefix.\n")
cat("Retained forest: <sheet>_forest_primary.pdf/png\n")
cat("Plotted estimates: <sheet>_forest_results.xlsx, sheet forest_primary_results\n")
cat("Run summary: ", summary_file, "\n", sep = "")
