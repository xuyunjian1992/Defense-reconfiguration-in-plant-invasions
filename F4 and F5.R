############################################################
# Figures 4 and 5: introduced-range versus native-range plants
# Standalone script with Figure S6 funnel plots and result tables.
# Run this entire file. No external helper scripts are needed.
# All manuscript figures use M2-RAW; IQR is numerical sensitivity analysis.
# Reads local input only; no automatic installation or network operations.
############################################################
local({
# Optional settings. Leave NULL to use automatic local/project discovery.
file_name <- "Table S2.xlsx"
original_input_file <- NULL
sheet_name <- 1
project_root_override <- NULL
input_file_override <- NULL
output_root_override <- NULL
clean_path_value <- function(x) {
 if(is.null(x) || length(x)!=1L || is.na(x) || !nzchar(x)) return("")
 path.expand(as.character(x))
}
project_candidates <- function() {
 source_files <- vapply(sys.frames(),function(frame) {
   tryCatch(clean_path_value(get0("ofile",envir=frame,inherits=FALSE)),error=function(e) "")
 },character(1))
 source_files <- rev(source_files[nzchar(source_files)])
 # RStudio's Run/Run All may have no source() frame or --file argument.
 editor_file <- ""
 if(requireNamespace("rstudioapi",quietly=TRUE)) {
   editor_file <- tryCatch({
     if(rstudioapi::isAvailable()) clean_path_value(rstudioapi::getSourceEditorContext()$path) else ""
   },error=function(e) "")
 }
 args <- commandArgs(trailingOnly=FALSE)
 script_files <- sub("^--file=","",args[grepl("^--file=",args)])
 script_files <- gsub("~+~"," ",script_files,fixed=TRUE)
 files <- unique(c(source_files,script_files,editor_file))
 files <- files[nzchar(files) & file.exists(files) & !dir.exists(files)]
 dirs <- if(length(files)) dirname(normalizePath(files,winslash="/",mustWork=TRUE)) else character()
 cwd <- normalizePath(getwd(),winslash="/",mustWork=TRUE)
 # A source/Rscript context is preferred. The current folder and its parent
 # also support pasted code and execution from a project's results folder.
 unique(c(dirs,cwd,dirname(cwd)))
}
resolve_input_location <- function(file_name, original_file=NULL,
  project_root_override=NULL, input_file_override=NULL,
  candidates=project_candidates(), interactive_session=interactive(), chooser=file.choose) {
 explicit_input <- clean_path_value(input_file_override)
 explicit_root <- clean_path_value(project_root_override)
 accept <- function(path, root, method) {
   if(!file.exists(path) || dir.exists(path)) stop("Input workbook not found: ",path,call.=FALSE)
   if(tolower(tools::file_ext(path))!="xlsx") stop("Select an .xlsx workbook: ",path,call.=FALSE)
   list(input=normalizePath(path,winslash="/",mustWork=TRUE),
        root=normalizePath(root,winslash="/",mustWork=TRUE),method=method)
 }
 infer_root <- function(path) {
   d <- dirname(normalizePath(path,winslash="/",mustWork=TRUE))
   if(tolower(basename(d))=="data") dirname(d) else d
 }
 if(nzchar(explicit_input)) {
   if(!file.exists(explicit_input) || dir.exists(explicit_input))
     stop("input_file_override does not point to an existing workbook: ",explicit_input,call.=FALSE)
   root <- if(nzchar(explicit_root)) explicit_root else infer_root(explicit_input)
   if(!dir.exists(root)) stop("project_root_override is not a directory: ",root,call.=FALSE)
   return(accept(explicit_input,root,"input_file_override"))
 }
 if(nzchar(explicit_root)) {
   if(!dir.exists(explicit_root)) stop("project_root_override is not a directory: ",explicit_root,call.=FALSE)
   paths <- c(file.path(explicit_root,"data",file_name),file.path(explicit_root,file_name))
   found <- paths[file.exists(paths) & !dir.exists(paths)]
   if(!length(found)) stop("Cannot find ",file_name," under project_root_override: ",explicit_root,call.=FALSE)
   return(accept(found[1],explicit_root,"project_root_override"))
 }
 original_file <- clean_path_value(original_file)
 if(nzchar(original_file) && file.exists(original_file) && !dir.exists(original_file))
   return(accept(original_file,dirname(original_file),"original_local_path"))
 for(root in unique(candidates)) {
   if(!dir.exists(root)) next
   paths <- c(file.path(root,"data",file_name),file.path(root,file_name))
   found <- paths[file.exists(paths) & !dir.exists(paths)]
   if(length(found)) return(accept(found[1],root,"script_or_working_directory"))
 }
 if(isTRUE(interactive_session)) {
   message("Could not locate ",file_name," automatically. Select the input workbook in the file dialog.")
   selected <- tryCatch(chooser(),error=function(e) stop("No workbook selected. Set input_file_override at the top of the script and rerun.",call.=FALSE))
   if(!file.exists(selected) || dir.exists(selected)) stop("No valid workbook was selected.",call.=FALSE)
   return(accept(selected,infer_root(selected),"interactive_file_selection"))
 }
 stop("Cannot find ",file_name,". Set input_file_override or project_root_override in section 1, ",
      "or run this file from its extracted project. No external helper files are needed.",call.=FALSE)
}

located <- resolve_input_location(file_name,original_input_file,project_root_override,input_file_override)
project_root <- located$root
input_file <- located$input
out_dir <- file.path(project_root,"F4_F5_results")

pkgs <- c("readxl","dplyr","tidyr","stringr","ggplot2","writexl","tibble","purrr","patchwork","clubSandwich","scales","metafor")
missing_pkgs <- pkgs[!vapply(pkgs,requireNamespace,logical(1),quietly=TRUE)]
if(length(missing_pkgs)) stop("Missing R packages: ",paste(missing_pkgs,collapse=", "),
 ". Install them separately as described in README.md, then rerun. No packages were installed automatically.",call.=FALSE)
invisible(lapply(pkgs,function(p) suppressPackageStartupMessages(library(p,character.only=TRUE))))
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
  temp_path <- tempfile(pattern = ".f23_", tmpdir = dirname(path),
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
  probe <- tempfile(pattern = ".f23check_", tmpdir = path)
  probe <- paste0(probe, strrep("x", max(0, filename_length - nchar(basename(probe)))))
  on.exit(unlink(probe), add = TRUE)
  safe_write_lines("F4/F5 output write check", probe)
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


# All file writers are defined in this script. Statistical calculations are unchanged.
read_f45_input <- function(path, sheet) {
 preview <- readxl::read_excel(path,sheet=sheet,col_names=FALSE,col_types="text",n_max=12,.name_repair="minimal")
 key <- function(x) gsub("[^a-z0-9]","",tolower(trimws(as.character(x))))
 h <- which(vapply(seq_len(nrow(preview)),function(i) {
   keys <- key(unlist(preview[i,],use.names=FALSE))
   all(c("studyid","systemid","trait","mi","mc","sdi","sdc","ni","nc") %in% keys)
 },logical(1)))
 if(length(h)!=1L) stop("Cannot uniquely identify the input header within the first 12 worksheet rows.",call.=FALSE)
 x <- readxl::read_excel(path,sheet=sheet,skip=h-1L,.name_repair="minimal")
 nm <- trimws(names(x))
 blank <- is.na(nm)|!nzchar(nm)
 nm[blank] <- paste0("Unnamed_column_",which(blank))
 aliases <- c("Study ID"="StudyID","System ID"="SystemID","Sdi"="SDi")
 for(old in names(aliases)) nm[nm==old] <- aliases[[old]]
 if(anyDuplicated(nm)) stop("Duplicate column names after input header mapping.")
 names(x) <- nm
 attr(x,"source_header_row") <- as.integer(h)
 x
}
main_plot_location <- function(stem, directory) {
 if(grepl("^Figure4_panel_[A-D]_FINAL_",stem))
   return(list(stem=paste0("F4",sub("^Figure4_panel_([A-D])_.*","\\1",stem)),directory=file.path(out_dir,"F4")))
 if(grepl("^Figure5_panel_[A-F]_FINAL_",stem))
   return(list(stem=paste0("F5",sub("^Figure5_panel_([A-F])_.*","\\1",stem)),directory=file.path(out_dir,"F5")))
 if(identical(stem,"Figure4_FINAL_M2_RAW_model_based_95CI"))
   return(list(stem="F4",directory=file.path(out_dir,"F4")))
 if(identical(stem,"Figure5_FINAL_six_panels_M2_RAW_model_based_95CI"))
   return(list(stem="F5",directory=file.path(out_dir,"F5")))
 list(stem=stem,directory=directory)
}
save_plot_pair <- function(plot_obj,stem,directory,width=pdf_w,height=pdf_h,save_tiff=TRUE) {
 loc <- main_plot_location(stem,directory)
 pdf_path <- file.path(loc$directory,paste0(loc$stem,".pdf"))
 assert_output_path(pdf_path)
 safe_ggsave(pdf_path,plot=plot_obj,width=width,height=height,units="in",device="pdf")
 if(isTRUE(save_tiff)) {
   tiff_path <- file.path(loc$directory,paste0(loc$stem,".tiff"))
   assert_output_path(tiff_path)
   safe_ggsave(tiff_path,plot=plot_obj,width=width,height=height,units="in",dpi=tiff_dpi,compression="lzw",device="tiff")
 }
}
# Presentation-only settings measured from supplementary Figure S5.
# Keep the one-line title, plot-internal annotations, three grey tones and
# the original title/axis : annotation font-size ratio of 7 : 5.
manuscript_funnel_p <- function(p, digits = 3L) {
  if (length(p) != 1L || !is.finite(p)) return("NA")
  cutoff <- 10^(-digits)
  if (p < cutoff) return(paste0("< ", formatC(cutoff, format = "f", digits = digits)))
  paste0("= ", formatC(p, format = "f", digits = digits))
}
draw_manuscript_funnel <- function(model, title, egger_p, begg_p,
    trimfill_k0, n_effects, n_studies, compact = FALSE, p_digits = 3L) {
  # Font size follows the physical plotting width. At the same insertion
  # width this reproduces the reference, regardless of PDF canvas size.
  plot_width_pt <- graphics::par("pin")[1] * 72
  axis_pt <- plot_width_pt * (4.7762 / 142.65)
  title_pt <- axis_pt
  annotation_pt <- title_pt * (5 / 7)
  metafor::funnel(model, refline = 0, level = c(90, 95, 99),
    shade = c("white", "grey90", "grey80"), back = "grey80",
    hlines = "white", steps = 5, xlab = "", ylab = "",
    cex = 0.7972 * plot_width_pt / 360,
    cex.axis = axis_pt / 12, cex.lab = axis_pt / 12)
  usr <- graphics::par("usr")
  left_in <- graphics::grconvertX(usr[1], "user", "inches")
  right_in <- graphics::grconvertX(usr[2], "user", "inches")
  top_in <- graphics::grconvertY(usr[4], "user", "inches")
  bottom_in <- graphics::grconvertY(usr[3], "user", "inches")
  x_user <- function(inches) graphics::grconvertX(inches, "inches", "user")
  y_user <- function(inches) graphics::grconvertY(inches, "inches", "user")
  graphics::text(x_user((left_in + right_in) / 2),
    y_user(bottom_in - 2.85 * axis_pt / 72), "lnRR",
    cex = axis_pt / 12, adj = c(0.5, 0.5), xpd = NA)
  graphics::text(x_user(left_in - 4 * axis_pt / 72),
    y_user((top_in + bottom_in) / 2), "Standard Error", srt = 90,
    cex = axis_pt / 12, adj = c(0.5, 0.5), xpd = NA)

  # The original title slightly extends beyond the axes but fits its cell.
  cell_width_pt <- graphics::par("fin")[1] * 72
  title_width_in <- graphics::strwidth(title, units = "inches", cex = title_pt / 12, font = 2)
  if (title_width_in * 72 > cell_width_pt - 8)
    stop("Funnel title does not fit its panel: ", title)
  cell_left_in <- left_in - graphics::par("mai")[2]
  cell_right_in <- cell_left_in + cell_width_pt / 72
  title_center_in <- min(max((left_in + right_in) / 2,
    cell_left_in + 4 / 72 + title_width_in / 2),
    cell_right_in - 4 / 72 - title_width_in / 2)
  graphics::text(x_user(title_center_in), y_user(top_in + 1.10 * title_pt / 72),
    title, cex = title_pt / 12, font = 2, adj = c(0.5, 0.5), xpd = NA)

  # These labels describe the actual fills. The supplied figure's original
  # labels did not match the shaded normal-theory P regions.
  graphics::legend("topleft", bty = "n", cex = annotation_pt / 12,
    legend = c("p >= 0.10", "0.05 <= p < 0.10", "p < 0.05"),
    fill = c("white", "grey90", "grey80"))
  info <- paste0("Egger P ", manuscript_funnel_p(egger_p, p_digits),
    "\nBegg P ", manuscript_funnel_p(begg_p, p_digits),
    "\nTrim-and-fill k0 = ",
    if (length(trimfill_k0) == 1L && is.finite(trimfill_k0))
      format(trimfill_k0, trim = TRUE) else "NA",
    "\nn = ", n_effects, ", k = ", n_studies)
  # Small panel-specific typography offsets keep the reference-style
  # diagnostic block clear of observed points. They change no plot values,
  # axis limits, fitted models, source rows or diagnostic calculations.
  info_shift_down_pt <- 0
  if (identical(title, "lnRR Plant nitrogen content VS. Phenolics-Plant nitrogen content"))
    info_shift_down_pt <- if (compact) 6 else 10
  if (identical(title, "lnRR Plant nitrogen content VS. Phenolics-Phenolics"))
    info_shift_down_pt <- if (compact) 13.25 else 24
  if (identical(title, "lnRR Plant carbon content VS. Cellulose-Plant carbon content"))
    info_shift_down_pt <- 20
  info_y <- usr[4] - 0.05 * (usr[4] - usr[3])
  info_y <- y_user(graphics::grconvertY(info_y, "user", "inches") - info_shift_down_pt / 72)
  graphics::text(usr[2] - 0.02 * (usr[2] - usr[1]),
    info_y, labels = info,
    adj = c(1, 1), cex = annotation_pt / 12)
}
manuscript_funnel_par <- function(nrow = 1L, ncol = 1L, compact = FALSE) {
  graphics::par(mfrow = c(nrow, ncol))
  cell_scale <- graphics::par("fin")[1] / 4
  # Physical margins prevent mfrow's implicit .66 text shrinkage and keep
  # the complete y-axis label inside every single, paired or combined page.
  graphics::par(cex = 1, ps = 12, family = "sans", oma = rep(0, 4),
    mai = c(0.425, 0.57, 0.255, 0.06) * cell_scale,
    mgp = c(2, 0.45 * cell_scale, 0), tcl = -0.34 * cell_scale,
    lwd = 0.7 * cell_scale, las = 0)
}

draw_f5_publication_funnel <- function(funnel_object, compact=FALSE) {
 draw_manuscript_funnel(funnel_object$model, funnel_object$title,
   funnel_object$egger_p, funnel_object$begg_p, funnel_object$trimfill_k0,
   funnel_object$n_effect_sizes, funnel_object$n_studies, compact, p_digits=3L)
}
save_f5_funnel_collection <- function(funnel_objects,stem,directory,nrow,ncol,width,height,compact=FALSE) {
 if(!length(funnel_objects)) return(invisible(NULL))
 if(length(funnel_objects)>nrow*ncol) stop("Funnel layout has fewer slots than plots: ",stem)
 draw <- function() {
   manuscript_funnel_par(nrow,ncol,compact)
   for(obj in funnel_objects) draw_f5_publication_funnel(obj,compact)
   if(length(funnel_objects)<nrow*ncol)
     for(i in seq_len(nrow*ncol-length(funnel_objects))) graphics::plot.new()
 }
 for(type in c("pdf","tiff")) {
   path <- file.path(directory,paste0(stem,".",type));assert_output_path(path)
   open <- local({tp<-type;function(temp) {
     if(tp=="pdf") grDevices::pdf(temp,width=width,height=height,family="sans")
     else grDevices::tiff(temp,width=width,height=height,units="in",res=tiff_dpi,compression="lzw")
   }})
   safe_plot_device(path,open,draw)
 }
 invisible(NULL)
}

# Stable output names are checked before any model fitting.
if(!is.null(output_root_override)) out_dir <- output_root_override
ensure_output_dir(out_dir)
out_dir <- normalizePath(out_dir,winslash="/",mustWork=TRUE)
publication_dir <- out_dir
funnel_dir <- file.path(out_dir,"F5_funnel")
main_pdf <- c(file.path("F4",paste0(c("F4",paste0("F4",LETTERS[1:4])),".pdf")),
 file.path("F5",paste0(c("F5",paste0("F5",LETTERS[1:6])),".pdf")))
funnel_stems <- c("F5A_C_funnel","F5A_Cellulose_funnel","F5B_C_funnel","F5B_LMA_funnel",
 "F5C_C_funnel","F5C_Terpenoids_funnel","F5D_N_funnel","F5D_Alkaloids_funnel",
 "F5E_N_funnel","F5E_Flavonoids_funnel","F5F_N_funnel","F5F_Terpenoids_funnel",paste0("F5",LETTERS[1:6],"_funnels"),"SF6")
expected_pdf <- c(main_pdf,file.path("F5_funnel",paste0(funnel_stems,".pdf")))
expected_files <- c(expected_pdf,sub("[.]pdf$",".tiff",expected_pdf),
 "F4/F4_results.xlsx","F4/F4_model_summary.csv","F5/F5_results.xlsx","F5/F5_model_summary.csv",
 "F5/Table_S3.csv","F5_funnel/F5_funnel_results.xlsx","F5_funnel/F5_funnel_summary.csv","panel_manifest.csv")
info_path <- file.path(out_dir,"run_info.txt")
input_md5 <- unname(tools::md5sum(input_file))
started_at <- format(Sys.time(),"%Y-%m-%d %H:%M:%S %Z")
captured_warnings <- character()
current_stage <- "output preflight"
message("Input: ",input_file)
message("Output: ",out_dir)
tryCatch(withCallingHandlers({
 check_existing_output(info_path)
 safe_write_lines(c("RUNNING",paste0("Started: ",started_at),paste0("Input: ",input_file),
 "Use results only after this file reports COMPLETE for the current attempt."),info_path)
 for(folder in c("F4","F5","F5_funnel")) probe_output_dir(file.path(out_dir,folder),45L)
 for(path in expected_files) check_existing_output(file.path(out_dir,path))
primary_method <- "M2_signed_sqrt_trait_RMS"
methods_to_run <- primary_method
## Only M2 is fitted in this publication program. RAW is used for all displayed
## figures; IQR-removed M2 results remain in the tables as sensitivity evidence.
run_iqr_sensitivity <- TRUE
iqr_k <- 1.5
eps <- 1e-12

use_for_regression_keep <- c("main")
pairing_level_keep <- c("exact", "baseline")
data_level_keep <- c("species_pair", "treatment_level", "population_pair")
tissue_keep <- c("leaf")

## Figure 5 scan follows the original F3/F5 framework:
## C -> every observed defense trait; N -> every observed chemical-defense trait.
## Set TRUE only if a separate exploratory P-specific scan is later required.
include_p_specific_f5 <- FALSE

## The four traits used in the submitted Figure 5 are exported again after the
## correction for direct comparison, but they are not automatically declared
## significant or retained as final panels.
original_f5_traits <- c("Cellulose", "Glycosides", "Flavonoids", "Terpenoids")

## Final Figure 5 panels requested after inspection of the corrected M2-RAW
## candidate screen. Their shapes also agree with the implemented automatic
## selection rule in the current corrected dataset. The explicit table below
## makes the final figure reproducible and prevents accidental panel reordering.
f5_publication_plan <- tibble::tribble(
  ~Panel, ~Nutrient, ~DefenseTrait, ~RequestedModel,
  "A", "C", "Cellulose",  "linear",
  "B", "C", "LMA",        "linear",
  "C", "C", "Terpenoids", "quadratic",
  "D", "N", "Alkaloids",  "quadratic",
  "E", "N", "Flavonoids", "quadratic",
  "F", "N", "Terpenoids", "quadratic"
)

## Model-shape rule implemented by the supplied analysis.
min_n_linear <- 3
min_n_quadratic <- 6
quadratic_min_delta_aicc <- 2
quadratic_min_delta_r2 <- 0.03
quadratic_max_p_ordinary <- 0.05
quadratic_min_systems <- 5
quadratic_vertex_quantiles <- c(0.05, 0.95)

## Evidence classification. These labels organize output; they do not delete results.
publication_min_n <- 10
publication_min_studies <- 3
publication_min_systems <- 5
publication_min_r2 <- 0.10
publication_alpha_cr2_fdr <- 0.05
publication_alpha_balanced_cr2 <- 0.10
supportive_alpha_ordinary_fdr <- 0.10

## Figure style: retained from the original program.
pdf_w <- 6.8
pdf_h <- 5.6
combined_pdf_w <- 10.4
combined_pdf_h <- 8.0
f5_grid_pdf_w <- 15.6
f5_grid_pdf_h <- 9.0
## Wider canvas used only for the final 2-column x 3-row Figure 5 grid,
## preventing long Cha axis titles from being clipped in the combined figure.
f5_final_grid_pdf_w <- 18.0
base_font_size <- 18
axis_title_size <- base_font_size + 1
axis_text_size <- base_font_size - 1
axis_line_width <- 0.55
axis_tick_width <- 0.50
axis_tick_length_cm <- 0.13
panel_tag_size <- 20
point_size_range <- c(2.2, 8.0)
tiff_dpi <- 600
confidence_level <- 0.95
confidence_band_method <- "model_based_WLS_fitted_mean"
confidence_band_alpha <- 0.50
## Non-publication candidate figures are deliberately not written. Their full
## numerical model results remain available in the workbook and CSV tables.
export_diagnostic_figures <- FALSE

## Figure 5 supplementary funnel-plot settings. These follow the old F5/S6
## organization: every selected relationship contributes one nutrient lnRR
## funnel and one defense-trait lnRR funnel from the exact RowID set used by
## that regression. Cha is used only in the regression, not in the funnels.
funnel_level <- c(90, 95, 99)
funnel_shade <- c("white", "grey90", "grey80")
funnel_min_effects_for_bias_tests <- 10
funnel_run_begg_ranktest <- TRUE
funnel_run_trimfill <- TRUE
funnel_individual_pdf_w <- 6.4
funnel_individual_pdf_h <- 5.2
funnel_pair_pdf_w <- 12.8
funnel_pair_pdf_h <- 5.2
funnel_grid_pdf_w <- 13.8
funnel_grid_pdf_h <- 17.2
## Publication Figure 5 panel tags are placed inside the upper-left safe area
## of each forest plot in the six-panel main figure only. Standalone Figure 5
## panels are exported without A-F tags. Negative coordinates are deliberately
## avoided because A/C/E were clipped at the outer boundary of the six-panel grid.
f5_panel_tag_position <- c(0.025, 0.985)
## Figure 5 companion-forest axes are narrow inside the six-panel grid.
## Limit them to at most three non-overlapping labels; this changes only the
## displayed ticks, never the effect estimates, CIs, zero line, or x limits.
f5_forest_max_x_breaks <- 3L
f5_forest_axis_text_size <- base_font_size - 4
f5_forest_axis_title_size <- base_font_size - 3
## Device-safe sans-serif family matches the supplied manuscript figures.
plot_font_family <- "sans"

# =========================
# 2) Text and trait helpers
# =========================
clean_key <- function(x) {
  x <- stringr::str_trim(as.character(x))
  x <- stringr::str_replace_all(x, "\\s+", "_")
  x <- stringr::str_replace_all(x, "[^A-Za-z0-9_\\-\\.]+", "-")
  stringr::str_replace_all(x, "-+", "-")
}

std_text <- function(x) {
  x <- stringr::str_trim(as.character(x))
  x <- stringr::str_replace_all(x, "\\s+", "_")
  tolower(x)
}

std_trait <- function(x) {
  x0 <- stringr::str_trim(as.character(x))
  x1 <- tolower(x0)
  x1 <- stringr::str_replace_all(x1, "[_\\-]+", " ")
  x1 <- stringr::str_squish(x1)
  dplyr::case_when(
    x1 %in% c("c", "carbon", "leaf c", "plant c", "plant carbon", "total c") ~ "C",
    x1 %in% c("n", "nitrogen", "leaf n", "plant n", "plant nitrogen", "total n") ~ "N",
    x1 %in% c("p", "phosphorus", "phosphorous", "leaf p", "plant p", "plant phosphorus", "total p") ~ "P",
    x1 %in% c("phenolics", "phenolic", "total phenolics", "total phenols", "phenols", "polyphenols") ~ "Phenolics",
    x1 %in% c("tannins", "tannin", "condensed tannins", "total tannins") ~ "Tannins",
    x1 %in% c("flavonoids", "flavonoid", "total flavonoids") ~ "Flavonoids",
    x1 %in% c("terpenoids", "terpene", "terpenes", "total terpenes", "total terpenoids") ~ "Terpenoids",
    x1 %in% c("alkaloids", "alkaloid", "total alkaloids") ~ "Alkaloids",
    x1 %in% c("glucosinolates", "glucosinolate") ~ "Glucosinolates",
    x1 %in% c("glycosides", "glycoside") ~ "Glycosides",
    x1 %in% c("saponins", "saponin") ~ "Saponins",
    x1 %in% c("proteinaceous toxins", "protein toxins", "protease inhibitors", "proteinase inhibitors") ~ "Proteinaceous toxins",
    x1 %in% c("cellulose") ~ "Cellulose",
    x1 %in% c("hemicellulose", "hemi cellulose") ~ "Hemicellulose",
    x1 %in% c("lignin") ~ "Lignin",
    x1 %in% c("wax", "waxes", "leaf wax") ~ "Wax",
    x1 %in% c("resin", "resins") ~ "Resin",
    x1 %in% c("latex") ~ "Latex",
    x1 %in% c("leaf toughness", "toughness") ~ "Leaf toughness",
    x1 %in% c("leaf thickness", "thickness") ~ "Leaf thickness",
    x1 %in% c("lma", "leaf mass per area", "leaf mass area") ~ "LMA",
    x1 %in% c("sla", "specific leaf area") ~ "SLA",
    x1 %in% c("trichome", "trichomes", "trichome density", "trichome number") ~ "Trichome density",
    TRUE ~ x0
  )
}

safe_file_label <- function(x) {
  x <- stringr::str_replace_all(as.character(x), "[^A-Za-z0-9_.-]+", "_")
  stringr::str_replace_all(x, "_+", "_")
}

## A conservative limit leaves room for device- and Windows-internal suffixes.
## This check runs immediately before every PDF/TIFF device is opened.
windows_safe_path_limit <- 240L
assert_output_path <- function(path) {
  path <- normalizePath(path, winslash = "/", mustWork = FALSE)
  parent <- dirname(path)
  if (!dir.exists(parent)) {
    dir.create(parent, recursive = TRUE, showWarnings = FALSE)
  }
  if (!dir.exists(parent)) {
    stop("Cannot create output directory: ", parent)
  }
  n_path <- nchar(path, type = "chars")
  if (.Platform$OS.type == "windows" && n_path > windows_safe_path_limit) {
    stop(
      "Output path is too long for a reliable Windows graphics device (",
      n_path, " characters; limit = ", windows_safe_path_limit, "): ", path,
      ". Set output_root_override to a shorter writable directory."
    )
  }
  invisible(path)
}

safe_rms <- function(x) {
  x <- as.numeric(x)
  s <- sqrt(mean(x[is.finite(x)]^2, na.rm = TRUE))
  if (!is.finite(s) || s <= 1e-10) 1 else s
}

weighted_mean_var <- function(value, variance) {
  value <- as.numeric(value)
  variance <- as.numeric(variance)
  ok <- is.finite(value) & is.finite(variance) & variance >= 0
  if (!any(ok)) return(tibble::tibble(value = NA_real_, variance = NA_real_))
  w <- 1 / pmax(variance[ok], eps)
  tibble::tibble(
    value = sum(w * value[ok]) / sum(w),
    variance = 1 / sum(w)
  )
}

# =========================
# 3) Correct transformations
# =========================
signed_sqrt_transform <- function(d, variance) {
  d <- as.numeric(d)
  variance <- as.numeric(variance)
  ## The signed square root is non-differentiable at exactly zero.
  ## The small floor makes the delta-method variance very large at zero,
  ## preventing an exact zero difference from receiving artificial high weight.
  derivative_floor <- 1e-8
  z <- sign(d) * sqrt(abs(d))
  derivative <- 0.5 / sqrt(pmax(abs(d), derivative_floor))
  tibble::tibble(z = z, variance = derivative^2 * variance)
}

transform_pair <- function(df, method) {
  tx <- signed_sqrt_transform(df$X_value_raw, df$Var_X_raw)
  ty <- signed_sqrt_transform(df$Y_value_raw, df$Var_Y_raw)
  out <- df
  out$zX <- tx$z
  out$Var_zX <- tx$variance
  out$zY <- ty$z
  out$vi <- ty$variance

  if (method == "M2_signed_sqrt_trait_RMS") {
    ## X occurs once for every paired Y row. Estimate its scale from unique
    ## biological X comparisons so systems with more defense measurements do
    ## not influence the axis scale more strongly merely through duplication.
    x_unique <- out %>%
      dplyr::distinct(StudyID, SystemID, .keep_all = TRUE)
    sx <- safe_rms(x_unique$zX)
    out$zX <- out$zX / sx
    out$Var_zX <- out$Var_zX / sx^2
    out <- out %>%
      dplyr::group_by(YTrait) %>%
      dplyr::group_modify(~ {
        sy <- safe_rms(.x$zY)
        .x$zY <- .x$zY / sy
        .x$vi <- .x$vi / sy^2
        .x
      }) %>%
      dplyr::ungroup()
  } else if (!(method %in% c("M1_signed_sqrt_Cha", "M3_symmetric_relative_Cha"))) {
    stop("Unknown corrected method: ", method)
  }

  out %>%
    dplyr::filter(
      is.finite(zX), is.finite(zY),
      is.finite(Var_zX), Var_zX >= 0,
      is.finite(vi), vi > 0
    ) %>%
    dplyr::mutate(w_for_model = 1 / pmax(vi, eps))
}

iqr_filter_2d <- function(df, k = 1.5) {
  if (nrow(df) < 4) return(df)
  qx <- stats::quantile(df$zX, c(.25, .75), na.rm = TRUE)
  qy <- stats::quantile(df$zY, c(.25, .75), na.rm = TRUE)
  ix <- qx[2] - qx[1]
  iy <- qy[2] - qy[1]
  df %>%
    dplyr::filter(
      zX >= qx[1] - k * ix, zX <= qx[2] + k * ix,
      zY >= qy[1] - k * iy, zY <= qy[2] + k * iy
    )
}

# =========================
# 4) Read, normalize and audit input
# =========================
current_stage <- "input import"
raw <- read_f45_input(input_file,sheet_name)
input_header_row <- attr(raw,"source_header_row")
if ("Sdi" %in% names(raw) && !("SDi" %in% names(raw))) {
  names(raw)[names(raw) == "Sdi"] <- "SDi"
}

required_cols <- c(
  "StudyID", "SystemID", "ComparisonType", "Trait", "TraitCategory", "TraitForm",
  "Tissue", "DataLevel", "PairingLevel", "UseForRegression",
  "Mc", "SDc", "Nc", "Mi", "SDi", "Ni"
)
missing_cols <- setdiff(required_cols, names(raw))
if (length(missing_cols)) {
  stop("Missing required columns: ", paste(missing_cols, collapse = ", "))
}
if (!("Unit" %in% names(raw))) raw$Unit <- NA_character_

dat_all <- raw %>%
  dplyr::mutate(
    RowID = dplyr::row_number(),
    StudyID = clean_key(StudyID),
    SystemID = clean_key(SystemID),
    ComparisonType = std_text(ComparisonType),
    Trait = std_trait(Trait),
    TraitCategory = std_text(TraitCategory),
    TraitForm = std_text(TraitForm),
    Tissue = std_text(Tissue),
    DataLevel = std_text(DataLevel),
    PairingLevel = std_text(PairingLevel),
    UseForRegression = std_text(UseForRegression),
    Mc = suppressWarnings(as.numeric(Mc)),
    SDc = suppressWarnings(as.numeric(SDc)),
    Nc = suppressWarnings(as.numeric(Nc)),
    Mi = suppressWarnings(as.numeric(Mi)),
    SDi = suppressWarnings(as.numeric(SDi)),
    Ni = suppressWarnings(as.numeric(Ni)),
    TraitCategory = dplyr::if_else(Trait == "Lignin", "physical_defense", TraitCategory),
    d = Mi - Mc,
    var_d = SDi^2 / pmax(Ni, eps) + SDc^2 / pmax(Nc, eps),
    symmetric_relative = dplyr::if_else(
      Mi > 0 & Mc > 0 & (Mi + Mc) > eps,
      2 * (Mi - Mc) / (Mi + Mc),
      NA_real_
    ),
    Var_symmetric_relative = dplyr::if_else(
      Mi > 0 & Mc > 0 & (Mi + Mc) > eps,
      (4 * Mc / (Mi + Mc)^2)^2 * SDi^2 / pmax(Ni, eps) +
        (-4 * Mi / (Mi + Mc)^2)^2 * SDc^2 / pmax(Nc, eps),
      NA_real_
    ),
    lnRR_value = dplyr::if_else(Mi > 0 & Mc > 0, log(Mi / Mc), NA_real_),
    Var_lnRR = dplyr::if_else(
      Mi > 0 & Mc > 0,
      SDi^2 / (pmax(Ni, eps) * Mi^2) + SDc^2 / (pmax(Nc, eps) * Mc^2),
      NA_real_
    ),
    valid_lnrr = is.finite(lnRR_value) & is.finite(Var_lnRR) & Var_lnRR >= 0
  )

input_audit <- tibble::tibble(
  Item = c(
    "input_file", "input_md5", "input_rows", "comparison_type", "direction", "M2",
    "RAW_role", "IQR_role", "F5_scan", "displayed_outputs", "legacy_R2_check"
  ),
  Value = c(
    input_file, unname(tools::md5sum(input_file)), nrow(dat_all), "introduced_vs_native_range",
    "Mi - Mc = introduced range - native range",
    "sign(d)*sqrt(abs(d)), then positive no-centering RMS scaling",
    "main analysis", "sensitivity only",
    "all observed C-defense and N-chemical-defense paired traits",
    "M2-RAW Figure 4, six M2-RAW Figure 5 panels and matching lnRR funnel plots",
    "diagnostic only; never stops corrected analysis"
  )
)

dat <- dat_all %>%
  dplyr::filter(
    ComparisonType == "introduced_vs_native_range",
    UseForRegression %in% std_text(use_for_regression_keep),
    PairingLevel %in% std_text(pairing_level_keep),
    DataLevel %in% std_text(data_level_keep),
    Tissue %in% std_text(tissue_keep),
    is.finite(Mi), is.finite(Mc), is.finite(SDi), is.finite(SDc),
    is.finite(Ni), is.finite(Nc), Ni > 0, Nc > 0, SDi >= 0, SDc >= 0,
    is.finite(d), is.finite(var_d), var_d >= 0
  )

if (nrow(dat) == 0) stop("No rows remain after the predefined filters.")

system_conflicts <- dat %>%
  dplyr::distinct(SystemID, StudyID) %>%
  dplyr::count(SystemID, name = "n_studies") %>%
  dplyr::filter(n_studies > 1)
if (nrow(system_conflicts)) {
  stop("At least one SystemID occurs in more than one StudyID; pairing would be ambiguous.")
}

exact_duplicates <- dat %>%
  dplyr::count(StudyID, SystemID, Trait, Mi, Mc, SDi, SDc, Ni, Nc, name = "n") %>%
  dplyr::filter(n > 1)
if (nrow(exact_duplicates)) {
  stop("Exact duplicated input effect-size rows were detected. Resolve them before fitting.")
}

## Non-destructive data-quality screen. Rows are only reported, never removed.
## It is designed to catch order-of-magnitude entries such as an N value entered
## on the C scale while allowing the user to verify legitimate biological cases.
input_qc_flags <- dat %>%
  dplyr::mutate(PairMean = (Mi + Mc) / 2) %>%
  dplyr::group_by(Trait, Unit) %>%
  dplyr::mutate(
    GroupN = dplyr::n(),
    GroupMedian = stats::median(PairMean[is.finite(PairMean) & PairMean > 0], na.rm = TRUE),
    FoldFromMedian = dplyr::if_else(
      is.finite(GroupMedian) & GroupMedian > 0 & PairMean > 0,
      pmax(PairMean / GroupMedian, GroupMedian / PairMean),
      NA_real_
    )
  ) %>%
  dplyr::ungroup() %>%
  dplyr::filter(GroupN >= 5, is.finite(FoldFromMedian), FoldFromMedian >= 10) %>%
  dplyr::select(
    RowID, StudyID, SystemID, Trait, Unit, Mc, Mi, GroupN,
    GroupMedian, FoldFromMedian
  )

# =========================
# 5) Pair construction
# =========================
method_metric <- function(method) {
  if (method == "M3_symmetric_relative_Cha") "symmetric_relative" else "raw_difference"
}

build_pair_data <- function(
    x_trait, y_category = NA_character_, y_trait = NA_character_,
    metric = c("raw_difference", "symmetric_relative")) {
  metric <- match.arg(metric)
  value_col <- if (metric == "raw_difference") "d" else "symmetric_relative"
  variance_col <- if (metric == "raw_difference") "var_d" else "Var_symmetric_relative"

  xdat <- dat %>%
    dplyr::filter(Trait == x_trait) %>%
    dplyr::group_by(StudyID, SystemID) %>%
    dplyr::group_modify(~ {
      ## Explicit .x[[...]] access avoids tidy-evaluation failures such as
      ## "object 'variance_col' not found" on different dplyr versions.
      wm <- weighted_mean_var(.x[[value_col]], .x[[variance_col]])
      tibble::tibble(
        X_value_raw = wm$value,
        Var_X_raw = wm$variance,
        X_RowIDs = paste(.x$RowID, collapse = ";")
      )
    }) %>%
    dplyr::ungroup()

  ydat <- dat %>% dplyr::filter(!(Trait %in% c("C", "N", "P")))
  if (!is.na(y_category) && nzchar(y_category)) {
    ydat <- ydat %>% dplyr::filter(TraitCategory == y_category)
  }
  if (!is.na(y_trait) && nzchar(y_trait)) {
    ydat <- ydat %>% dplyr::filter(Trait == y_trait)
  }

  ydat %>%
    dplyr::transmute(
      StudyID, SystemID, Y_RowID = RowID,
      YTrait = Trait, YCategory = TraitCategory, YTraitForm = TraitForm, YUnit = Unit,
      Y_value_raw = .data[[value_col]], Var_Y_raw = .data[[variance_col]]
    ) %>%
    dplyr::inner_join(xdat, by = c("StudyID", "SystemID")) %>%
    dplyr::filter(
      is.finite(X_value_raw), is.finite(Var_X_raw), Var_X_raw >= 0,
      is.finite(Y_value_raw), is.finite(Var_Y_raw), Var_Y_raw >= 0
    )
}

# =========================
# 6) WLS, CR2, AICc and LOSO
# =========================
extract_cr2_p <- function(test_table, term) {
  if (is.null(test_table) || !nrow(test_table)) return(NA_real_)
  rn <- rownames(test_table)
  idx <- which(rn == term)
  if (!length(idx) && "Coef" %in% names(test_table)) {
    idx <- which(as.character(test_table$Coef) == term)
  }
  pcol <- grep("^p", names(test_table), value = TRUE)[1]
  if (!length(idx) || length(pcol) == 0 || is.na(pcol)) return(NA_real_)
  as.numeric(test_table[[pcol]][idx[1]])
}

weighted_r2 <- function(y, fitted, w) {
  ybar <- sum(w * y) / sum(w)
  sst <- sum(w * (y - ybar)^2)
  if (!is.finite(sst) || sst <= 0) return(NA_real_)
  1 - sum(w * (y - fitted)^2) / sst
}

fit_wls <- function(
    df, form = c("linear", "quadratic"), weight_col = "w_for_model",
    trait_adjusted = FALSE) {
  form <- match.arg(form)
  min_n <- ifelse(form == "quadratic", min_n_quadratic, min_n_linear)
  min_unique_x <- ifelse(form == "quadratic", 3, 2)
  if (nrow(df) < min_n || dplyr::n_distinct(df$zX) < min_unique_x) return(NULL)

  use <- df
  use$w_analysis <- as.numeric(use[[weight_col]])
  use <- use %>% dplyr::filter(is.finite(w_analysis), w_analysis > 0)
  if (nrow(use) < min_n) return(NULL)
  if (isTRUE(trait_adjusted) && dplyr::n_distinct(use$YTrait) < 2) {
    trait_adjusted <- FALSE
  }
  formula <- if (isTRUE(trait_adjusted)) {
    if (form == "quadratic") {
      zY ~ zX + I(zX^2) + factor(YTrait)
    } else {
      zY ~ zX + factor(YTrait)
    }
  } else {
    if (form == "quadratic") zY ~ zX + I(zX^2) else zY ~ zX
  }
  fit <- stats::lm(formula, data = use, weights = w_analysis, model = TRUE, x = TRUE, y = TRUE)
  sm <- summary(fit)
  coef_names <- names(stats::coef(fit))
  if (any(!is.finite(stats::coef(fit)))) return(NULL)

  cr2 <- NULL
  cr2_vcov <- NULL
  cr2_status <- "not_run_fewer_than_2_StudyID_clusters"
  if (dplyr::n_distinct(use$StudyID) >= 2) {
    cr2_status <- "available"
    cr2_bundle <- tryCatch({
      V <- clubSandwich::vcovCR(
        fit,
        cluster = as.factor(use$StudyID),
        type = "CR2"
      )
      list(
        V = V,
        test = clubSandwich::coef_test(fit, vcov = V, test = "Satterthwaite")
      )
    }, error = function(e) {
      cr2_status <<- paste0("failed: ", conditionMessage(e))
      NULL
    })
    if (!is.null(cr2_bundle)) {
      cr2 <- cr2_bundle$test
      cr2_vcov <- cr2_bundle$V
    }
  }

  ## logLik(lm) counts both regression coefficients and the fitted scale
  ## parameter. Using its df keeps the AICc correction consistent with AIC().
  kpar <- attr(stats::logLik(fit), "df")
  aic <- stats::AIC(fit)
  aicc <- if (nrow(use) > kpar + 1) {
    aic + 2 * kpar * (kpar + 1) / (nrow(use) - kpar - 1)
  } else NA_real_

  r2_total <- weighted_r2(use$zY, stats::fitted(fit), use$w_analysis)
  model_vcov <- stats::vcov(fit)
  model_vcov_symmetric <- (model_vcov + t(model_vcov)) / 2
  model_vcov_eigen <- eigen(
    model_vcov_symmetric, symmetric = TRUE, only.values = TRUE
  )$values
  model_vcov_tolerance <- 1e-10 * max(1, max(abs(model_vcov_eigen), na.rm = TRUE))
  if (any(!is.finite(model_vcov_eigen)) ||
      min(model_vcov_eigen) < -model_vcov_tolerance) {
    stop("Model-based WLS covariance matrix is not positive semidefinite.")
  }
  if (isTRUE(trait_adjusted)) {
    baseline_fit <- stats::lm(
      zY ~ factor(YTrait), data = use, weights = w_analysis,
      model = TRUE, x = TRUE, y = TRUE
    )
    sse_full <- sum(use$w_analysis * stats::residuals(fit)^2)
    sse_baseline <- sum(use$w_analysis * stats::residuals(baseline_fit)^2)
    partial_r2 <- if (is.finite(sse_baseline) && sse_baseline > 0) {
      1 - sse_full / sse_baseline
    } else NA_real_
  } else {
    partial_r2 <- r2_total
  }

  list(
    fit = fit,
    data = use,
    form = form,
    b0 = unname(stats::coef(fit)[1]),
    b1 = if ("zX" %in% coef_names) unname(stats::coef(fit)["zX"]) else NA_real_,
    b2 = if ("I(zX^2)" %in% coef_names) unname(stats::coef(fit)["I(zX^2)"]) else NA_real_,
    p1 = if ("zX" %in% rownames(sm$coefficients)) sm$coefficients["zX", 4] else NA_real_,
    p2 = if ("I(zX^2)" %in% rownames(sm$coefficients)) sm$coefficients["I(zX^2)", 4] else NA_real_,
    p1_CR2 = extract_cr2_p(cr2, "zX"),
    p2_CR2 = extract_cr2_p(cr2, "I(zX^2)"),
    R2 = r2_total,
    Partial_R2 = partial_r2,
    AICc = aicc,
    TraitAdjusted = isTRUE(trait_adjusted),
    CR2_available = !is.null(cr2),
    CR2_status = cr2_status,
    V_CR2 = cr2_vcov,
    V_model = model_vcov_symmetric,
    CI_df_model = max(1, fit$df.residual),
    ConfidenceBandMethod = confidence_band_method
  )
}

loso_cv_r2 <- function(df, form, weight_col = "w_for_model", trait_adjusted = FALSE) {
  studies <- unique(df$StudyID)
  if (length(studies) < 3) return(NA_real_)
  sse_model <- 0
  sse_null <- 0
  used <- 0
  for (study in studies) {
    train <- df %>% dplyr::filter(StudyID != study)
    test <- df %>% dplyr::filter(StudyID == study)
    min_unique_x <- ifelse(form == "quadratic", 3, 2)
    min_n <- ifelse(form == "quadratic", min_n_quadratic, min_n_linear)
    if (nrow(train) < min_n || dplyr::n_distinct(train$zX) < min_unique_x) return(NA_real_)
    if (isTRUE(trait_adjusted) &&
        !all(unique(test$YTrait) %in% unique(train$YTrait))) return(NA_real_)
    formula <- if (isTRUE(trait_adjusted)) {
      if (form == "quadratic") {
        zY ~ zX + I(zX^2) + factor(YTrait)
      } else {
        zY ~ zX + factor(YTrait)
      }
    } else {
      if (form == "quadratic") zY ~ zX + I(zX^2) else zY ~ zX
    }
    train$w_cv <- as.numeric(train[[weight_col]])
    fit <- tryCatch(
      stats::lm(formula, data = train, weights = w_cv),
      error = function(e) NULL
    )
    if (is.null(fit)) return(NA_real_)
    pred <- tryCatch(stats::predict(fit, newdata = test), error = function(e) rep(NA_real_, nrow(test)))
    if (any(!is.finite(pred))) return(NA_real_)
    w_test <- test[[weight_col]]
    w_train <- train[[weight_col]]
    null_pred <- sum(w_train * train$zY) / sum(w_train)
    sse_model <- sse_model + sum(w_test * (test$zY - pred)^2)
    sse_null <- sse_null + sum(w_test * (test$zY - null_pred)^2)
    used <- used + nrow(test)
  }
  if (used == 0 || !is.finite(sse_null) || sse_null <= 0) return(NA_real_)
  1 - sse_model / sse_null
}

result_row <- function(df, result, method, dataset, nutrient, defense_trait, defense_category, trait_form, model, sensitivity = "original_inverse_variance") {
  selected_p <- if (model == "quadratic") result$p2 else result$p1
  selected_p_cr2 <- if (model == "quadratic") result$p2_CR2 else result$p1_CR2
  selected_term <- if (model == "quadratic") result$b2 else result$b1
  wcol <- if (sensitivity == "original_inverse_variance") "w_for_model" else "w_analysis"
  weights_now <- result$data[[wcol]]
  study_share <- tapply(weights_now, result$data$StudyID, sum)
  tibble::tibble(
    Method = method,
    Dataset = dataset,
    Nutrient = nutrient,
    DefenseTrait = defense_trait,
    DefenseCategory = defense_category,
    TraitForm = trait_form,
    Model = model,
    Sensitivity = sensitivity,
    n = nrow(result$data),
    Studies = dplyr::n_distinct(result$data$StudyID),
    Systems = dplyr::n_distinct(result$data$SystemID),
    R2_weighted = result$R2,
    Partial_R2_nutrient = result$Partial_R2,
    TraitAdjusted = result$TraitAdjusted,
    LOSO_CV_R2 = loso_cv_r2(result$data, model, wcol, result$TraitAdjusted),
    AICc = result$AICc,
    b0 = result$b0,
    b1 = result$b1,
    b2 = result$b2,
    Selected_term = selected_term,
    P_selected_term_ordinary = selected_p,
    P_selected_term_CR2 = selected_p_cr2,
    CR2_available = result$CR2_available,
    CR2_status = result$CR2_status,
    Displayed_CI = result$ConfidenceBandMethod,
    Max_study_weight_share = max(study_share) / sum(study_share)
  )
}

choose_model_shape <- function(linear, quadratic, df) {
  if (is.null(quadratic)) {
    return(list(
      selected = "linear", delta_aicc = NA_real_, delta_r2 = NA_real_,
      vertex = NA_real_, vertex_inside = FALSE, quadratic_eligible = FALSE
    ))
  }
  delta_aicc <- linear$AICc - quadratic$AICc
  delta_r2 <- quadratic$Partial_R2 - linear$Partial_R2
  vertex <- if (is.finite(quadratic$b2) && quadratic$b2 != 0) {
    -quadratic$b1 / (2 * quadratic$b2)
  } else NA_real_
  vertex_limits <- stats::quantile(
    df$zX, probs = quadratic_vertex_quantiles, na.rm = TRUE, names = FALSE
  )
  vertex_inside <- is.finite(vertex) &&
    vertex >= vertex_limits[1] && vertex <= vertex_limits[2]
  eligible <- nrow(quadratic$data) >= publication_min_n &&
    dplyr::n_distinct(quadratic$data$SystemID) >= quadratic_min_systems &&
    is.finite(delta_aicc) && delta_aicc >= quadratic_min_delta_aicc &&
    is.finite(delta_r2) && delta_r2 >= quadratic_min_delta_r2 &&
    is.finite(quadratic$p2) && quadratic$p2 <= quadratic_max_p_ordinary &&
    vertex_inside
  list(
    selected = ifelse(eligible, "quadratic", "linear"),
    delta_aicc = delta_aicc, delta_r2 = delta_r2,
    vertex = vertex, vertex_inside = vertex_inside,
    quadratic_eligible = eligible
  )
}

# =========================
# 7) Plotting
# =========================
nutrient_axis_label <- function(nutrient) {
  dplyr::recode(
    as.character(nutrient),
    C = "Plant carbon content",
    N = "Plant nitrogen content",
    P = "Plant phosphorus content",
    .default = as.character(nutrient)
  )
}

make_prediction_band <- function(df, result) {
  ## Use a strictly increasing prediction grid across the observed RAW range.
  ## No point is removed and the curve is never extrapolated beyond the data.
  xgrid <- sort(unique(seq(min(df$zX), max(df$zX), length.out = 320)))
  if (isTRUE(result$TraitAdjusted)) {
    trait_levels <- sort(unique(as.character(df$YTrait)))
    newdat <- tidyr::expand_grid(zX = xgrid, YTrait = trait_levels)
    newdat$YTrait <- factor(newdat$YTrait, levels = trait_levels)
    mm <- stats::model.matrix(
      stats::delete.response(stats::terms(result$fit)), newdat,
      contrasts.arg = result$fit$contrasts, xlev = result$fit$xlevels
    )
    mm <- mm[, names(stats::coef(result$fit)), drop = FALSE]
    average_rows <- lapply(seq_along(xgrid), function(ii) {
      colMeans(mm[newdat$zX == xgrid[ii], , drop = FALSE])
    })
    average_mm <- do.call(rbind, average_rows)
    fit_value <- as.numeric(average_mm %*% stats::coef(result$fit))
    vc <- result$V_model
    pred_variance <- rowSums((average_mm %*% vc) * average_mm)
  } else {
    newdat <- data.frame(zX = xgrid)
    mm <- stats::model.matrix(
      stats::delete.response(stats::terms(result$fit)), newdat,
      contrasts.arg = result$fit$contrasts, xlev = result$fit$xlevels
    )
    mm <- mm[, names(stats::coef(result$fit)), drop = FALSE]
    fit_value <- as.numeric(mm %*% stats::coef(result$fit))
    vc <- result$V_model
    pred_variance <- rowSums((mm %*% vc) * mm)
  }
  variance_tolerance <- 1e-10 * max(1, max(abs(pred_variance), na.rm = TRUE))
  if (any(!is.finite(pred_variance)) ||
      min(pred_variance) < -variance_tolerance) {
    stop("Invalid model-based prediction variance; confidence band was not drawn.")
  }
  se_value <- sqrt(pmax(pred_variance, 0))
  alpha_tail <- (1 - confidence_level) / 2
  tcrit <- stats::qt(1 - alpha_tail, df = result$CI_df_model)
  band <- tibble::tibble(
    zX = xgrid,
    fit = fit_value,
    lwr = fit_value - tcrit * se_value,
    upr = fit_value + tcrit * se_value,
    se_fit = se_value
  )
  if (any(!is.finite(as.matrix(band))) ||
      any(diff(band$zX) <= 0) || any(band$lwr > band$upr)) {
    stop("Confidence-band self-check failed.")
  }
  band
}

confidence_band_audit_row <- function(
    df, result, figure, panel, nutrient, response_trait, model) {
  band <- make_prediction_band(df, result)
  widths <- band$upr - band$lwr
  tibble::tibble(
    Figure = figure,
    Panel = panel,
    Nutrient = nutrient,
    Response = response_trait,
    Model = model,
    n = nrow(df),
    Studies = dplyr::n_distinct(df$StudyID),
    Systems = dplyr::n_distinct(df$SystemID),
    ConfidenceLevel = confidence_level,
    ConfidenceBandMethod = confidence_band_method,
    GridPoints = nrow(band),
    X_min = min(band$zX),
    X_max = max(band$zX),
    MinBandWidth = min(widths),
    MaxBandWidth = max(widths),
    FiniteAndOrdered = all(is.finite(as.matrix(band))) &&
      all(diff(band$zX) > 0) && all(band$lwr <= band$upr)
  )
}

build_plot <- function(df, result, method, nutrient, response_label, panel_tag = NULL) {
  band <- make_prediction_band(df, result)
  r2_text <- if (isTRUE(result$TraitAdjusted)) {
    sprintf("Partial R² = %.3f", result$Partial_R2)
  } else {
    sprintf("R² = %.3f", result$R2)
  }
  p <- ggplot2::ggplot(df, ggplot2::aes(zX, zY)) +
    ggplot2::geom_hline(yintercept = 0, color = "grey82", linewidth = .45, linetype = "dashed") +
    ggplot2::geom_vline(xintercept = 0, color = "grey82", linewidth = .45, linetype = "dashed") +
    ggplot2::geom_ribbon(
      data = band, ggplot2::aes(x = zX, ymin = lwr, ymax = upr),
      inherit.aes = FALSE, fill = "grey80", alpha = confidence_band_alpha,
      color = NA
    ) +
    ggplot2::geom_line(
      data = band, ggplot2::aes(x = zX, y = fit), inherit.aes = FALSE,
      color = "#2F77C4", linewidth = .9, linetype = "dashed"
    ) +
    ggplot2::geom_point(
      ggplot2::aes(size = w_for_model), color = "grey32", alpha = .62,
      stroke = base_font_size / 22
    ) +
    ggplot2::scale_size_continuous(range = point_size_range, guide = "none") +
    ggplot2::annotate(
      "text", x = -Inf, y = Inf, hjust = -.15, vjust = 1.35,
      label = r2_text,
      family = plot_font_family, fontface = "bold", size = 4.8
    ) +
    ggplot2::labs(
      x = paste0("Cha-", nutrient_axis_label(nutrient)),
      y = paste0("Cha-", response_label),
      tag = panel_tag
    ) +
    ggplot2::theme_classic(base_size = base_font_size, base_family = plot_font_family) +
    ggplot2::theme(
      text = ggplot2::element_text(family = plot_font_family, color = "black"),
      axis.title = ggplot2::element_text(
        family = plot_font_family, face = "bold", size = axis_title_size
      ),
      axis.text = ggplot2::element_text(
        family = plot_font_family, face = "plain", size = axis_text_size,
        color = "black"
      ),
      axis.line = ggplot2::element_line(
        linewidth = axis_line_width, color = "black"
      ),
      axis.ticks = ggplot2::element_line(
        linewidth = axis_tick_width, color = "black"
      ),
      axis.ticks.length = grid::unit(axis_tick_length_cm, "cm"),
      plot.tag = ggplot2::element_text(
        family = plot_font_family, face = "bold", size = panel_tag_size,
        color = "black"
      ),
      plot.tag.position = c(-.04, 1.02),
      plot.margin = ggplot2::margin(10, 10, 10, 16)
    )
  p
}

# =========================
# 8) Complete Figure 4
# =========================
current_stage <- "F4 models and graphics"
f4_plan <- tibble::tribble(
  ~Panel, ~Nutrient, ~DefenseCategory, ~ResponseLabel,
  "A", "C", "physical_defense", "Physical defense",
  "B", "C", "chemical_defense", "Chemical defense",
  "C", "N", "chemical_defense", "Chemical defense",
  "D", "P", "chemical_defense", "Chemical defense"
)

f4_rows <- list()
f4_all_model_rows <- list()
f4_sensitivity_rows <- list()
f4_points <- list()
f4_plot_store <- list()
f4_fit_store <- list()

for (method in methods_to_run) {
  for (ii in seq_len(nrow(f4_plan))) {
    target <- f4_plan[ii, ]
    pair <- build_pair_data(
      target$Nutrient, y_category = target$DefenseCategory,
      metric = method_metric(method)
    )
    transformed <- transform_pair(pair, method)
    datasets <- list(RAW = transformed)
    if (run_iqr_sensitivity) datasets$IQR_removed <- iqr_filter_2d(transformed, iqr_k)
    for (dataset in names(datasets)) {
      use <- datasets[[dataset]]
      ## Primary Figure 4 estimand: pooled relationship after positive M2
      ## scaling. No defense-trait mean is subtracted from the primary model.
      linear <- fit_wls(use, "linear", trait_adjusted = FALSE)
      quadratic <- fit_wls(use, "quadratic", trait_adjusted = FALSE)
      if (is.null(linear)) stop("Figure 4 linear model failed for panel ", target$Panel, ".")
      decision <- choose_model_shape(linear, quadratic, use)
      result <- if (decision$selected == "quadratic") quadratic else linear
      key <- paste(method, dataset, target$Panel, sep = "__")

      f4_all_model_rows[[paste(key, "linear", sep = "__")]] <- result_row(
        use, linear, method, dataset, target$Nutrient,
        target$ResponseLabel, target$DefenseCategory, "pooled_traits", "linear"
      ) %>% dplyr::mutate(Panel = target$Panel, ModelRole = "pooled_primary_candidate", .before = 1)
      if (!is.null(quadratic)) {
        f4_all_model_rows[[paste(key, "quadratic", sep = "__")]] <- result_row(
          use, quadratic, method, dataset, target$Nutrient,
          target$ResponseLabel, target$DefenseCategory, "pooled_traits", "quadratic"
        ) %>% dplyr::mutate(Panel = target$Panel, ModelRole = "pooled_primary_candidate", .before = 1)
      }

      ## Defense-trait fixed effects answer a different, within-trait-adjusted
      ## question and are therefore retained only as sensitivity analyses.
      adjusted_linear <- fit_wls(use, "linear", trait_adjusted = TRUE)
      adjusted_quadratic <- fit_wls(use, "quadratic", trait_adjusted = TRUE)
      if (!is.null(adjusted_linear)) {
        f4_all_model_rows[[paste(key, "trait_adjusted_linear", sep = "__")]] <- result_row(
          use, adjusted_linear, method, dataset, target$Nutrient,
          target$ResponseLabel, target$DefenseCategory, "pooled_traits", "linear"
        ) %>% dplyr::mutate(Panel = target$Panel, ModelRole = "trait_adjusted_sensitivity", .before = 1)
      }
      if (!is.null(adjusted_quadratic)) {
        f4_all_model_rows[[paste(key, "trait_adjusted_quadratic", sep = "__")]] <- result_row(
          use, adjusted_quadratic, method, dataset, target$Nutrient,
          target$ResponseLabel, target$DefenseCategory, "pooled_traits", "quadratic"
        ) %>% dplyr::mutate(Panel = target$Panel, ModelRole = "trait_adjusted_sensitivity", .before = 1)
      }

      f4_plot_store[[key]] <- build_plot(
        use, result, method, target$Nutrient, target$ResponseLabel, target$Panel
      )
      f4_fit_store[[key]] <- result
      if (isTRUE(export_diagnostic_figures)) {
        diagnostic_dir <- file.path(out_dir, "M2_diagnostic_figures", "Figure4_all_panels")
        dir.create(diagnostic_dir, recursive = TRUE, showWarnings = FALSE)
        save_plot_pair(
          f4_plot_store[[key]],
          paste("Figure4_panel", target$Panel, method, dataset,
                decision$selected, "pooled_original_style", sep = "_"),
          diagnostic_dir
        )
      }
      f4_rows[[key]] <- result_row(
        use, result, method, dataset, target$Nutrient,
        target$ResponseLabel, target$DefenseCategory, "pooled_traits", decision$selected
      ) %>%
        dplyr::mutate(
          Panel = target$Panel,
          SelectedModel = decision$selected,
          DeltaAICc_linear_minus_quadratic = decision$delta_aicc,
          DeltaR2_quadratic_minus_linear = decision$delta_r2,
          QuadraticVertexX = decision$vertex,
          VertexInsideObservedRange = decision$vertex_inside,
          ShapeSelectionRule = paste0(
            "quadratic only if n>=", publication_min_n,
            ", systems>=", quadratic_min_systems,
            ", DeltaAICc>=", quadratic_min_delta_aicc,
            ", DeltaR2>=", quadratic_min_delta_r2,
            ", p(b2)<=", quadratic_max_p_ordinary,
            ", and vertex is within the central observed x range; otherwise linear"
          ),
          .before = 1
        )

      if (dataset == "RAW") {
        selected_form <- decision$selected
        ## These weights test dominance without deleting or averaging away
        ## individual paired effects. They do not determine the primary curve.
        cap_use <- use %>%
          dplyr::mutate(
            w_analysis = pmin(w_for_model, stats::quantile(w_for_model, .95, na.rm = TRUE))
          )
        study_use <- use %>%
          dplyr::group_by(StudyID) %>%
          dplyr::mutate(w_analysis = w_for_model / sum(w_for_model)) %>%
          dplyr::ungroup()
        trait_use <- use %>%
          dplyr::group_by(YTrait) %>%
          dplyr::mutate(w_analysis = w_for_model / sum(w_for_model)) %>%
          dplyr::ungroup()
        for (sensitivity_name in c("weight_cap_95pct", "study_balanced", "trait_balanced")) {
          sensitivity_data <- switch(
            sensitivity_name,
            weight_cap_95pct = cap_use,
            study_balanced = study_use,
            trait_balanced = trait_use
          )
          sensitivity_fit <- fit_wls(
            sensitivity_data, selected_form, "w_analysis", trait_adjusted = FALSE
          )
          if (!is.null(sensitivity_fit)) {
            sensitivity_key <- paste(key, sensitivity_name, sep = "__")
            f4_sensitivity_rows[[sensitivity_key]] <- result_row(
              sensitivity_data, sensitivity_fit, method, dataset, target$Nutrient,
              target$ResponseLabel, target$DefenseCategory, "pooled_traits",
              selected_form, sensitivity_name
            ) %>% dplyr::mutate(Panel = target$Panel, .before = 1)
          }
        }
      }
      f4_points[[key]] <- use %>%
        dplyr::mutate(Panel = target$Panel, Method = method, Dataset = dataset)
    }
  }
}

if (isTRUE(export_diagnostic_figures)) {
  diagnostic_dir <- file.path(out_dir, "M2_diagnostic_figures", "Figure4_all_panels")
  dir.create(diagnostic_dir, recursive = TRUE, showWarnings = FALSE)
  for (method in methods_to_run) {
    for (dataset in c("RAW", "IQR_removed")) {
      keys <- paste(method, dataset, c("A", "B", "C", "D"), sep = "__")
      if (!all(keys %in% names(f4_plot_store))) stop("A Figure 4 panel is missing before combination.")
      pp <- f4_plot_store[keys]
      combo <- (pp[[1]] | pp[[2]]) / (pp[[3]] | pp[[4]])
      stem <- paste("Figure4", method, dataset, "selected_shape_pooled_original_style", sep = "_")
      save_plot_pair(combo, stem, diagnostic_dir, combined_pdf_w, combined_pdf_h)
    }
  }
}

## Dedicated Figure 4 for submission: corrected M2, all RAW points, and the
## selected linear/quadratic shapes. The confidence ribbons use the same WLS
## fits as the curves and show confidence in the fitted mean, not prediction
## intervals for individual observations.
f4_publication_keys <- paste(primary_method, "RAW", c("A", "B", "C", "D"), sep = "__")
if (!all(f4_publication_keys %in% names(f4_plot_store))) {
  stop("Publication Figure 4 cannot be assembled because at least one M2-RAW panel is missing.")
}
f4_publication_plots <- f4_plot_store[f4_publication_keys]
f4_publication_plot <-
  (f4_publication_plots[[1]] | f4_publication_plots[[2]]) /
  (f4_publication_plots[[3]] | f4_publication_plots[[4]])
save_plot_pair(
  f4_publication_plot,
  "Figure4_FINAL_M2_RAW_model_based_95CI",
  publication_dir,
  combined_pdf_w,
  combined_pdf_h
)
## Export the four Figure 4 panels separately without A-D tags. The tagged
## objects above remain unchanged, preserving the submitted combined figure.
for (ii in seq_len(nrow(f4_plan))) {
  target <- f4_plan[ii, ]
  key <- paste(primary_method, "RAW", target$Panel, sep = "__")
  panel_result <- f4_fit_store[[key]]
  panel_untagged <- f4_plot_store[[key]] +
    ggplot2::labs(tag = NULL) +
    ggplot2::theme(plot.tag = ggplot2::element_blank())
  save_plot_pair(
    panel_untagged,
    paste0(
      "Figure4_panel_", target$Panel, "_FINAL_M2_RAW_",
      target$Nutrient, "_vs_", safe_file_label(target$ResponseLabel), "_",
      panel_result$form, "_model_based_95CI"
    ),
    publication_dir,
    width = pdf_w,
    height = pdf_h,
    save_tiff = TRUE
  )
}
f4_publication_band_audit <- purrr::map_dfr(seq_len(nrow(f4_plan)), function(ii) {
  target <- f4_plan[ii, ]
  key <- paste(primary_method, "RAW", target$Panel, sep = "__")
  confidence_band_audit_row(
    f4_points[[key]], f4_fit_store[[key]], "Figure4", target$Panel,
    target$Nutrient, target$ResponseLabel,
    f4_fit_store[[key]]$form
  )
})

f4_summary <- dplyr::bind_rows(f4_rows)
f4_all_models <- dplyr::bind_rows(f4_all_model_rows)
f4_weight_sensitivity <- dplyr::bind_rows(f4_sensitivity_rows)
f4_point_data <- dplyr::bind_rows(f4_points)

# =========================
# 9) Complete Figure 5 rescreen
# =========================
parse_row_ids <- function(x) {
  x <- as.character(x)
  x <- x[!is.na(x) & nzchar(x)]
  if (!length(x)) return(integer(0))
  ids <- suppressWarnings(as.integer(trimws(unlist(strsplit(x, ";", fixed = TRUE)))))
  ids[is.finite(ids)]
}

## Typical within-effect sampling variance used by metafor for I2 in a
## conventional random-effects model.  For the present multilevel model, the
## denominator is extended with the SUM of all estimated random-intercept
## variance components.  This makes total I2 directly comparable with the
## earlier Table S3 convention while retaining StudyID/SystemID dependence.
typical_sampling_variance <- function(vi) {
  vi <- as.numeric(vi)
  vi <- vi[is.finite(vi) & vi >= 0]
  if (length(vi) <= 1L) return(NA_real_)
  wi <- 1 / pmax(vi, eps)
  sw <- sum(wi)
  denominator <- sw - sum(wi^2) / sw
  if (!is.finite(denominator) || denominator <= 0) return(NA_real_)
  (length(vi) - 1) / denominator
}

multilevel_heterogeneity <- function(result, vi) {
  sigma2 <- suppressWarnings(as.numeric(result$sigma2))
  sigma2 <- sigma2[is.finite(sigma2) & sigma2 >= 0]
  tau2_total <- if (length(sigma2)) sum(sigma2) else NA_real_
  tau2_study <- if (length(sigma2) >= 1L) sigma2[1] else NA_real_
  ## random = ~1|StudyID/SystemID expands to a StudyID component followed by
  ## the nested StudyID/SystemID component.  If metafor ever returns more than
  ## two nested components, all lower-level components are combined here.
  tau2_system <- if (length(sigma2) >= 2L) sum(sigma2[-1]) else NA_real_
  typical_vi <- typical_sampling_variance(vi)
  total_variance <- tau2_total + typical_vi
  i2_total <- if (
    is.finite(tau2_total) && is.finite(total_variance) && total_variance > 0
  ) 100 * tau2_total / total_variance else NA_real_
  i2_study <- if (
    is.finite(tau2_study) && is.finite(total_variance) && total_variance > 0
  ) 100 * tau2_study / total_variance else NA_real_
  i2_system <- if (
    is.finite(tau2_system) && is.finite(total_variance) && total_variance > 0
  ) 100 * tau2_system / total_variance else NA_real_
  list(
    tau2 = tau2_total,
    I2 = i2_total,
    tau2_StudyID = tau2_study,
    tau2_SystemID = tau2_system,
    I2_StudyID = i2_study,
    I2_SystemID = i2_system,
    TypicalSamplingVariance = typical_vi,
    QE = suppressWarnings(as.numeric(result$QE)[1]),
    QE_p = suppressWarnings(as.numeric(result$QEp)[1])
  )
}

pool_lnrr_one_trait <- function(frame) {
  frame <- frame %>%
    dplyr::filter(valid_lnrr, is.finite(lnRR_value), is.finite(Var_lnRR), Var_lnRR >= 0)
  k <- nrow(frame)
  if (!k) return(tibble::tibble(
    Estimate = NA_real_, SE = NA_real_,
    CI_lower = NA_real_, CI_upper = NA_real_, P_value = NA_real_,
    tau2 = NA_real_, I2 = NA_real_,
    tau2_StudyID = NA_real_, tau2_SystemID = NA_real_,
    I2_StudyID = NA_real_, I2_SystemID = NA_real_,
    TypicalSamplingVariance = NA_real_, QE = NA_real_, QE_p = NA_real_,
    k = 0L, Studies = 0L, Systems = 0L, Model = "no_data"
  ))
  if (k == 1L) {
    se <- sqrt(pmax(frame$Var_lnRR[1], eps))
    estimate <- frame$lnRR_value[1]
    z975 <- stats::qnorm(.975)
    return(tibble::tibble(
      Estimate = estimate, SE = se,
      CI_lower = estimate - z975 * se,
      CI_upper = estimate + z975 * se,
      P_value = 2 * stats::pnorm(-abs(estimate / se)),
      tau2 = NA_real_, I2 = NA_real_,
      tau2_StudyID = NA_real_, tau2_SystemID = NA_real_,
      I2_StudyID = NA_real_, I2_SystemID = NA_real_,
      TypicalSamplingVariance = NA_real_, QE = NA_real_, QE_p = NA_real_,
      k = k, Studies = 1L, Systems = 1L, Model = "single_effect"
    ))
  }
  result <- tryCatch(
    metafor::rma.mv(
      yi = lnRR_value, V = Var_lnRR,
      random = ~ 1 | StudyID/SystemID,
      data = frame, method = "REML"
    ),
    error = function(e) NULL
  )
  if (is.null(result)) {
    result_uni <- metafor::rma.uni(
      yi = lnRR_value, vi = Var_lnRR, data = frame, method = "REML"
    )
    return(tibble::tibble(
      Estimate = as.numeric(result_uni$b[1]),
      SE = as.numeric(result_uni$se[1]),
      CI_lower = as.numeric(result_uni$ci.lb[1]),
      CI_upper = as.numeric(result_uni$ci.ub[1]),
      P_value = as.numeric(result_uni$pval[1]),
      tau2 = as.numeric(result_uni$tau2),
      I2 = as.numeric(result_uni$I2),
      tau2_StudyID = NA_real_, tau2_SystemID = NA_real_,
      I2_StudyID = NA_real_, I2_SystemID = NA_real_,
      TypicalSamplingVariance = typical_sampling_variance(frame$Var_lnRR),
      QE = as.numeric(result_uni$QE), QE_p = as.numeric(result_uni$QEp),
      k = k,
      Studies = dplyr::n_distinct(frame$StudyID),
      Systems = dplyr::n_distinct(frame$SystemID),
      Model = "rma.uni_REML_fallback"
    ))
  }
  heterogeneity <- multilevel_heterogeneity(result, frame$Var_lnRR)
  tibble::tibble(
    Estimate = as.numeric(result$b[1]),
    SE = as.numeric(result$se[1]),
    CI_lower = as.numeric(result$ci.lb[1]),
    CI_upper = as.numeric(result$ci.ub[1]),
    P_value = as.numeric(result$pval[1]),
    tau2 = heterogeneity$tau2,
    I2 = heterogeneity$I2,
    tau2_StudyID = heterogeneity$tau2_StudyID,
    tau2_SystemID = heterogeneity$tau2_SystemID,
    I2_StudyID = heterogeneity$I2_StudyID,
    I2_SystemID = heterogeneity$I2_SystemID,
    TypicalSamplingVariance = heterogeneity$TypicalSamplingVariance,
    QE = heterogeneity$QE,
    QE_p = heterogeneity$QE_p,
    k = k,
    Studies = dplyr::n_distinct(frame$StudyID),
    Systems = dplyr::n_distinct(frame$SystemID),
    Model = "rma.mv_REML_StudyID_SystemID"
  )
}

forest_axis_breaks <- function(x_range, max_breaks = f5_forest_max_x_breaks) {
  x_range <- range(x_range, finite = TRUE)
  if (length(x_range) != 2L || any(!is.finite(x_range)) || x_range[1] >= x_range[2]) {
    return(0)
  }
  tolerance <- 1e-10 * max(1, max(abs(x_range)))
  candidates <- pretty(x_range, n = 2)
  candidates <- candidates[
    candidates >= x_range[1] - tolerance &
      candidates <= x_range[2] + tolerance
  ]
  if (x_range[1] <= 0 && x_range[2] >= 0) candidates <- c(candidates, 0)
  candidates <- sort(unique(candidates[is.finite(candidates)]))
  if (!length(candidates)) candidates <- mean(x_range)

  if (length(candidates) > max_breaks) {
    zero_index <- which(abs(candidates) <= tolerance)
    if (length(zero_index)) {
      negative <- candidates[candidates < -tolerance]
      positive <- candidates[candidates > tolerance]
      chosen <- 0
      if (length(negative)) chosen <- c(min(negative), chosen)
      if (length(positive)) chosen <- c(chosen, max(positive))
      if (length(chosen) < max_breaks) {
        remaining <- setdiff(candidates, chosen)
        chosen <- c(chosen, remaining[seq_len(min(
          length(remaining), max_breaks - length(chosen)
        ))])
      }
      candidates <- sort(unique(chosen))[seq_len(min(max_breaks, length(unique(chosen))))]
    } else {
      keep <- unique(round(seq(1, length(candidates), length.out = max_breaks)))
      candidates <- candidates[keep]
    }
  }
  candidates
}

forest_axis_labels <- function(x) {
  x <- as.numeric(x)
  nonzero <- abs(x[is.finite(x) & abs(x) > 1e-12])
  smallest <- if (length(nonzero)) min(nonzero) else 1
  digits <- if (smallest >= 1) 0L else if (smallest >= 0.1) 1L else if (smallest >= 0.01) 2L else 3L
  labels <- formatC(x, format = "f", digits = digits)
  labels[is.finite(x) & abs(x) <= 1e-12] <- "0"
  labels
}

make_companion_forest <- function(reg_dat, nutrient, defense_trait) {
  x_ids <- parse_row_ids(reg_dat$X_RowIDs)
  y_ids <- suppressWarnings(as.integer(reg_dat$Y_RowID))
  ids <- unique(c(x_ids, y_ids[is.finite(y_ids)]))
  if (!length(ids)) return(NULL)
  forest_rows <- dat_all %>%
    dplyr::filter(RowID %in% ids, Trait %in% c(nutrient, defense_trait), valid_lnrr)
  if (!nrow(forest_rows)) return(NULL)
  summary_rows <- forest_rows %>%
    dplyr::group_by(Trait) %>%
    dplyr::group_modify(~ pool_lnrr_one_trait(.x)) %>%
    dplyr::ungroup() %>%
    dplyr::filter(is.finite(Estimate), is.finite(CI_lower), is.finite(CI_upper))
  if (!nrow(summary_rows)) return(NULL)
  trait_order <- c(nutrient, defense_trait)
  trait_order <- trait_order[trait_order %in% as.character(summary_rows$Trait)]
  n_lookup <- stats::setNames(summary_rows$k, as.character(summary_rows$Trait))
  display_lookup <- stats::setNames(
    ifelse(
      trait_order == nutrient,
      nutrient_axis_label(trait_order),
      trait_order
    ),
    trait_order
  )
  label_levels <- paste0(display_lookup[rev(trait_order)], "\n(n=", n_lookup[rev(trait_order)], ")")
  summary_rows <- summary_rows %>%
    dplyr::mutate(
      Trait = factor(Trait, levels = rev(trait_order)),
      TraitLabel = paste0(display_lookup[as.character(Trait)], "\n(n=", k, ")"),
      TraitLabel = factor(TraitLabel, levels = label_levels),
      SignifStar = dplyr::case_when(
        P_value < 0.001 ~ "***", P_value < 0.01 ~ "**",
        P_value < 0.05 ~ "*", TRUE ~ ""
      )
    )
  xr <- range(c(summary_rows$CI_lower, summary_rows$CI_upper, 0), finite = TRUE)
  span <- diff(xr)
  if (!is.finite(span) || span <= 0) span <- 1
  star_x <- xr[2] + 0.16 * span
  limits <- c(xr[1] - 0.16 * span, xr[2] + 0.34 * span)
  forest_break_values <- forest_axis_breaks(xr)
  forest_break_labels <- forest_axis_labels(forest_break_values)
  if (length(forest_break_values) > f5_forest_max_x_breaks ||
      anyDuplicated(forest_break_labels)) {
    stop("Figure 5 companion-forest x-axis tick self-check failed.")
  }
  summary_rows$StarX <- star_x
  p <- ggplot2::ggplot(summary_rows, ggplot2::aes(x = Estimate, y = TraitLabel)) +
    ggplot2::geom_vline(xintercept = 0, color = "grey72", linewidth = .5, linetype = "dashed") +
    ggplot2::geom_segment(
      ggplot2::aes(x = CI_lower, xend = CI_upper, y = TraitLabel, yend = TraitLabel),
      linewidth = .75, color = "black"
    ) +
    ggplot2::geom_point(shape = 21, size = 2.8, stroke = .35, fill = "grey45") +
    ggplot2::geom_text(
      ggplot2::aes(x = StarX, label = SignifStar),
      family = plot_font_family, color = "#D7191C",
      fontface = "bold", size = 4.8
    ) +
    ggplot2::scale_x_continuous(
      breaks = forest_break_values,
      labels = forest_break_labels
    ) +
    ggplot2::coord_cartesian(xlim = limits, clip = "on") +
    ggplot2::labs(x = "Effect size (lnRR)", y = NULL) +
    ggplot2::theme_classic(base_size = base_font_size, base_family = plot_font_family) +
    ggplot2::theme(
      text = ggplot2::element_text(family = plot_font_family, color = "black"),
      axis.title.x = ggplot2::element_text(
        family = plot_font_family, face = "bold",
        size = f5_forest_axis_title_size,
        margin = ggplot2::margin(t = 7)
      ),
      axis.text.x = ggplot2::element_text(
        family = plot_font_family, face = "plain",
        size = f5_forest_axis_text_size, color = "black"
      ),
      axis.text.y = ggplot2::element_text(
        family = plot_font_family, face = "bold",
        size = base_font_size - 3, color = "black"
      ),
      axis.line = ggplot2::element_line(
        linewidth = axis_line_width, color = "black"
      ),
      axis.ticks = ggplot2::element_line(
        linewidth = axis_tick_width, color = "black"
      ),
      axis.ticks.length = grid::unit(axis_tick_length_cm, "cm"),
      plot.margin = ggplot2::margin(10, 10, 10, 6)
    )
  list(
    plot = p,
    summary = summary_rows,
    axis_breaks = forest_break_values,
    axis_labels = forest_break_labels
  )
}

make_combined_f5_panel <- function(forest_plot, regression_plot) {
  if (is.null(forest_plot)) return(regression_plot)
  forest_plot + regression_plot + patchwork::plot_layout(widths = c(1.08, 2.50))
}

## Build one lnRR funnel object from the exact source RowIDs used in one
## M2-RAW Figure 5 regression. This deliberately uses lnRR rather than Cha,
## matching the old Figure S6 and the companion forest plot on the left.
make_f5_publication_funnel <- function(
    reg_dat, panel, nutrient, defense_trait,
    role = c("nutrient", "defense")) {
  role <- match.arg(role)
  if (role == "nutrient") {
    row_ids <- parse_row_ids(reg_dat$X_RowIDs)
    target_trait <- nutrient
    role_order <- 1L
    role_label <- "Nutrient"
  } else {
    row_ids <- suppressWarnings(as.integer(reg_dat$Y_RowID))
    row_ids <- row_ids[is.finite(row_ids)]
    target_trait <- defense_trait
    role_order <- 2L
    role_label <- "Defense"
  }
  row_ids <- unique(row_ids)
  if (!length(row_ids)) return(NULL)

  fdat <- dat_all %>%
    dplyr::filter(
      RowID %in% row_ids,
      Trait == target_trait,
      valid_lnrr,
      is.finite(lnRR_value),
      is.finite(Var_lnRR), Var_lnRR > 0
    ) %>%
    dplyr::mutate(SE_lnRR = sqrt(Var_lnRR)) %>%
    dplyr::arrange(StudyID, SystemID, RowID)

  ## Three effect sizes are the minimum retained from the old F5 funnel logic.
  if (nrow(fdat) < 3L) return(NULL)

  funnel_model <- tryCatch(
    metafor::rma.uni(
      yi = lnRR_value, vi = Var_lnRR,
      data = fdat, method = "REML"
    ),
    error = function(e) NULL
  )
  funnel_model_type <- "rma.uni_REML"
  if (is.null(funnel_model)) {
    funnel_model <- tryCatch(
      metafor::rma.uni(
        yi = lnRR_value, vi = Var_lnRR,
        data = fdat, method = "FE"
      ),
      error = function(e) NULL
    )
    funnel_model_type <- "rma.uni_FE_fallback"
  }
  if (is.null(funnel_model)) return(NULL)

  pooled <- pool_lnrr_one_trait(fdat)
  egger_p <- NA_real_
  begg_p <- NA_real_
  trimfill_k0 <- NA_real_
  if (nrow(fdat) >= funnel_min_effects_for_bias_tests) {
    egger_p <- tryCatch(
      as.numeric(metafor::regtest(funnel_model, model = "rma")$pval),
      error = function(e) NA_real_
    )
    if (isTRUE(funnel_run_begg_ranktest)) {
      begg_p <- tryCatch(
        ## Tied effect sizes/variances are expected in paired meta-analytic
        ## data.  Use the asymptotic Kendall test explicitly so cor.test does
        ## not warn that an exact P value is unavailable in the presence of ties.
        as.numeric(metafor::ranktest(funnel_model, exact = FALSE)$pval),
        error = function(e) NA_real_
      )
    }
    if (isTRUE(funnel_run_trimfill)) {
      trimfill_object <- tryCatch(
        metafor::trimfill(funnel_model),
        error = function(e) NULL
      )
      if (!is.null(trimfill_object)) {
        trimfill_k0 <- suppressWarnings(as.numeric(trimfill_object$k0))
      }
    }
  }

  display_trait <- if (role == "nutrient") {
    nutrient_axis_label(target_trait)
  } else {
    target_trait
  }
  funnel_panel <- paste0(panel, role_order)
  relationship <- paste0(nutrient, " vs. ", defense_trait)
  old_style_title <- paste0("lnRR ",nutrient_axis_label(nutrient)," VS. ",defense_trait,"-",display_trait)
  summary_row <- pooled %>%
    dplyr::transmute(
      Panel = panel,
      FunnelPanel = funnel_panel,
      PanelOrder = match(panel, LETTERS),
      RoleOrder = role_order,
      Role = role_label,
      Relationship = relationship,
      Nutrient = nutrient,
      DefenseTrait = defense_trait,
      FunnelTrait = target_trait,
      DisplayTrait = display_trait,
      FunnelTitle = old_style_title,
      n_effect_sizes = nrow(fdat),
      n_studies = dplyr::n_distinct(fdat$StudyID),
      n_systems = dplyr::n_distinct(fdat$SystemID),
      PooledEstimate_lnRR = Estimate,
      PooledSE_lnRR = SE,
      CI_lower_lnRR = CI_lower,
      CI_upper_lnRR = CI_upper,
      Pooled_P = P_value,
      tau2_total = tau2,
      I2_total = I2,
      tau2_StudyID = tau2_StudyID,
      tau2_SystemID = tau2_SystemID,
      I2_StudyID = I2_StudyID,
      I2_SystemID = I2_SystemID,
      TypicalSamplingVariance = TypicalSamplingVariance,
      Heterogeneity_Q = QE,
      Heterogeneity_Q_P = QE_p,
      PooledModel = Model,
      FunnelModel = funnel_model_type,
      Egger_P = egger_p,
      Begg_P = begg_p,
      TrimFill_k0 = trimfill_k0
    )

  list(
    data = fdat,
    model = funnel_model,
    summary = summary_row,
    panel = panel,
    funnel_panel = funnel_panel,
    role = role_label,
    role_order = role_order,
    nutrient = nutrient,
    defense_trait = defense_trait,
    target_trait = target_trait,
    display_trait = display_trait,
    title = old_style_title,
    egger_p = egger_p,
    begg_p = begg_p,
    trimfill_k0 = trimfill_k0,
    n_effect_sizes = nrow(fdat),
    n_studies = dplyr::n_distinct(fdat$StudyID)
  )
}

funnel_format_p <- function(p) {
  if (!length(p) || is.na(p) || !is.finite(p)) return("NA")
  if (p < 0.001) return("< 0.001")
  formatC(p, digits = 3, format = "f")
}

current_stage <- "F5 models and graphics"
defense_inventory <- dat %>%
  dplyr::filter(!(Trait %in% c("C", "N", "P"))) %>%
  dplyr::group_by(Trait, TraitCategory) %>%
  dplyr::summarise(TraitForm = paste(sort(unique(TraitForm)), collapse = ";"), .groups = "drop")

f5_plan <- dplyr::bind_rows(
  defense_inventory %>%
    dplyr::transmute(
      Nutrient = "C", DefenseTrait = Trait,
      DefenseCategory = TraitCategory, TraitForm
    ),
  defense_inventory %>%
    dplyr::filter(TraitCategory == "chemical_defense") %>%
    dplyr::transmute(
      Nutrient = "N", DefenseTrait = Trait,
      DefenseCategory = TraitCategory, TraitForm
    ),
  if (include_p_specific_f5) {
    defense_inventory %>%
      dplyr::filter(TraitCategory == "chemical_defense") %>%
      dplyr::transmute(
        Nutrient = "P", DefenseTrait = Trait,
        DefenseCategory = TraitCategory, TraitForm
      )
  } else NULL
) %>%
  dplyr::distinct() %>%
  dplyr::arrange(Nutrient, DefenseCategory, DefenseTrait)

availability_rows <- list()
f5_all_model_rows <- list()
f5_selected_rows <- list()
f5_sensitivity_rows <- list()
f5_point_rows <- list()
f5_plot_store <- list()
f5_regression_plot_store <- list()
f5_forest_summaries <- list()

for (ii in seq_len(nrow(f5_plan))) {
  target <- f5_plan[ii, ]
  pair_raw <- build_pair_data(
    target$Nutrient, y_trait = target$DefenseTrait,
    metric = "raw_difference"
  )
  availability_rows[[paste(target$Nutrient, target$DefenseTrait, sep = "__")]] <- tibble::tibble(
    Nutrient = target$Nutrient,
    DefenseTrait = target$DefenseTrait,
    DefenseCategory = target$DefenseCategory,
    TraitForm = target$TraitForm,
    PairedRows = nrow(pair_raw),
    Studies = dplyr::n_distinct(pair_raw$StudyID),
    Systems = dplyr::n_distinct(pair_raw$SystemID),
    FitPossible = nrow(pair_raw) >= min_n_linear && dplyr::n_distinct(pair_raw$X_value_raw) >= 2
  )
  if (nrow(pair_raw) < min_n_linear || dplyr::n_distinct(pair_raw$X_value_raw) < 2) next

  for (method in methods_to_run) {
    pair <- if (method_metric(method) == "raw_difference") {
      pair_raw
    } else {
      build_pair_data(
        target$Nutrient, y_trait = target$DefenseTrait,
        metric = method_metric(method)
      )
    }
    transformed <- transform_pair(pair, method)
    datasets <- list(RAW = transformed)
    if (run_iqr_sensitivity) datasets$IQR_removed <- iqr_filter_2d(transformed, iqr_k)

    for (dataset in names(datasets)) {
      use <- datasets[[dataset]]
      linear <- fit_wls(use, "linear")
      quadratic <- fit_wls(use, "quadratic")
      if (is.null(linear)) next

      linear_row <- result_row(
        use, linear, method, dataset, target$Nutrient, target$DefenseTrait,
        target$DefenseCategory, target$TraitForm, "linear"
      )
      f5_all_model_rows[[paste(method, dataset, target$Nutrient, target$DefenseTrait, "linear", sep = "__")]] <- linear_row

      if (!is.null(quadratic)) {
        quadratic_row <- result_row(
          use, quadratic, method, dataset, target$Nutrient, target$DefenseTrait,
          target$DefenseCategory, target$TraitForm, "quadratic"
        )
        f5_all_model_rows[[paste(method, dataset, target$Nutrient, target$DefenseTrait, "quadratic", sep = "__")]] <- quadratic_row
      }

      decision <- choose_model_shape(linear, quadratic, use)
      selected_form <- decision$selected
      selected_result <- if (selected_form == "quadratic") quadratic else linear
      selected_row <- result_row(
        use, selected_result, method, dataset, target$Nutrient, target$DefenseTrait,
        target$DefenseCategory, target$TraitForm, selected_form
      ) %>%
        dplyr::mutate(
          SelectedModel = selected_form,
          DeltaAICc_linear_minus_quadratic = decision$delta_aicc,
          DeltaR2_quadratic_minus_linear = decision$delta_r2,
          QuadraticVertexX = decision$vertex,
          VertexInsideObservedRange = decision$vertex_inside,
          ShapeSelectionRule = paste0(
            "quadratic only if n>=", publication_min_n,
            ", systems>=", quadratic_min_systems,
            ", DeltaAICc>=", quadratic_min_delta_aicc,
            ", DeltaR2>=", quadratic_min_delta_r2,
            ", p(b2)<=", quadratic_max_p_ordinary,
            ", and vertex is within the central observed x range; otherwise linear"
          )
        )
      selected_key <- paste(method, dataset, target$Nutrient, target$DefenseTrait, sep = "__")
      f5_selected_rows[[selected_key]] <- selected_row
      f5_point_rows[[selected_key]] <- use %>%
        dplyr::mutate(
          Method = method, Dataset = dataset,
          Nutrient = target$Nutrient, DefenseTrait = target$DefenseTrait
        )

      regression_plot <- build_plot(
        use, selected_result, method, target$Nutrient, target$DefenseTrait, NULL
      )
      forest_obj <- make_companion_forest(use, target$Nutrient, target$DefenseTrait)
      combined_plot <- make_combined_f5_panel(
        if (is.null(forest_obj)) NULL else forest_obj$plot,
        regression_plot
      )
      f5_regression_plot_store[[selected_key]] <- regression_plot
      f5_plot_store[[selected_key]] <- combined_plot
      if (!is.null(forest_obj)) {
        f5_forest_summaries[[selected_key]] <- forest_obj$summary %>%
          dplyr::mutate(
            Method = method, Dataset = dataset,
            Nutrient = target$Nutrient, DefenseTrait = target$DefenseTrait,
            SelectedModel = selected_form
          )
      }
      stem <- paste(
        "Figure5_candidate", method, dataset,
        paste0(target$Nutrient, "_vs_", safe_file_label(target$DefenseTrait)),
        selected_form, "forest_plus_regression", sep = "_"
      )
      ## Candidate statistics and paired data are retained in the result tables.
      ## Candidate figures are omitted from the publication-only output unless
      ## export_diagnostic_figures is explicitly changed to TRUE.
      if (isTRUE(export_diagnostic_figures)) {
        diagnostic_dir <- file.path(out_dir, "M2_diagnostic_figures", "Figure5_all_candidates")
        dir.create(diagnostic_dir, recursive = TRUE, showWarnings = FALSE)
        save_plot_pair(
          combined_plot, stem, diagnostic_dir,
          width = combined_pdf_w, height = pdf_h, save_tiff = FALSE
        )
      }

      if (dataset == "RAW") {
        ## Sensitivity 1: cap individual weights at the 95th percentile.
        cap_use <- use %>%
          dplyr::mutate(
            w_analysis = pmin(w_for_model, stats::quantile(w_for_model, .95, na.rm = TRUE))
          )
        cap_result <- fit_wls(cap_use, selected_form, "w_analysis")
        if (!is.null(cap_result)) {
          f5_sensitivity_rows[[paste(selected_key, "weight_cap_95pct", sep = "__")]] <-
            result_row(
              cap_use, cap_result, method, dataset, target$Nutrient, target$DefenseTrait,
              target$DefenseCategory, target$TraitForm, selected_form, "weight_cap_95pct"
            )
        }

        ## Sensitivity 2: every StudyID receives equal total weight.
        balanced_use <- use %>%
          dplyr::group_by(StudyID) %>%
          dplyr::mutate(w_analysis = w_for_model / sum(w_for_model)) %>%
          dplyr::ungroup()
        balanced_result <- fit_wls(balanced_use, selected_form, "w_analysis")
        if (!is.null(balanced_result)) {
          f5_sensitivity_rows[[paste(selected_key, "study_balanced", sep = "__")]] <-
            result_row(
              balanced_use, balanced_result, method, dataset, target$Nutrient, target$DefenseTrait,
              target$DefenseCategory, target$TraitForm, selected_form, "study_balanced"
            )
        }
      }
    }
  }
}

f5_availability <- dplyr::bind_rows(availability_rows)
f5_all_models <- dplyr::bind_rows(f5_all_model_rows)
f5_selected <- dplyr::bind_rows(f5_selected_rows)
f5_sensitivity <- dplyr::bind_rows(f5_sensitivity_rows)
f5_point_data <- dplyr::bind_rows(f5_point_rows)
f5_forest_summary <- dplyr::bind_rows(f5_forest_summaries)

## Publication-ready Table S3 block: exactly two forest rows (defense first,
## nutrient second) for each of the six final Figure 5 relationships.
f5_publication_forest_table <- f5_forest_summary %>%
  dplyr::mutate(Trait = as.character(Trait)) %>%
  dplyr::filter(Method == primary_method, Dataset == "RAW") %>%
  dplyr::inner_join(
    f5_publication_plan %>%
      dplyr::select(Panel, Nutrient, DefenseTrait, RequestedModel),
    by = c(
      "Nutrient", "DefenseTrait",
      "SelectedModel" = "RequestedModel"
    )
  ) %>%
  dplyr::mutate(
    TraitOrder = dplyr::if_else(Trait == DefenseTrait, 1L, 2L),
    TraitDisplay = dplyr::case_when(
      Trait == "C" ~ "Plant carbon content",
      Trait == "N" ~ "Plant nitrogen content",
      Trait == "P" ~ "Plant phosphorus content",
      TRUE ~ Trait
    )
  ) %>%
  dplyr::arrange(factor(Panel, levels = LETTERS[1:6]), TraitOrder) %>%
  dplyr::transmute(
    Figure_Name = paste0("Figure 5", Panel),
    Trait = TraitDisplay,
    Estimate, SE, CI_lower, CI_upper, P_value,
    tau2, I2, n = k, SignifStar,
    ComparisonType = "introduced_vs_native_range",
    AnalysisID = paste0(Nutrient, "_vs_", DefenseTrait),
    Studies, Systems, Model,
    tau2_StudyID, tau2_SystemID,
    I2_StudyID, I2_SystemID,
    TypicalSamplingVariance,
    Heterogeneity_Q = QE,
    Heterogeneity_Q_P = QE_p
  )

forest_heterogeneity_bad <- f5_forest_summary %>%
  dplyr::filter(Model %in% c(
    "rma.mv_REML_StudyID_SystemID", "rma.uni_REML_fallback"
  )) %>%
  dplyr::filter(
    !is.finite(SE) | SE <= 0 |
      !is.finite(tau2) | tau2 < 0 |
      !is.finite(I2) | I2 < -1e-8 | I2 > 100 + 1e-8 |
      !is.finite(TypicalSamplingVariance) | TypicalSamplingVariance <= 0
  )

forest_component_bad <- f5_forest_summary %>%
  dplyr::filter(Model == "rma.mv_REML_StudyID_SystemID") %>%
  dplyr::mutate(
    tau2_component_sum = tau2_StudyID + tau2_SystemID,
    I2_component_sum = I2_StudyID + I2_SystemID,
    tau2_tolerance = 1e-8 * pmax(1, abs(tau2)),
    I2_tolerance = 1e-8 * pmax(1, abs(I2))
  ) %>%
  dplyr::filter(
    !is.finite(tau2_component_sum) |
      !is.finite(I2_component_sum) |
      abs(tau2 - tau2_component_sum) > tau2_tolerance |
      abs(I2 - I2_component_sum) > I2_tolerance
  )

if (nrow(forest_heterogeneity_bad)) {
  stop(
    "Self-check failed: invalid SE/tau2/I2 in F5 companion forest results; bad rows = ",
    nrow(forest_heterogeneity_bad)
  )
}
if (nrow(forest_component_bad)) {
  stop(
    "Self-check failed: multilevel heterogeneity components do not sum to totals; bad rows = ",
    nrow(forest_component_bad)
  )
}
if (nrow(f5_publication_forest_table) != 12L) {
  stop(
    "Self-check failed: Table S3 Figure 5 block must contain exactly 12 rows; found ",
    nrow(f5_publication_forest_table)
  )
}

if (nrow(f5_selected)) {
  f5_selected <- f5_selected %>%
    dplyr::group_by(Method, Dataset) %>%
    dplyr::mutate(
      Q_selected_term_ordinary_BH = stats::p.adjust(P_selected_term_ordinary, method = "BH"),
      Q_selected_term_CR2_BH = stats::p.adjust(P_selected_term_CR2, method = "BH")
    ) %>%
    dplyr::ungroup()
}

balanced_evidence <- f5_sensitivity %>%
  dplyr::filter(Sensitivity == "study_balanced") %>%
  dplyr::select(
    Method, Nutrient, DefenseTrait,
    Balanced_R2 = R2_weighted,
    Balanced_LOSO_CV_R2 = LOSO_CV_R2,
    Balanced_term = Selected_term,
    Balanced_P_CR2 = P_selected_term_CR2,
    Balanced_max_study_share = Max_study_weight_share
  )

shape_stability <- f5_selected %>%
  dplyr::select(Method, Dataset, Nutrient, DefenseTrait, SelectedModel, Selected_term) %>%
  tidyr::pivot_wider(
    names_from = Dataset,
    values_from = c(SelectedModel, Selected_term)
  ) %>%
  dplyr::mutate(
    RAW_IQR_shape_match = SelectedModel_RAW == SelectedModel_IQR_removed,
    RAW_IQR_direction_match = sign(Selected_term_RAW) == sign(Selected_term_IQR_removed)
  ) %>%
  dplyr::select(Method, Nutrient, DefenseTrait, RAW_IQR_shape_match, RAW_IQR_direction_match)

f5_selected <- f5_selected %>%
  dplyr::left_join(balanced_evidence, by = c("Method", "Nutrient", "DefenseTrait")) %>%
  dplyr::left_join(shape_stability, by = c("Method", "Nutrient", "DefenseTrait")) %>%
  dplyr::mutate(
    ReplicationEligible = n >= publication_min_n &
      Studies >= publication_min_studies & Systems >= publication_min_systems,
    MainEvidenceCandidate = Dataset == "RAW" &
      n >= publication_min_n &
      Studies >= publication_min_studies &
      Systems >= publication_min_systems &
      R2_weighted >= publication_min_r2 &
      is.finite(Q_selected_term_CR2_BH) & Q_selected_term_CR2_BH <= publication_alpha_cr2_fdr &
      is.finite(LOSO_CV_R2) & LOSO_CV_R2 > 0 &
      is.finite(Balanced_R2) & Balanced_R2 >= publication_min_r2 &
      is.finite(Balanced_P_CR2) & Balanced_P_CR2 <= publication_alpha_balanced_cr2 &
      is.finite(Balanced_LOSO_CV_R2) & Balanced_LOSO_CV_R2 > 0 &
      sign(Selected_term) == sign(Balanced_term) &
      RAW_IQR_shape_match & RAW_IQR_direction_match,
    SupportiveCandidate = Dataset == "RAW" & ReplicationEligible &
      R2_weighted >= publication_min_r2 &
      is.finite(Q_selected_term_ordinary_BH) &
      Q_selected_term_ordinary_BH <= supportive_alpha_ordinary_fdr,
    EvidenceClass = dplyr::case_when(
      MainEvidenceCandidate ~ "main_evidence_candidate",
      SupportiveCandidate ~ "supportive_only",
      ReplicationEligible ~ "replication_only",
      TRUE ~ "insufficient_replication"
    ),
    EvidenceRank = dplyr::case_when(
      EvidenceClass == "main_evidence_candidate" ~ 1L,
      EvidenceClass == "supportive_only" ~ 2L,
      EvidenceClass == "replication_only" ~ 3L,
      TRUE ~ 4L
    )
  ) %>%
  dplyr::group_by(Method, Dataset) %>%
  dplyr::arrange(EvidenceRank, dplyr::desc(R2_weighted), .by_group = TRUE) %>%
  dplyr::mutate(CandidateRankWithinMethodDataset = dplyr::row_number()) %>%
  dplyr::ungroup()

## -------------------------
## Final six-panel Figure 5
## -------------------------
## Refit each requested shape explicitly from the already transformed M2-RAW
## paired data. This does not alter the data or weights; it only guarantees the
## stated panel order and model form in the final publication file. Agreement
## with the automatic shape rule is exported as an audit column.
f5_publication_selection <- f5_publication_plan %>%
  dplyr::left_join(
    f5_selected %>%
      dplyr::filter(Method == primary_method, Dataset == "RAW") %>%
      dplyr::select(
        Nutrient, DefenseTrait,
        AutomaticModel = SelectedModel,
        Automatic_R2 = R2_weighted,
        Automatic_P_ordinary = P_selected_term_ordinary,
        Automatic_P_CR2 = P_selected_term_CR2,
        Automatic_Q_CR2_BH = Q_selected_term_CR2_BH,
        Studies, Systems, EvidenceClass
      ),
    by = c("Nutrient", "DefenseTrait")
  ) %>%
  dplyr::mutate(
    CandidateAvailable = !is.na(AutomaticModel),
    RequestedMatchesAutomatic = RequestedModel == AutomaticModel
  )

if (any(!f5_publication_selection$CandidateAvailable)) {
  missing_panels <- f5_publication_selection %>%
    dplyr::filter(!CandidateAvailable) %>%
    dplyr::transmute(label = paste0(Nutrient, "-", DefenseTrait)) %>%
    dplyr::pull(label)
  stop(
    "Final Figure 5 cannot be assembled; missing corrected M2-RAW candidate(s): ",
    paste(missing_panels, collapse = ", ")
  )
}
if (any(!f5_publication_selection$RequestedMatchesAutomatic)) {
  warning(
    "At least one requested Figure 5 shape differs from the automatic rule. ",
    "The requested shape is drawn and the mismatch is recorded in the audit table."
  )
}

f5_publication_plots <- list()
f5_publication_model_rows <- list()
f5_publication_band_rows <- list()

for (ii in seq_len(nrow(f5_publication_plan))) {
  spec <- f5_publication_plan[ii, ]
  key <- paste(primary_method, "RAW", spec$Nutrient, spec$DefenseTrait, sep = "__")
  use <- f5_point_rows[[key]]
  if (is.null(use) || !nrow(use)) {
    stop("No transformed M2-RAW data for final Figure 5 panel ", spec$Panel, ".")
  }
  requested_result <- fit_wls(use, spec$RequestedModel)
  if (is.null(requested_result)) {
    stop("Requested model could not be fitted for final Figure 5 panel ", spec$Panel, ".")
  }
  trait_info <- f5_plan %>%
    dplyr::filter(
      Nutrient == spec$Nutrient,
      DefenseTrait == spec$DefenseTrait
    ) %>%
    dplyr::slice(1)

  regression_plot <- build_plot(
    use, requested_result, primary_method,
    spec$Nutrient, spec$DefenseTrait, NULL
  )
  forest_obj <- make_companion_forest(use, spec$Nutrient, spec$DefenseTrait)
  if (is.null(forest_obj)) {
    ## Standalone export: no A-F tag. Tagged copy: combined publication figure.
    final_panel_untagged <- regression_plot
    final_panel_tagged <- regression_plot +
      ggplot2::labs(tag = spec$Panel) +
      ggplot2::theme(
        plot.tag = ggplot2::element_text(
          family = plot_font_family, face = "bold",
          size = panel_tag_size, hjust = 0, vjust = 1, color = "black"
        ),
        plot.tag.position = f5_panel_tag_position
      )
  } else {
    final_panel_untagged <- make_combined_f5_panel(
      forest_obj$plot, regression_plot
    )
    forest_plot_tagged <- forest_obj$plot +
      ggplot2::labs(tag = spec$Panel) +
      ggplot2::theme(
        plot.tag = ggplot2::element_text(
          family = plot_font_family, face = "bold",
          size = panel_tag_size, hjust = 0, vjust = 1, color = "black"
        ),
        plot.tag.position = f5_panel_tag_position
      )
    final_panel_tagged <- make_combined_f5_panel(
      forest_plot_tagged, regression_plot
    )
  }
  f5_publication_plots[[spec$Panel]] <- final_panel_tagged

  f5_publication_model_rows[[spec$Panel]] <- result_row(
    use, requested_result, primary_method, "RAW",
    spec$Nutrient, spec$DefenseTrait,
    trait_info$DefenseCategory, trait_info$TraitForm,
    spec$RequestedModel
  ) %>%
    dplyr::mutate(Panel = spec$Panel, RequestedModel = spec$RequestedModel, .before = 1)

  f5_publication_band_rows[[spec$Panel]] <- confidence_band_audit_row(
    use, requested_result, "Figure5", spec$Panel,
    spec$Nutrient, spec$DefenseTrait, spec$RequestedModel
  )

  save_plot_pair(
    final_panel_untagged,
    paste0(
      "Figure5_panel_", spec$Panel, "_FINAL_M2_RAW_",
      spec$Nutrient, "_vs_", safe_file_label(spec$DefenseTrait), "_",
      spec$RequestedModel, "_model_based_95CI"
    ),
    publication_dir,
    width = combined_pdf_w,
    height = pdf_h,
    save_tiff = TRUE
  )
}

f5_publication_model_results <- dplyr::bind_rows(f5_publication_model_rows) %>%
  dplyr::left_join(
    f5_publication_selection %>%
      dplyr::select(
        Panel, AutomaticModel, RequestedMatchesAutomatic,
        Automatic_Q_CR2_BH, EvidenceClass
      ),
    by = "Panel"
  )
f5_publication_band_audit <- dplyr::bind_rows(f5_publication_band_rows)

f5_publication_grid <- patchwork::wrap_plots(
  f5_publication_plots[f5_publication_plan$Panel], ncol = 2
)
save_plot_pair(
  f5_publication_grid,
  "Figure5_FINAL_six_panels_M2_RAW_model_based_95CI",
  publication_dir,
  width = f5_final_grid_pdf_w,
  height = 13.5,
  save_tiff = TRUE
)

## -------------------------------------------------------
## Figure 5 supplementary funnels: exact paired RowID sets
## -------------------------------------------------------
## For every final Figure 5 relationship, reproduce the old S6 logic:
##   1 = nutrient lnRR funnel; 2 = defense-trait lnRR funnel.
## The order A1, A2, B1, B2, ..., F1, F2 is fixed and audited below.
current_stage <- "F5 funnel diagnostics and graphics"
f5_publication_funnel_objects <- list()
f5_publication_funnel_summary_rows <- list()
f5_publication_funnel_effect_rows <- list()

for (ii in seq_len(nrow(f5_publication_plan))) {
  spec <- f5_publication_plan[ii, ]
  key <- paste(primary_method, "RAW", spec$Nutrient, spec$DefenseTrait, sep = "__")
  use <- f5_point_rows[[key]]
  if (is.null(use) || !nrow(use)) {
    stop("No M2-RAW paired data for Figure 5 funnel panel ", spec$Panel, ".")
  }

  panel_funnels <- list()
  for (role in c("nutrient", "defense")) {
    funnel_object <- make_f5_publication_funnel(
      use, spec$Panel, spec$Nutrient, spec$DefenseTrait, role
    )
    if (is.null(funnel_object)) {
      stop(
        "F5 funnel could not be generated for panel ", spec$Panel,
        " (", role, "). Check exact RowIDs and lnRR variances."
      )
    }

    requested_ids <- if(role=="nutrient") parse_row_ids(use$X_RowIDs) else unique(as.integer(use$Y_RowID))
    expected_ids <- dat_all %>% dplyr::filter(RowID %in% requested_ids,Trait==funnel_object$target_trait,
       valid_lnrr,is.finite(Var_lnRR),Var_lnRR>0) %>% dplyr::pull(RowID)
    if(!setequal(funnel_object$data$RowID,expected_ids)) stop("F5 funnel source-row check failed: ",spec$Panel," ",role)
    role_code <- ifelse(role == "nutrient", "N", "D")
    stem <- paste0("F5",spec$Panel,"_",safe_file_label(funnel_object$target_trait),"_funnel")
    save_f5_funnel_collection(
      list(funnel_object), stem, funnel_dir,
      nrow = 1, ncol = 1,
      width = funnel_individual_pdf_w,
      height = funnel_individual_pdf_h,
      compact = FALSE
    )

    funnel_object$summary <- funnel_object$summary %>%
      dplyr::mutate(
        IndependentPdfFile = file.path(
          "F5_funnel", paste0(stem, ".pdf")
        ),
        IndependentTiffFile = file.path(
          "F5_funnel", paste0(stem, ".tiff")
        )
      )
    f5_publication_funnel_summary_rows[[funnel_object$funnel_panel]] <-
      funnel_object$summary
    f5_publication_funnel_effect_rows[[funnel_object$funnel_panel]] <-
      funnel_object$data %>%
      dplyr::transmute(
        Panel = spec$Panel,
        FunnelPanel = funnel_object$funnel_panel,
        Role = funnel_object$role,
        Relationship = paste0(spec$Nutrient, " vs. ", spec$DefenseTrait),
        Nutrient = spec$Nutrient,
        DefenseTrait = spec$DefenseTrait,
        FunnelTrait = funnel_object$target_trait,
        RowID, StudyID, SystemID,
        Mi, Mc, SDi, SDc, Ni, Nc,
        lnRR_value, Var_lnRR, SE_lnRR
      )
    panel_funnels[[role]] <- funnel_object
    f5_publication_funnel_objects[[funnel_object$funnel_panel]] <- funnel_object
  }

  pair_stem <- paste0("F5",spec$Panel,"_funnels")
  save_f5_funnel_collection(
    panel_funnels[c("nutrient", "defense")],
    pair_stem, funnel_dir,
    nrow = 1, ncol = 2,
    width = funnel_pair_pdf_w,
    height = funnel_pair_pdf_h,
    compact = FALSE
  )
}

f5_funnel_expected_order <- as.vector(rbind(
  paste0(f5_publication_plan$Panel, "1"),
  paste0(f5_publication_plan$Panel, "2")
))
if (!identical(names(f5_publication_funnel_objects), f5_funnel_expected_order)) {
  stop("Figure 5 funnel order self-check failed.")
}

save_f5_funnel_collection(
  f5_publication_funnel_objects[f5_funnel_expected_order],
  "SF6",
  funnel_dir,
  nrow = 4, ncol = 3,
  width = funnel_grid_pdf_w,
  height = funnel_grid_pdf_h,
  compact = TRUE
)

f5_publication_funnel_summary <-
  dplyr::bind_rows(f5_publication_funnel_summary_rows) %>%
  dplyr::left_join(
    f5_publication_plan %>%
      dplyr::transmute(
        Panel,
        PairedPdfFile = file.path(
          "F5_funnel", paste0("F5",Panel,"_funnels.pdf")
        ),
        PairedTiffFile = sub("\\.pdf$", ".tiff", PairedPdfFile)
      ),
    by = "Panel"
  ) %>%
  dplyr::mutate(
    CombinedPdfFile = file.path(
      "F5_funnel", "SF6.pdf"
    ),
    CombinedTiffFile = file.path(
      "F5_funnel", "SF6.tiff"
    )
  ) %>%
  dplyr::arrange(PanelOrder, RoleOrder)
f5_publication_funnel_effect_data <-
  dplyr::bind_rows(f5_publication_funnel_effect_rows) %>%
  dplyr::mutate(
    PanelOrder = match(Panel, LETTERS),
    RoleOrder = ifelse(Role == "Nutrient", 1L, 2L)
  ) %>%
  dplyr::arrange(PanelOrder, RoleOrder, StudyID, SystemID, RowID) %>%
  dplyr::select(-PanelOrder, -RoleOrder)

publication_band_audit <- dplyr::bind_rows(
  f4_publication_band_audit,
  f5_publication_band_audit
)
if (any(!publication_band_audit$FiniteAndOrdered)) {
  stop("Publication confidence-band audit failed.")
}

## Corrected version of the four panels used in the submitted Figure 5.
## This is a comparison figure only; final retention must follow the rescreen table.
ranked_rows <- f5_selected %>%
  dplyr::filter(
    Method == primary_method, Dataset == "RAW",
    ReplicationEligible, EvidenceClass %in% c("main_evidence_candidate", "supportive_only")
  ) %>%
  dplyr::arrange(EvidenceRank, dplyr::desc(R2_weighted)) %>%
  dplyr::slice_head(n = 4)

if (isTRUE(export_diagnostic_figures)) {
diagnostic_grid_dir <- file.path(out_dir, "M2_diagnostic_figures", "Figure5_candidate_grids")
dir.create(diagnostic_grid_dir, recursive = TRUE, showWarnings = FALSE)
for (method in methods_to_run) {
  comparison_keys <- paste(method, "RAW", "C", original_f5_traits, sep = "__")
  available <- comparison_keys %in% names(f5_plot_store)
  if (all(available)) {
    comparison_plots <- f5_plot_store[comparison_keys]
    comparison_plots <- Map(
      function(p, tag) p + ggplot2::labs(tag = tag),
      comparison_plots,
      LETTERS[seq_along(comparison_plots)]
    )
    comparison_grid <-
      (comparison_plots[[1]] | comparison_plots[[2]]) /
      (comparison_plots[[3]] | comparison_plots[[4]])
    save_plot_pair(
      comparison_grid,
      paste("Figure5_SUBMITTED_four_after_correction", method, "RAW", sep = "_"),
      diagnostic_grid_dir,
      width = f5_grid_pdf_w,
      height = f5_grid_pdf_h
    )
  }
}

## A ranked candidate overview for the primary method. It is deliberately
## labelled CANDIDATES rather than FINAL because scientific panel selection
## still requires inspection of replication and sensitivity columns.
ranked_rows <- f5_selected %>%
  dplyr::filter(
    Method == primary_method, Dataset == "RAW",
    ReplicationEligible, EvidenceClass %in% c("main_evidence_candidate", "supportive_only")
  ) %>%
  dplyr::arrange(EvidenceRank, dplyr::desc(R2_weighted)) %>%
  dplyr::slice_head(n = 4)
if (nrow(ranked_rows)) {
  ranked_keys <- paste(
    ranked_rows$Method, ranked_rows$Dataset,
    ranked_rows$Nutrient, ranked_rows$DefenseTrait, sep = "__"
  )
  ranked_plots <- f5_plot_store[ranked_keys]
  ranked_plots <- ranked_plots[!vapply(ranked_plots, is.null, logical(1))]
  ranked_plots <- Map(
    function(p, tag) p + ggplot2::labs(tag = tag),
    ranked_plots,
    LETTERS[seq_along(ranked_plots)]
  )
  ranked_grid <- patchwork::wrap_plots(ranked_plots, ncol = 2)
  save_plot_pair(
    ranked_grid,
    paste("Figure5_TOP_ranked_CANDIDATES", primary_method, "RAW", sep = "_"),
    diagnostic_grid_dir,
    width = f5_grid_pdf_w,
    height = max(pdf_h, 5.4 * ceiling(length(ranked_plots) / 2))
  )
}

## Overview grids are generated after selection; no candidate is silently discarded.
for (method in methods_to_run) {
  for (dataset in c("RAW", "IQR_removed")) {
    rows_now <- f5_selected %>%
      dplyr::filter(Method == method, Dataset == dataset) %>%
      dplyr::arrange(Nutrient, DefenseTrait)
    if (!nrow(rows_now)) next
    keys <- paste(method, dataset, rows_now$Nutrient, rows_now$DefenseTrait, sep = "__")
    plots <- f5_plot_store[keys]
    plots <- plots[!vapply(plots, is.null, logical(1))]
    if (length(plots)) {
      grid <- patchwork::wrap_plots(plots, ncol = 2)
      save_plot_pair(
        grid,
        paste("Figure5_ALL_candidates", method, dataset, sep = "_"),
        diagnostic_grid_dir,
        width = f5_grid_pdf_w,
        height = max(pdf_h, 5.4 * ceiling(length(plots) / 2)),
        save_tiff = FALSE
      )
    }

    if (dataset == "RAW") {
      main_rows <- rows_now %>% dplyr::filter(MainEvidenceCandidate)
      if (nrow(main_rows)) {
        main_keys <- paste(method, dataset, main_rows$Nutrient, main_rows$DefenseTrait, sep = "__")
        main_plots <- f5_plot_store[main_keys]
        main_plots <- main_plots[!vapply(main_plots, is.null, logical(1))]
        main_grid <- patchwork::wrap_plots(main_plots, ncol = 2)
        save_plot_pair(
          main_grid,
          paste("Figure5_MAIN_evidence_candidates", method, dataset, sep = "_"),
          diagnostic_grid_dir,
          width = f5_grid_pdf_w,
          height = max(pdf_h, 5.4 * ceiling(length(main_plots) / 2))
        )
      }
    }
  }
}
}

# =========================
# 10) Mandatory self-checks
# =========================
all_points <- dplyr::bind_rows(
  f4_point_data %>% dplyr::mutate(Source = "F4"),
  f5_point_data %>% dplyr::mutate(Source = "F5")
)

sign_check <- all_points %>%
  dplyr::group_by(Source, Method) %>%
  dplyr::summarise(
    bad_sign_x = sum(X_value_raw != 0 & sign(X_value_raw) != sign(zX)),
    bad_sign_y = sum(Y_value_raw != 0 & sign(Y_value_raw) != sign(zY)),
    bad_zero_x = sum(X_value_raw == 0 & zX != 0),
    bad_zero_y = sum(Y_value_raw == 0 & zY != 0),
    .groups = "drop"
  )
if (any(sign_check$bad_sign_x + sign_check$bad_sign_y + sign_check$bad_zero_x + sign_check$bad_zero_y > 0)) {
  stop("Self-check failed: a corrected method changed at least one source sign or zero.")
}

numeric_bad <- all_points %>%
  dplyr::summarise(
    n = sum(
      !is.finite(zX) | !is.finite(zY) | !is.finite(Var_zX) |
        !is.finite(vi) | !is.finite(w_for_model) |
        Var_zX < 0 | vi <= 0 | w_for_model <= 0
    )
  ) %>%
  dplyr::pull(n)
if (numeric_bad != 0) stop("Self-check failed: invalid transformed value, variance, or weight.")

f4_duplicate_pairs <- f4_point_data %>%
  dplyr::count(Method, Panel, Dataset, StudyID, SystemID, Y_RowID, name = "n") %>%
  dplyr::filter(n > 1)
if (nrow(f4_duplicate_pairs)) stop("Self-check failed: duplicated Figure 4 paired rows.")

f5_duplicate_pairs <- f5_point_data %>%
  dplyr::count(Method, Nutrient, DefenseTrait, Dataset, StudyID, SystemID, Y_RowID, name = "n") %>%
  dplyr::filter(n > 1)
if (nrow(f5_duplicate_pairs)) stop("Self-check failed: duplicated Figure 5 paired rows.")

f4_expected <- tidyr::expand_grid(
  Method = methods_to_run,
  Dataset = c("RAW", "IQR_removed"),
  Panel = c("A", "B", "C", "D")
)
f4_missing <- f4_expected %>%
  dplyr::anti_join(f4_summary, by = c("Method", "Dataset", "Panel"))
if (nrow(f4_missing)) stop("Self-check failed: one or more Figure 4 outputs are missing.")

## Confirm that both tagged combined figures and all untagged standalone
## Figure 4/Figure 5 panels were written in PDF and TIFF formats.
publication_pair_files <- function(stem) {
 loc <- main_plot_location(stem,publication_dir)
 c(file.path(loc$directory,paste0(loc$stem,".pdf")),file.path(loc$directory,paste0(loc$stem,".tiff")))
}
f4_standalone_stems <- vapply(seq_len(nrow(f4_plan)), function(ii) {
  target <- f4_plan[ii, ]
  key <- paste(primary_method, "RAW", target$Panel, sep = "__")
  paste0(
    "Figure4_panel_", target$Panel, "_FINAL_M2_RAW_",
    target$Nutrient, "_vs_", safe_file_label(target$ResponseLabel), "_",
    f4_fit_store[[key]]$form, "_model_based_95CI"
  )
}, character(1))
f5_standalone_stems <- vapply(seq_len(nrow(f5_publication_plan)), function(ii) {
  spec <- f5_publication_plan[ii, ]
  paste0(
    "Figure5_panel_", spec$Panel, "_FINAL_M2_RAW_",
    spec$Nutrient, "_vs_", safe_file_label(spec$DefenseTrait), "_",
    spec$RequestedModel, "_model_based_95CI"
  )
}, character(1))
publication_figure_expected_files <- c(
  publication_pair_files("Figure4_FINAL_M2_RAW_model_based_95CI"),
  unlist(lapply(f4_standalone_stems, publication_pair_files), use.names = FALSE),
  publication_pair_files("Figure5_FINAL_six_panels_M2_RAW_model_based_95CI"),
  unlist(lapply(f5_standalone_stems, publication_pair_files), use.names = FALSE)
)
publication_figure_missing_files <- publication_figure_expected_files[
  !file.exists(publication_figure_expected_files)
]
if (length(publication_figure_missing_files)) {
  stop(
    "Self-check failed: one or more final Figure 4/Figure 5 files are missing: ",
    paste(basename(publication_figure_missing_files), collapse = "; ")
  )
}

raw_keys <- all_points %>%
  dplyr::filter(Dataset == "RAW") %>%
  dplyr::distinct(Source, Method, Panel, Nutrient, DefenseTrait, StudyID, SystemID, Y_RowID)
iqr_keys <- all_points %>%
  dplyr::filter(Dataset == "IQR_removed") %>%
  dplyr::distinct(Source, Method, Panel, Nutrient, DefenseTrait, StudyID, SystemID, Y_RowID)
extra_iqr <- iqr_keys %>%
  dplyr::anti_join(
    raw_keys,
    by = intersect(names(iqr_keys), names(raw_keys))
  )
if (nrow(extra_iqr)) stop("Self-check failed: IQR data are not a subset of RAW data.")

f5_funnel_count_check <- f5_publication_funnel_effect_data %>%
  dplyr::count(FunnelPanel, name = "ObservedEffectRows") %>%
  dplyr::left_join(
    f5_publication_funnel_summary %>%
      dplyr::select(FunnelPanel, ExpectedEffectRows = n_effect_sizes),
    by = "FunnelPanel"
  ) %>%
  dplyr::mutate(CountMatches = ObservedEffectRows == ExpectedEffectRows)
if (nrow(f5_publication_funnel_summary) != 12L ||
    !identical(
      as.character(f5_publication_funnel_summary$FunnelPanel),
      f5_funnel_expected_order
    )) {
  stop("Self-check failed: Figure 5 funnel summary is incomplete or out of order.")
}
if (nrow(f5_funnel_count_check) != 12L || any(!f5_funnel_count_check$CountMatches)) {
  stop("Self-check failed: Figure 5 funnel effect-row counts do not match.")
}

f5_funnel_independent_files <- unlist(lapply(
  f5_publication_funnel_summary$IndependentPdfFile,
  function(path_now) c(
    file.path(out_dir, path_now),
    sub("\\.pdf$", ".tiff", file.path(out_dir, path_now))
  )
))
f5_funnel_pair_stems <- paste0("F5",f5_publication_plan$Panel,"_funnels")
f5_funnel_pair_files <- unlist(lapply(
  f5_funnel_pair_stems,
  function(stem) c(
    file.path(funnel_dir, paste0(stem, ".pdf")),
    file.path(funnel_dir, paste0(stem, ".tiff"))
  )
))
f5_funnel_combined_files <- c(
  file.path(funnel_dir,"SF6.pdf"), file.path(funnel_dir,"SF6.tiff")
)
f5_funnel_expected_files <- c(
  f5_funnel_independent_files,
  f5_funnel_pair_files,
  f5_funnel_combined_files
)
f5_funnel_path_lengths <- nchar(
  normalizePath(f5_funnel_expected_files, winslash = "/", mustWork = FALSE),
  type = "chars"
)
max_f5_funnel_path_chars <- max(f5_funnel_path_lengths, na.rm = TRUE)
if (
  .Platform$OS.type == "windows" &&
    is.finite(max_f5_funnel_path_chars) &&
    max_f5_funnel_path_chars > windows_safe_path_limit
) {
  stop(
    "Self-check failed: planned Figure 5 funnel path exceeds the Windows-safe limit; ",
    "maximum = ", max_f5_funnel_path_chars, " characters."
  )
}
f5_funnel_missing_files <- f5_funnel_expected_files[
  !file.exists(f5_funnel_expected_files)
]
if (length(f5_funnel_missing_files)) {
  stop(
    "Self-check failed: one or more Figure 5 funnel files are missing: ",
    paste(basename(f5_funnel_missing_files), collapse = "; ")
  )
}

self_checks <- dplyr::bind_rows(
  sign_check %>%
    dplyr::transmute(
      Check = paste(Source, Method, "sign_and_zero_preservation", sep = ":"),
      Status = "PASS",
      Observed = paste0(
        "bad_sign_x=", bad_sign_x, ";bad_sign_y=", bad_sign_y,
        ";bad_zero_x=", bad_zero_x, ";bad_zero_y=", bad_zero_y
      )
    ),
  tibble::tibble(
    Check = c(
      "finite_values_variances_weights",
      "unique_F4_paired_rows",
      "unique_F5_paired_rows",
      "all_8_M2_F4_results_present",
      "IQR_subset_of_RAW",
      "F5_twelve_funnels_complete",
      "F5_funnel_effect_counts_match",
      "F5_funnel_files_created",
      "F5_funnel_paths_Windows_safe",
      "F5_forest_heterogeneity_complete",
      "F5_forest_components_sum_to_total",
      "Table_S3_Figure5_has_12_rows",
      "legacy_paper_R2"
    ),
    Status = c(
      "PASS", "PASS", "PASS", "PASS", "PASS",
      "PASS", "PASS", "PASS", "PASS", "PASS", "PASS", "PASS",
      "NOT_A_STOPPING_CHECK"
    ),
    Observed = c(
      paste0("bad_rows=", numeric_bad),
      paste0("duplicates=", nrow(f4_duplicate_pairs)),
      paste0("duplicates=", nrow(f5_duplicate_pairs)),
      paste0("missing=", nrow(f4_missing)),
      paste0("extra_IQR_keys=", nrow(extra_iqr)),
      paste0("funnels=", nrow(f5_publication_funnel_summary), "/12"),
      paste0("mismatches=", sum(!f5_funnel_count_check$CountMatches)),
      paste0("missing_files=", length(f5_funnel_missing_files)),
      paste0(
        "max_chars=", max_f5_funnel_path_chars,
        ";limit=", windows_safe_path_limit
      ),
      paste0(
        "rows=", nrow(f5_forest_summary),
        ";invalid=", nrow(forest_heterogeneity_bad),
        ";columns=SE,tau2,I2"
      ),
      paste0("component_mismatches=", nrow(forest_component_bad)),
      paste0("rows=", nrow(f5_publication_forest_table), "/12"),
      "old erroneous transformation is excluded from corrected analysis"
    )
  )
)

self_checks <- dplyr::bind_rows(
  self_checks,
  tibble::tibble(
    Check = c(
      "publication_confidence_bands",
      "F5_final_six_candidates_available",
      "F5_requested_shapes_match_automatic"
    ),
    Status = c(
      "PASS",
      "PASS",
      ifelse(
        all(f5_publication_selection$RequestedMatchesAutomatic),
        "PASS", "REVIEW"
      )
    ),
    Observed = c(
      paste0(
        "bands=", nrow(publication_band_audit),
        ";invalid=", sum(!publication_band_audit$FiniteAndOrdered),
        ";method=", confidence_band_method
      ),
      paste0(
        "available=", sum(f5_publication_selection$CandidateAvailable),
        "/", nrow(f5_publication_selection)
      ),
      paste0(
        "matched=", sum(f5_publication_selection$RequestedMatchesAutomatic),
        "/", nrow(f5_publication_selection)
      )
    )
  )
)

f5_submitted_four_corrected <- f5_selected %>%
  dplyr::filter(
    Dataset == "RAW", Nutrient == "C",
    DefenseTrait %in% original_f5_traits
  ) %>%
  dplyr::mutate(
    SubmittedPanel = factor(
      DefenseTrait,
      levels = original_f5_traits,
      labels = LETTERS[seq_along(original_f5_traits)]
    )
  ) %>%
  dplyr::arrange(Method, SubmittedPanel)

# =========================
# 11) Export tables
# =========================
config <- tibble::tibble(
  Item = c(
    "input_file", "sheet_name", "output_directory", "methods", "primary_method",
    "iqr_k", "F4_primary_model", "shape_rule", "F5_main_evidence_rule",
    "displayed_confidence_band", "CR2_rule", "F5_publication_panels",
    "publication_figure_scope", "F5_funnel_rule",
    "F5_forest_heterogeneity_rule", "windows_path_rule"
  ),
  Value = c(
    input_file, as.character(sheet_name), out_dir,
    paste(methods_to_run, collapse = ";"), primary_method, as.character(iqr_k),
    "M2 pooled inverse-variance WLS; factor(YTrait) and balanced weights are sensitivity analyses",
    "quadratic only if n>=10; SystemID>=5; DeltaAICc>=2; DeltaR2>=0.03; p(b2)<=0.05; vertex within 5th-95th percentile of x; otherwise linear",
    "n>=10; StudyID>=3; SystemID>=5; R2>=0.10; BH-CR2<=0.05; LOSO>0; study-balanced and RAW-IQR sensitivity supported",
    "model-based 95% confidence interval for the fitted WLS mean; full observed x range; no data removal",
    "clubSandwich CR2 on lm, clustered by StudyID; retained for coefficient inference and evidence grading; not used as the displayed ribbon",
    paste0(
      f5_publication_plan$Panel, ":", f5_publication_plan$Nutrient, "-",
      f5_publication_plan$DefenseTrait, "(", f5_publication_plan$RequestedModel, ")",
      collapse = ";"
    ),
    "only M2-RAW Figure 4, six M2-RAW Figure 5 panels, and matching Figure 5 funnel outputs are written; M2 RAW/IQR numerical results remain in tables",
    "two lnRR funnels per final Figure 5 relationship from the exact regression RowID sets: nutrient first, defense trait second; 12 independent, 6 paired and one 12-panel grid",
    "companion forests retain rma.mv REML with StudyID/SystemID random intercepts; total tau2=sum(sigma2); total I2=100*total_tau2/(total_tau2+typical sampling variance); component tau2 and I2 are also exported",
    "short output directory and compact funnel filenames; every graphics path is checked against a conservative 240-character Windows limit before the device is opened"
  )
)

# Per-figure result workbooks retain full-precision model and sensitivity tables.
current_stage <- "table export"
source_rows <- function(x) {
 if("Y_RowID" %in% names(x)) x$SourceExcelRow_Y <- x$Y_RowID+input_header_row
 if("X_RowIDs" %in% names(x)) x$SourceExcelRows_X <- vapply(strsplit(as.character(x$X_RowIDs),";",fixed=TRUE),
   function(ids) paste(as.integer(ids)+input_header_row,collapse=";"),character(1))
 x
}
f4_point_data <- source_rows(f4_point_data)
f5_point_data <- source_rows(f5_point_data)
f4_plot_data <- f4_point_data %>% dplyr::filter(Method==primary_method,Dataset=="RAW")
f5_plot_data <- f5_point_data %>%
 dplyr::filter(Method==primary_method,Dataset=="RAW") %>%
 dplyr::inner_join(f5_publication_plan %>% dplyr::select(Panel,Nutrient,DefenseTrait),by=c("Nutrient","DefenseTrait"))
f5_publication_funnel_effect_data$SourceExcelRow <- f5_publication_funnel_effect_data$RowID+input_header_row
input_audit <- input_audit %>% dplyr::filter(Item!="legacy_R2_check")
input_audit <- dplyr::bind_rows(input_audit,tibble::tibble(Item="header_row",Value=as.character(input_header_row)))
self_checks <- self_checks %>% dplyr::filter(Check!="legacy_paper_R2")
config <- dplyr::bind_rows(config,tibble::tibble(
 Item=c("output_structure","contour_key","forest_vs_funnel_models"),
 Value=c("F4; F5; F5_funnel. Stable filenames are replaced on rerun.",
 "Two-sided normal-theory contours about lnRR=0. Original fills: white p>=.10; grey90 .05<=p<.10; grey80 p<.05. The p=.01 contour line remains, with grey80 on both sides.",
 "Forests use original multilevel rma.mv REML; funnel asymmetry diagnostics use original rma.uni REML with FE fallback.")))
f4_prediction <- purrr::map_dfr(seq_len(nrow(f4_plan)),function(i) {
 s<-f4_plan[i,];key<-paste(primary_method,"RAW",s$Panel,sep="__")
 make_prediction_band(f4_points[[key]],f4_fit_store[[key]]) %>% dplyr::mutate(Panel=s$Panel,.before=1)
})
f5_prediction <- purrr::map_dfr(seq_len(nrow(f5_publication_plan)),function(i) {
 s<-f5_publication_plan[i,];key<-paste(primary_method,"RAW",s$Nutrient,s$DefenseTrait,sep="__")
 df<-f5_point_rows[[key]]
 make_prediction_band(df,fit_wls(df,s$RequestedModel)) %>% dplyr::mutate(Panel=s$Panel,.before=1)
})
safe_write_xlsx(list(
 F4_Model_Summary=f4_summary,
 F4_Plot_Data=f4_plot_data,
 Prediction_Data=f4_prediction,
 F4_All_Models=f4_all_models,
 F4_Weight_Sensitivity=f4_weight_sensitivity,
 F4_Transformed_Data=f4_point_data,
 Confidence_Band_Audit=f4_publication_band_audit,
 Config=config,Input_Audit=input_audit,Input_QC_Flags=input_qc_flags,Self_Checks=self_checks
),file.path(out_dir,"F4","F4_results.xlsx"))
safe_write_xlsx(list(
 F5_Publication_Models=f5_publication_model_results,
 F5_Plot_Data=f5_plot_data,
 Table_S3_Figure5_M2_RAW=f5_publication_forest_table,
 Prediction_Data=f5_prediction,
 F5_Publication_Selection=f5_publication_selection,
 Confidence_Band_Audit=f5_publication_band_audit,
 F5_Candidate_Availability=f5_availability,
 F5_All_Models=f5_all_models,
 F5_Selected_Shape=f5_selected,
 F5_Submitted4_Corrected=f5_submitted_four_corrected,
 F5_Ranked_Candidates=ranked_rows,
 F5_Weight_Sensitivity=f5_sensitivity,
 F5_RAW_IQR_Shape_Stability=shape_stability,
 F5_Companion_Forest_Summary=f5_forest_summary,
 F5_Transformed_Data=f5_point_data,
 Config=config,Input_Audit=input_audit,Input_QC_Flags=input_qc_flags,Self_Checks=self_checks
),file.path(out_dir,"F5","F5_results.xlsx"))
funnel_methods <- tibble::tibble(
 Item=c("comparison","source_rows","models","Egger","Begg","trimfill","counts","SF6","contours","missing_values"),
 Description=c("lnRR=log(introduced-range/native-range). Regression Cha is not used in these plots.",
 "Exact original nutrient/defense RowIDs from each F5 M2-RAW paired dataset, restricted to valid positive lnRR variance.",
 "Diagnostic model: rma.uni REML, FE fallback. Pooled columns: original companion-forest multilevel REML with its original fallback.",
 "metafor::regtest(model, model='rma') for at least 10 effect-size records.",
 "metafor::ranktest(model, exact=FALSE) for at least 10 effect-size records.",
 "metafor::trimfill(model) for at least 10 effect-size records. k0 reported; observed effects plotted and primary fits unchanged.",
 "n=effect-size records; k=studies. Diagnostic tests do not apply StudyID-clustered CR2 correction.",
 "12 separate analyses in row-major order A1,A2,B1,B2,C1,C2,D1,D2,E1,E2,F1,F2; no re-pooling across traits.",
 "Original fills: white p>=.10; grey90 .05<=p<.10; grey80 p<.05. The p=.01 contour line remains, with grey80 on both sides. Two-sided normal-theory contours about lnRR=0.",
 "NA means a diagnostic was unavailable or could not be estimated, not a zero P value or zero missing effects."))
safe_write_xlsx(list(
 F5_Publication_Funnels=f5_publication_funnel_summary,
 F5_Funnel_Effect_Data=f5_publication_funnel_effect_data,
 F5_Funnel_Count_Check=f5_funnel_count_check,
 Methods=funnel_methods
),file.path(funnel_dir,"F5_funnel_results.xlsx"))
safe_write_csv(f4_summary,file.path(out_dir,"F4","F4_model_summary.csv"),row.names=FALSE,na="",fileEncoding="UTF-8")
safe_write_csv(f5_publication_model_results,file.path(out_dir,"F5","F5_model_summary.csv"),row.names=FALSE,na="",fileEncoding="UTF-8")
safe_write_csv(f5_publication_forest_table,file.path(out_dir,"F5","Table_S3.csv"),row.names=FALSE,na="",fileEncoding="UTF-8")
safe_write_csv(f5_publication_funnel_summary,file.path(funnel_dir,"F5_funnel_summary.csv"),row.names=FALSE,na="",fileEncoding="UTF-8")
manifest <- dplyr::bind_rows(
 f4_summary %>% dplyr::filter(Dataset=="RAW",Method==primary_method) %>%
   dplyr::transmute(Figure="F4",Panel=paste0("F4",Panel),Nutrient,Response=DefenseTrait,Dataset,Model=SelectedModel),
 f5_publication_model_results %>%
   dplyr::transmute(Figure="F5",Panel=paste0("F5",Panel),Nutrient,Response=DefenseTrait,Dataset,Model=RequestedModel)
) %>% dplyr::mutate(PdfFile=paste0(Figure,"/",Panel,".pdf"),TiffFile=paste0(Figure,"/",Panel,".tiff"),
 InputFile=basename(input_file),InputMD5=input_md5,HeaderRow=input_header_row)
safe_write_csv(manifest,file.path(out_dir,"panel_manifest.csv"),row.names=FALSE,na="",fileEncoding="UTF-8")
current_stage <- "final verification"
if(!all(file.exists(file.path(out_dir,expected_files)))) stop("An expected output file is missing.")
if(!identical(unname(tools::md5sum(input_file)),input_md5)) stop("Input workbook changed during this run.")
safe_write_lines(c("COMPLETE",paste0("Started: ",started_at),paste0("Finished: ",format(Sys.time(),"%Y-%m-%d %H:%M:%S %Z")),
 paste0("Input: ",input_file),paste0("Input location method: ",located$method),
 paste0("Input MD5: ",input_md5),paste0("Header row: ",input_header_row),
 paste0("Input data rows: ",nrow(raw)),paste0("Output: ",out_dir),
 "Input workbook was read only. Analysis uses introduced_vs_native_range.",
 "Outputs: F4 four panels+composite; F5 six panels+composite; 12 individual funnels+6 pairs+SF6.",
 "31 PDFs and 31 matching 600-dpi TIFFs. Stable files are replaced on rerun.",
 "Main figures use M2-RAW. IQR results remain numerical sensitivity analyses.",
 "F5 right panels display weighted R-squared; left forests test pooled lnRR.",
 "Reference SF6 layout: per-panel internal keys/diagnostics, five y ticks; original white/grey90/grey80 fills, all three contour boundaries and unclipped axes. Observed lnRR points and diagnostic formulas retained.",
 "", "Warnings captured:",if(length(captured_warnings)) captured_warnings else "No warnings captured.",
 "","Session information:",capture.output(sessionInfo())),info_path)
message("Completed F4, F5 and SF6. Output folder: ",out_dir)

},warning=function(w) {
 captured_warnings <<- c(captured_warnings,paste0("[",current_stage,"] ",conditionMessage(w)))
}),error=function(e) {
 original <- conditionMessage(e)
 tryCatch(safe_write_lines(c("FAILED",paste0("Started: ",started_at),paste0("Stage: ",current_stage),
 paste0("Original error: ",original),"Files may be from different attempts. Rerun successfully before using results.",captured_warnings),info_path),
 error=function(io) message("Could not save diagnostic: ",conditionMessage(io)))
 stop(original,"\nStage: ",current_stage,"\nOutput: ",out_dir,call.=FALSE)
})
})
