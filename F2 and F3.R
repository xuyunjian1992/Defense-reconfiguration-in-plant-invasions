############################################################
# F2/F3 analyses - standalone script with F3 funnels
# Updated 2026-10-08: simple folders and eight F3 funnels + SF5. Run the complete file.
# All analysis functions are included below; no R/ folder needed.
# F3: left forest = pooled-effect significance; right fit = R-squared only.
# No automatic software installation or data upload.
############################################################
local({
# 1) User configuration. These overrides are optional.
shared_file_name <- "Table S2.xlsx"
project_root_override <- NULL  # Example: "D:/my_project/F2_F3_v2_project"
input_file_override <- NULL    # Optional full path to the input .xlsx workbook
output_root_override <- NULL   # Optional short writable output folder
shared_sheet_name <- 1
# End of user configuration.

# Path discovery is optional. The local Windows input does not need a project.
# No external R/ directory or project file is required to load the analysis.
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
      "or run this file from its extracted project. No R/helpers.R file is needed.",call.=FALSE)
}

located <- resolve_input_location(shared_file_name,
 project_root_override=project_root_override,input_file_override=input_file_override)
project_root <- located$root
shared_input_file <- located$input
output_root <- file.path(project_root,"F2_F3_results")

required_packages <- c("readxl","dplyr","stringr","ggplot2","writexl","tibble","scales","patchwork","metafor")
missing <- required_packages[!vapply(required_packages,requireNamespace,logical(1),quietly=TRUE)]
if(length(missing)) stop("Missing packages: ",paste(missing,collapse=", "),
 ". Install them once using the command in README.md, then rerun. No packages were installed automatically.")
suppressPackageStartupMessages(library(dplyr))
# ---- Embedded helpers.R ----
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
  safe_write_lines("F2/F3 output write check", probe)
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


# Read only: detect the header, then normalize documented presentation aliases.
read_input_table <- function(input_file, sheet_name) {
 preview <- readxl::read_excel(input_file, sheet=sheet_name, col_names=FALSE,
   col_types="text", n_max=12, .name_repair="minimal")
 key <- function(x) gsub("[^a-z0-9]", "", tolower(trimws(as.character(x))))
 required_keys <- c("systemid","trait","mc","mi","sdc","sdi","nc","ni")
 header_candidates <- which(vapply(seq_len(nrow(preview)), function(i) {
   keys <- key(unlist(preview[i, ], use.names=FALSE))
   all(required_keys %in% keys) && any(c("studyid","number") %in% keys)
 }, logical(1)))
 if (length(header_candidates)!=1L) stop("Cannot uniquely detect the Excel header in the first 12 rows: ", input_file)
 h <- header_candidates[1]
 raw <- readxl::read_excel(input_file, sheet=sheet_name, skip=h-1L, .name_repair="minimal")
 nm <- trimws(names(raw))
 nm[is.na(nm)|nm==""] <- paste0("Unnamed_column_",which(is.na(nm)|nm==""))
 alias <- c("Study ID"="StudyID", "System ID"="SystemID")
 for (old in names(alias)) nm[nm==old] <- alias[[old]]
 if (anyDuplicated(nm)) stop("Duplicate column names after header normalization.")
 names(raw) <- nm
 if (!"Trait" %in% nm) stop("Trait column is missing.")
 # The spreadsheet remains untouched; only these exact normalized aliases change.
 tr <- tolower(trimws(as.character(raw$Trait)))
 nutrient_alias <- c("plant carbon"="C", "plant nitrogen"="N", "plant phosphorus"="P")
 for (old in names(nutrient_alias)) raw$Trait[!is.na(tr) & tr==old] <- nutrient_alias[[old]]
 observed <- c(header_row=h, data_rows=nrow(raw))
 if (!identical(unname(as.integer(observed)), unname(as.integer(c(input_info$header_row,input_info$data_rows)))))
   stop("Input changed between initial inspection and module read.")
 raw
}
inspect_input <- function(path, sheet) {
 p <- readxl::read_excel(path, sheet=sheet, col_names=FALSE, col_types="text", n_max=12, .name_repair="minimal")
 key <- function(x) gsub("[^a-z0-9]", "", tolower(trimws(as.character(x))))
 candidates <- which(vapply(seq_len(nrow(p)), function(i) { keys <- key(unlist(p[i,],use.names=FALSE)); all(c("systemid","trait","mc","mi","sdc","sdi","nc","ni") %in% keys) && any(c("studyid","number") %in% keys) }, logical(1)))
 if(length(candidates)!=1L) stop("Expected one recognizable header row in the first 12 Excel rows.")
 raw <- readxl::read_excel(path,sheet=sheet,skip=candidates-1L,.name_repair="minimal")
 list(header_row=as.integer(candidates), data_rows=nrow(raw), md5=unname(tools::md5sum(path)))
}
nutrient_axis_label <- function(trait) {
 label <- c(C="Plant carbon content", N="Plant nitrogen content", P="Plant phosphorus content")
 paste0("lnRR-", unname(label[trait]))
}
forest_trait_labels <- function(x) {
 x <- sub("^C\\n", "Plant carbon\ncontent\n", x)
 x <- sub("^N\\n", "Plant nitrogen\ncontent\n", x)
 sub("^P\\n", "Plant phosphorus\ncontent\n", x)
}
assert_system_study_keys <- function(base) {
 n <- tapply(base$StudyID, base$SystemID, function(x) length(unique(x)))
 if(any(n>1L)) stop("Some SystemID values span multiple StudyID values. Resolve the input pairing keys before fitting.")
}
validate_fit_input <- function(dat, spec) {
 if(nrow(dat) < ifelse(spec$ModelForm=="quadratic",6L,3L)) stop(spec$Panel, ": too few fitting rows.")
 if(anyDuplicated(dat$RegressionID)) stop(spec$Panel, ": duplicated RegressionID.")
 if(any(!is.finite(dat$zX)|!is.finite(dat$zY)|!is.finite(dat$w_for_model)|dat$w_for_model<=0))
   stop(spec$Panel, ": invalid fitted value/weight.")
}
model_summary_row <- function(spec, dat, fit, r2) {
 sm <- summary(fit); cf <- stats::coef(fit)
 get_coef <- function(term,default=NA_real_) if(term %in% names(cf)) unname(cf[term]) else default
 get_p <- function(term) if(term %in% rownames(sm$coefficients)) unname(sm$coefficients[term,4]) else NA_real_
 tibble::tibble(Panel=spec$Panel, AnalysisID=spec$AnalysisID, Dataset=spec$Dataset,
  Metric="LNRR", ComparisonType="invasive_vs_native", XTrait=spec$XTrait,
  Method="WLS", ModelForm=spec$ModelForm, n_rows=nrow(dat),
  n_studies=dplyr::n_distinct(dat$StudyID), n_systems=dplyr::n_distinct(dat$SystemID),
  n_response_traits=dplyr::n_distinct(dat$YTrait),
  Intercept=get_coef("(Intercept)"), Slope_linear=get_coef("zX"), Slope_quadratic=get_coef("I(zX^2)"),
  P_linear=get_p("zX"), P_quadratic=get_p("I(zX^2)"), R2=r2, AIC=stats::AIC(fit),
  Residual_df=fit$df.residual, OriginalSelectionStatus="fixed_quadratic_in_original_F2")
}
panel_result <- function(spec, dat_raw, dat, fit, band, plot, summary) {
 dat$Panel <- spec$Panel
 dat$SourceExcelRow_Y <- dat$Y_RowID + input_info$header_row
 dat$SourceExcelRows_X <- vapply(strsplit(dat$X_RowIDs,";",fixed=TRUE),
   function(x) paste(as.integer(x)+input_info$header_row,collapse=";"),character(1))
 qx <- quantile(dat_raw$zX,c(.25,.75)); qy <- quantile(dat_raw$zY,c(.25,.75))
 bounds <- c(qx[1]-1.5*diff(qx), qx[2]+1.5*diff(qx), qy[1]-1.5*diff(qy), qy[2]+1.5*diff(qy))
 audit <- tibble::tibble(Panel=spec$Panel, Dataset=spec$Dataset,
   RegressionID=dat_raw$RegressionID, StudyID=dat_raw$StudyID, SystemID=dat_raw$SystemID,
   Y_RowID=dat_raw$Y_RowID, X_RowIDs=dat_raw$X_RowIDs,
   Retained=dat_raw$RegressionID %in% dat$RegressionID,
   Rule=ifelse(spec$Dataset=="RAW","No IQR removal","Both axes within Q1/Q3 +/- 1.5 IQR; inclusive"),
   X_lower=if(spec$Dataset=="RAW") NA_real_ else bounds[1],
   X_upper=if(spec$Dataset=="RAW") NA_real_ else bounds[2],
   Y_lower=if(spec$Dataset=="RAW") NA_real_ else bounds[3],
   Y_upper=if(spec$Dataset=="RAW") NA_real_ else bounds[4])
 summary$n_before_IQR <- nrow(dat_raw)
 summary$n_removed_IQR <- nrow(dat_raw)-nrow(dat)
 summary$PdfFile <- paste0(spec$Panel,".pdf")
 band$Panel <- spec$Panel
 band$Dataset <- spec$Dataset
 cf <- summary(fit)$coefficients
 ci <- stats::confint(fit)
 coefficients <- data.frame(Panel=spec$Panel, Term=rownames(cf), Estimate=cf[,1],
   SE=cf[,2], t_value=cf[,3], P_value=cf[,4], CI_lower=ci[,1], CI_upper=ci[,2], row.names=NULL)
 list(spec=spec, data=dat, summary=summary, plot=plot, band=band,
      audit=audit, coefficients=coefficients, model=fit)
}
# Complete figures only. F3's panel is already forest+regression.
export_figure <- function(items, figure, out_dir) {
 d <- file.path(out_dir,figure); ensure_output_dir(d)
 plots <- vector("list",4L)
 for(i in seq_along(items)) {
   item <- items[[i]]
   p <- item$plot
   if(figure=="F3") p <- p + patchwork::plot_annotation(title=LETTERS[i],
     theme=ggplot2::theme(plot.title=ggplot2::element_text(size=22,face="bold",hjust=0)))
   else p <- p + ggplot2::labs(tag=LETTERS[i]) + ggplot2::theme(plot.tag=ggplot2::element_text(size=22,face="bold"))
   plots[[i]] <- p
   width <- if(figure=="F2") 6.8 else 11.6
   base <- file.path(d,item$spec$Panel)
   save_figure_file(p,paste0(base,".pdf"),width,5.6)
   save_figure_file(p,paste0(base,".png"),width,5.6)
 }
 # Freeze each compound F3 panel as a grob to keep its letter and forest together.
 if(figure=="F3") plots <- lapply(plots,function(p) patchwork::wrap_elements(full=patchwork::patchworkGrob(p)))
 combined <- patchwork::wrap_plots(plots,ncol=2,nrow=2)
 width <- if(figure=="F2") 13.6 else 23.2
 save_figure_file(combined,file.path(d,paste0(figure,".pdf")),width,11.2)
 save_figure_file(combined,file.path(d,paste0(figure,".png")),width,11.2)
 bind <- function(field) dplyr::bind_rows(lapply(items,`[[`,field))
 book <- list(Model_Summary=bind("summary"), Prepared_Fit_Data=bind("data"),
   Coefficients=bind("coefficients"), Prediction_Data=bind("band"))
 if(figure=="F3") book <- c(book,list(Forest_Summary=bind("forest_summary"),
   Forest_Source_Rows=bind("forest_data"), Selection_Audit=bind("selection")))
 book <- c(book,list(IQR_Audit=bind("audit")))
 safe_write_xlsx(book,file.path(d,paste0(figure,"_results.xlsx")))
 safe_write_csv(book$Model_Summary,file.path(d,paste0(figure,"_model_summary.csv")),row.names=FALSE,na="",fileEncoding="UTF-8")
 invisible(book$Model_Summary)
}
save_figure_file <- function(p,path,width,height) {
 if(tolower(tools::file_ext(path))=="pdf") {
   device <- if(capabilities("cairo")) grDevices::cairo_pdf else grDevices::pdf
   safe_ggsave(path,plot=p,width=width,height=height,device=device,bg="white",limitsize=FALSE)
 } else {
   safe_ggsave(path,plot=p,width=width,height=height,device="png",dpi=300,bg="white",limitsize=FALSE)
 }
}


# Publication-bias diagnostics use each F3 forest's original lnRR rows.
# A combined funnel figure is an assembly of eight separate analyses.
funnel_trait_label <- function(trait) {
 labels <- c(C="Plant carbon content",N="Plant nitrogen content",P="Plant phosphorus content")
 if(trait %in% names(labels)) unname(labels[trait]) else trait
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

draw_f3_funnel <- function(obj, spec, trait, compact=FALSE) {
 draw_manuscript_funnel(obj$model,
   title=paste0("lnRR ", funnel_trait_label(spec$XTrait), " VS. ", spec$YTrait, "-", funnel_trait_label(trait)),
   egger_p=obj$egger_p, begg_p=obj$begg_p, trimfill_k0=obj$trimfill_k0,
   n_effects=obj$n_effect_sizes, n_studies=obj$k_studies,
   compact=compact, p_digits=4L)
}
save_funnel_graphic <- function(path,width,height,draw) {
 open <- function(temp) {
   if(tolower(tools::file_ext(path))=="pdf") {
     grDevices::pdf(temp,width=width,height=height,family="sans",pointsize=12)
   } else grDevices::png(temp,width=width,height=height,units="in",res=300,pointsize=12,bg="white")
 }
 safe_plot_device(path,open,draw)
}
export_f3_funnels <- function(items,out_dir) {
 d <- file.path(out_dir,"F3_funnel");ensure_output_dir(d)
 panel_objects <- list();summaries <- list();rows <- list();diagnostics <- list()
 counter <- 0L
 for(item in items) {
   for(trait in c(item$spec$XTrait,item$spec$YTrait)) {
     obj <- item$funnels[[trait]]
     if(is.null(obj)) stop(item$spec$Panel,": funnel unavailable for ",trait)
     counter <- counter+1L
     key <- paste0(item$spec$Panel,"_",trait,"_funnel")
     panel_objects[[counter]] <- list(object=obj,spec=item$spec,trait=trait)
     draw_one <- local({x<-obj;s<-item$spec;t<-trait;function() {
       manuscript_funnel_par()
       draw_f3_funnel(x,s,t)
     }})
     for(ext in c("pdf","png")) save_funnel_graphic(file.path(d,paste0(key,".",ext)),6.4,5.0,draw_one)
     summaries[[counter]] <- cbind(data.frame(Panel=item$spec$Panel,Dataset=item$spec$Dataset,
       Trait=trait,Role=ifelse(trait==item$spec$XTrait,"nutrient","defense"),
       Egger_P=obj$egger_p,Begg_P=obj$begg_p,Trimfill_k0=obj$trimfill_k0,
       n_effect_sizes=obj$n_effect_sizes,k_studies=obj$k_studies,
       FunnelModelMethod=obj$model$method,PdfFile=paste0(key,".pdf"),
       DiagnosticNote=paste(obj$warnings,collapse=" | ")),
       as.data.frame(obj$pooled))
     dat <- obj$data
     dat$Panel <- item$spec$Panel;dat$Dataset <- item$spec$Dataset
     dat$SourceExcelRow <- dat$RowID+input_info$header_row
     rows[[counter]] <- dat
     diagnostics[[counter]] <- c(paste0("## ",key),capture.output(print(obj$model)),
       paste0("Egger P: ",format(obj$egger_p,digits=17)),paste0("Begg P: ",format(obj$begg_p,digits=17)),
       paste0("Trim-and-fill k0: ",obj$trimfill_k0),
       if(length(obj$warnings)) obj$warnings else "No funnel diagnostic warnings.","")
   }
 }
 stopifnot(counter==8L)
 draw_combined <- function() {
   manuscript_funnel_par(3L,3L,TRUE)
   for(entry in panel_objects) draw_f3_funnel(entry$object,entry$spec,entry$trait,compact=TRUE)
   graphics::plot.new() # Ninth cell is intentionally blank, as in the supplied SF5.
 }
 for(ext in c("pdf","png")) save_funnel_graphic(file.path(d,paste0("SF5.",ext)),12,9,draw_combined)
 summary <- dplyr::bind_rows(summaries)
 data <- dplyr::bind_rows(rows)
 notes <- data.frame(Item=c("source_rows","funnel_models","Egger","Begg","trim_and_fill","combined_figure","contours"),
   Description=c("Same original RowIDs as each corresponding F3 forest trait; positive finite lnRR variance.",
    "Original REML rma.uni; original FE fallback if REML cannot fit.",
    "metafor::regtest(model, model='rma'), for n >= 10 effect-size rows.",
    "metafor::ranktest(model), for n >= 10; ties can yield an approximate P value and are recorded.",
    "metafor::trimfill(model), for n >= 10. k0 is reported; imputed points are not added to the funnel display or forest.",
    "SF5 assembles eight separate funnel plots in original order; it does not pool different F3 relationships.",
    "Two-sided normal-theory P contours about zero. Original fills: white p>=0.10; grey90 0.05<=p<0.10; grey80 p<0.05. The p=0.01 contour line remains, with grey80 on both sides."))
 safe_write_xlsx(list(Funnel_Summary=summary,Funnel_Source_Rows=data,Methods=notes),file.path(d,"F3_funnel_results.xlsx"))
 safe_write_csv(summary,file.path(d,"F3_funnel_summary.csv"),row.names=FALSE,na="",fileEncoding="UTF-8")
 safe_write_lines(unlist(diagnostics),file.path(d,"F3_funnel_diagnostics.txt"))
 invisible(summary)
}

# ---- Embedded f2_module.R ----
# Statistical functions retained from the supplied F2andF3.R.
# Scope: manuscript-selected invasive-vs-native panels only.
run_f2 <- function(input_file, sheet_name, manifest) {
use_for_regression_keep <- "main"
pairing_level_keep <- c("exact", "baseline")
data_level_keep <- c("species_pair", "treatment_level", "population_pair")
tissue_keep <- "leaf"
force_lignin_to_physical <- TRUE
diff_standardization_mode <- "global_signed_norm"
standardize_lnrr <- FALSE
lnrr_standardization_mode <- "global_signed_norm"
cap_wls_weights <- FALSE
weight_cap_quantile <- 0.99
iqr_k <- 1.5
eps <- 1e-12
base_font_size <- 18
clip_panel <- TRUE
trim_unstable_ci_edges <- TRUE
min_ci_points_to_draw <- 12
point_size_range <- c(2.2, 8.0)
ci_width_cap_multiplier <- 4.0

clean_key <- function(x) {
  x <- as.character(x)
  x <- stringr::str_trim(x)
  x <- stringr::str_replace_all(x, "\\s+", "_")
  x <- stringr::str_replace_all(x, "[^A-Za-z0-9_\\-\\.]+", "-")
  x <- stringr::str_replace_all(x, "-+", "-")
  x
}

std_text <- function(x) {
  x <- as.character(x)
  x <- stringr::str_trim(x)
  x <- stringr::str_replace_all(x, "\\s+", "_")
  tolower(x)
}

std_trait <- function(x) {
  x0 <- stringr::str_trim(as.character(x))
  x1 <- tolower(x0)
  x1 <- stringr::str_replace_all(x1, "[_\\-]+", " ")
  x1 <- stringr::str_squish(x1)
  dplyr::case_when(
    x1 %in% c("c", "carbon", "leaf c", "plant c", "total c") ~ "C",
    x1 %in% c("n", "nitrogen", "leaf n", "plant n", "total n") ~ "N",
    x1 %in% c("p", "phosphorus", "phosphorous", "leaf p", "plant p", "total p") ~ "P",
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

fmt_num <- function(x, d = 3) {
  if (length(x) == 0 || is.na(x) || !is.finite(x)) return("NA")
  formatC(as.numeric(x), digits = d, format = "f")
}

var_delta <- function(SDc, Nc, SDi, Ni) {
  (SDi^2) / (Ni + eps) + (SDc^2) / (Nc + eps)
}

var_lnrr <- function(Mc, SDc, Nc, Mi, SDi, Ni) {
  (SDi^2) / ((Ni + eps) * (Mi^2 + eps)) +
    (SDc^2) / ((Nc + eps) * (Mc^2 + eps))
}

weighted_R2 <- function(y, yhat, w, clamp_zero = TRUE) {
  ok <- is.finite(y) & is.finite(yhat) & is.finite(w) & w > 0
  y <- y[ok]; yhat <- yhat[ok]; w <- w[ok]
  if (length(y) < 3) return(NA_real_)
  ybar <- sum(w * y) / sum(w)
  sst <- sum(w * (y - ybar)^2)
  sse <- sum(w * (y - yhat)^2)
  if (!is.finite(sst) || sst <= 0) return(NA_real_)
  r2 <- 1 - sse / sst
  if (!is.finite(r2)) return(NA_real_)
  if (isTRUE(clamp_zero)) r2 <- max(0, r2)
  r2
}

signed_norm_vec <- function(x, var = NULL, eps2 = 1e-8) {
  x <- as.numeric(x)
  if (is.null(var)) var <- rep(NA_real_, length(x))
  var <- as.numeric(var)
  z <- rep(NA_real_, length(x))
  dz <- rep(NA_real_, length(x))

  global_sd <- stats::sd(x[is.finite(x)], na.rm = TRUE)
  global_mu <- mean(x[is.finite(x)], na.rm = TRUE)
  if (!is.finite(global_sd) || global_sd <= eps2) global_sd <- 1
  if (!is.finite(global_mu)) global_mu <- 0

  pos <- x[is.finite(x) & x > 0]
  neg <- x[is.finite(x) & x < 0]
  mu_pos <- if (length(pos) >= 2) mean(pos) else global_mu
  sd_pos <- if (length(pos) >= 2) stats::sd(pos) else global_sd
  mu_neg <- if (length(neg) >= 2) mean(neg) else global_mu
  sd_neg <- if (length(neg) >= 2) stats::sd(neg) else global_sd
  if (!is.finite(sd_pos) || sd_pos <= eps2) sd_pos <- global_sd
  if (!is.finite(sd_neg) || sd_neg <= eps2) sd_neg <- global_sd

  for (i in seq_along(x)) {
    if (!is.finite(x[i])) next
    if (x[i] > 0) {
      z[i] <- (x[i] - mu_pos) / sd_pos
      dz[i] <- 1 / sd_pos
    } else if (x[i] < 0) {
      z[i] <- (x[i] - mu_neg) / sd_neg
      dz[i] <- 1 / sd_neg
    } else {
      z[i] <- 0
      dz[i] <- 1 / global_sd
    }
  }
  list(z = z, dz = dz, var_z = (dz^2) * var)
}

standardize_diff <- function(df) {
  tx <- signed_norm_vec(df$X_value_raw, df$Var_X_raw)
  df$zX <- tx$z
  df$Var_zX <- tx$var_z

  ty_global <- signed_norm_vec(df$Y_value_raw, df$Var_Y_raw)
  df$zY <- ty_global$z
  df$vi <- ty_global$var_z
  df$Y_standardization <- "global_signed_norm"

  if (diff_standardization_mode == "within_trait_signed_norm") {
    spl <- split(seq_len(nrow(df)), df$YTrait)
    for (ii in spl) {
      if (length(ii) >= 3) {
        ty <- signed_norm_vec(df$Y_value_raw[ii], df$Var_Y_raw[ii])
        df$zY[ii] <- ty$z
        df$vi[ii] <- ty$var_z
        df$Y_standardization[ii] <- "within_trait_signed_norm"
      }
    }
  }
  df
}

standardize_lnrr_if_needed <- function(df) {
  if (!isTRUE(standardize_lnrr)) {
    df$zX <- df$X_value_raw
    df$zY <- df$Y_value_raw
    df$vi <- df$Var_Y_raw
    df$Var_zX <- df$Var_X_raw
    df$Y_standardization <- "raw_lnRR_no_standardization"
    return(df)
  }

  tx <- signed_norm_vec(df$X_value_raw, df$Var_X_raw)
  df$zX <- tx$z
  df$Var_zX <- tx$var_z

  ty_global <- signed_norm_vec(df$Y_value_raw, df$Var_Y_raw)
  df$zY <- ty_global$z
  df$vi <- ty_global$var_z
  df$Y_standardization <- "lnRR_global_signed_norm"

  if (lnrr_standardization_mode == "within_trait_signed_norm") {
    spl <- split(seq_len(nrow(df)), df$YTrait)
    for (ii in spl) {
      if (length(ii) >= 3) {
        ty <- signed_norm_vec(df$Y_value_raw[ii], df$Var_Y_raw[ii])
        df$zY[ii] <- ty$z
        df$vi[ii] <- ty$var_z
        df$Y_standardization[ii] <- "lnRR_within_trait_signed_norm"
      }
    }
  }
  df
}

make_xgrid <- function(x, n = 240) {
  xr <- range(x, na.rm = TRUE)
  if (!all(is.finite(xr)) || diff(xr) == 0) return(seq(-1, 1, length.out = n))
  seq(xr[1], xr[2], length.out = n)
}

iqr_filter_2d <- function(df, xcol = "zX", ycol = "zY", k = 1.5) {
  if (nrow(df) < 4) return(df)
  qx <- stats::quantile(df[[xcol]], probs = c(0.25, 0.75), na.rm = TRUE)
  qy <- stats::quantile(df[[ycol]], probs = c(0.25, 0.75), na.rm = TRUE)
  iqx <- qx[2] - qx[1]; iqy <- qy[2] - qy[1]
  lx <- qx[1] - k * iqx; ux <- qx[2] + k * iqx
  ly <- qy[1] - k * iqy; uy <- qy[2] + k * iqy
  df %>% dplyr::filter(.data[[xcol]] >= lx, .data[[xcol]] <= ux,
                        .data[[ycol]] >= ly, .data[[ycol]] <= uy)
}

safe_AIC <- function(obj) {
  out <- tryCatch(stats::AIC(obj), error = function(e) NA_real_)
  if (length(out) == 0 || !is.finite(out[1])) return(NA_real_)
  as.numeric(out[1])
}

make_analysis_base <- function(comparison_type, metric) {
  valid_col <- if (metric == "LNRR") "valid_lnrr" else "valid_diff"
  dat <- raw_clean %>%
    dplyr::filter(
      .data[[valid_col]],
      ComparisonType == std_text(comparison_type),
      UseForRegression %in% std_text(use_for_regression_keep),
      PairingLevel %in% std_text(pairing_level_keep),
      DataLevel %in% std_text(data_level_keep)
    )
  if (length(tissue_keep) > 0) dat <- dat %>% dplyr::filter(Tissue %in% std_text(tissue_keep))
  dat
}

build_fit_data <- function(analysis_base, plan_row, comparison_type, metric) {
  x_trait <- plan_row$x_trait
  y_category <- plan_row$y_category
  analysis_id <- plan_row$analysis_id

  value_col <- if (metric == "LNRR") "lnRR_value" else "d_value"
  var_col   <- if (metric == "LNRR") "Var_lnRR" else "Var_d"

  xdat <- analysis_base %>%
    dplyr::filter(Trait == x_trait) %>%
    dplyr::mutate(wX = 1 / pmax(.data[[var_col]], eps)) %>%
    dplyr::group_by(SystemID) %>%
    dplyr::summarise(
      StudyID = dplyr::first(StudyID),
      ComparisonType = dplyr::first(ComparisonType),
      InvasiveSpecies = dplyr::first(InvasiveSpecies),
      ComparatorSpecies = dplyr::first(ComparatorSpecies),
      Treatment = dplyr::first(Treatment),
      Tissue = dplyr::first(Tissue),
      ExperimentType = dplyr::first(ExperimentType),
      PairDesign = dplyr::first(PairDesign),
      DataLevel = dplyr::first(DataLevel),
      PairingLevel = dplyr::first(PairingLevel),
      XTrait = dplyr::first(Trait),
      X_value_raw = stats::weighted.mean(.data[[value_col]], w = wX, na.rm = TRUE),
      Var_X_raw = 1 / sum(wX, na.rm = TRUE),
      X_n_rows = dplyr::n(),
      X_RowIDs = paste(RowID, collapse = ";"),
      .groups = "drop"
    )

  ydat <- analysis_base %>%
    dplyr::filter(TraitCategory == y_category, !(Trait %in% c("C", "N", "P"))) %>%
    dplyr::transmute(
      SystemID, StudyID, Y_RowID = RowID,
      YTrait = Trait,
      YTraitCategory = TraitCategory,
      YTraitForm = TraitForm,
      YUnit = Unit,
      YFigureTable = FigureTable,
      Y_value_raw = .data[[value_col]],
      Var_Y_raw = pmax(.data[[var_col]], eps),
      Mi_y = Mi, Mc_y = Mc, SDi_y = SDi, SDc_y = SDc, Ni_y = Ni, Nc_y = Nc,
      Notes
    )

  fit <- ydat %>%
    dplyr::inner_join(xdat, by = c("SystemID", "StudyID")) %>%
    dplyr::mutate(
      analysis_id = analysis_id,
      comparison_type = comparison_type,
      metric = metric,
      x_trait = x_trait,
      y_category = y_category
    )

  if (nrow(fit) == 0) return(fit)

  if (metric == "LNRR") {
    fit <- standardize_lnrr_if_needed(fit)
  } else {
    fit <- standardize_diff(fit)
  }

  fit <- fit %>%
    dplyr::filter(is.finite(zX), is.finite(zY), is.finite(vi), vi > 0) %>%
    dplyr::mutate(
      weight_raw = 1 / pmax(vi, eps),
      w_for_model = weight_raw,
      RegressionID = paste(SystemID, metric, x_trait, YTrait, Y_RowID, sep = "__")
    )

  if (isTRUE(cap_wls_weights) && nrow(fit) >= 4) {
    cap <- stats::quantile(fit$w_for_model, probs = weight_cap_quantile, na.rm = TRUE)
    if (is.finite(cap) && cap > 0) fit$w_for_model <- pmin(fit$w_for_model, cap)
  }
  fit
}

fit_formula <- function(form) {
  if (form == "quadratic") stats::as.formula("zY ~ zX + I(zX^2)") else stats::as.formula("zY ~ zX")
}

predict_band_lm <- function(fit, df, form, robust_vcov = NULL) {
  xg <- make_xgrid(df$zX)
  newdat <- data.frame(zX = xg)
  if (is.null(robust_vcov)) {
    pr <- stats::predict(fit, newdata = newdat, se.fit = TRUE)
    tval <- stats::qt(0.975, df = max(1, fit$df.residual))
    return(data.frame(x = xg, fit = as.numeric(pr$fit),
                      lwr = as.numeric(pr$fit - tval * pr$se.fit),
                      upr = as.numeric(pr$fit + tval * pr$se.fit)))
  }
  mm <- stats::model.matrix(fit_formula(form), data = newdat)
  b <- stats::coef(fit)
  common <- intersect(colnames(mm), names(b))
  mm <- mm[, common, drop = FALSE]
  b <- b[common]
  V <- robust_vcov[common, common, drop = FALSE]
  fitv <- as.numeric(mm %*% b)
  se <- sqrt(pmax(0, rowSums((mm %*% V) * mm)))
  df_ci <- max(5, dplyr::n_distinct(df$StudyID) - 1)
  tval <- stats::qt(0.975, df = df_ci)
  data.frame(x = xg, fit = fitv, lwr = fitv - tval * se, upr = fitv + tval * se)
}

run_WLS <- function(df, form = "linear") {
  f <- fit_formula(form)
  fit <- stats::lm(f, data = df, weights = w_for_model)
  s <- summary(fit)
  b <- stats::coef(fit)
  b0 <- unname(b[1])
  b1 <- if ("zX" %in% names(b)) unname(b["zX"]) else NA_real_
  b2 <- if ("I(zX^2)" %in% names(b)) unname(b["I(zX^2)"]) else 0
  p <- tryCatch(as.numeric(s$coefficients["zX", 4]), error = function(e) NA_real_)
  p2 <- tryCatch(as.numeric(s$coefficients["I(zX^2)", 4]), error = function(e) NA_real_)
  r2 <- weighted_R2(df$zY, stats::fitted(fit), df$w_for_model, clamp_zero = TRUE)
  band <- predict_band_lm(fit, df, form)
  list(method = paste0("WLS_", form), model_family = "WLS", model_form = form, fit = fit,
       b0 = b0, b1 = b1, b2 = b2, p = p, p_quadratic = p2, R2 = r2,
       AIC = safe_AIC(fit), band = band)
}

build_plot <- function(df, res, comparison_type, plan_row, dataset_tag, metric) {
  labs <- get_axis_labels(comparison_type, plan_row$x_trait, plan_row$y_category, metric)
  if (nrow(df) < 3 || !is.finite(res$b1) || nrow(res$band) == 0) return(NULL)

  r2_label <- paste0("bold(R)^2 == bold('", fmt_num(res$R2, 3), "')")
  line_type <- "dashed"  # user-requested final style: blue dashed fit line for all selected plots

  df <- df %>% dplyr::mutate(point_weight = w_for_model)

  ## Panel limits are based on observed points. This keeps the visual scale comparable to V4
  ## and avoids compressing points because of very wide edge confidence intervals.
  xr <- range(df$zX, na.rm = TRUE)
  yr <- range(df$zY, na.rm = TRUE)
  if (!all(is.finite(xr)) || diff(xr) == 0) xr <- c(-1, 1)
  if (!all(is.finite(yr)) || diff(yr) == 0) yr <- c(-1, 1)
  xpad <- 0.08 * diff(xr); if (!is.finite(xpad) || xpad == 0) xpad <- 0.2
  ypad <- 0.10 * diff(yr); if (!is.finite(ypad) || ypad == 0) ypad <- 0.2
  xlim <- c(xr[1] - xpad, xr[2] + xpad)
  ylim <- c(yr[1] - ypad, yr[2] + ypad)

  ## Prepare fitted curve and confidence ribbon.
  ## Important: the model and R2 are unchanged. These operations only affect plotting.
  band0 <- res$band %>%
    dplyr::filter(is.finite(x), is.finite(fit), is.finite(lwr), is.finite(upr)) %>%
    dplyr::arrange(x)
  if (nrow(band0) == 0) return(NULL)

  ## Avoid very wide CI polygons at extreme x values. This is common for quadratic WLS fits
  ## when only one or two observations occur near an edge. Such ribbons can visually dominate
  ## the panel while adding little useful information. We trim only the ribbon in those unstable
  ## regions; the fitted curve and all observed points remain plotted.
  band_ribbon <- band0
  if (isTRUE(trim_unstable_ci_edges) && nrow(band_ribbon) >= min_ci_points_to_draw) {
    y_iqr <- stats::IQR(df$zY, na.rm = TRUE)
    if (!is.finite(y_iqr) || y_iqr <= 0) y_iqr <- diff(yr)
    if (!is.finite(y_iqr) || y_iqr <= 0) y_iqr <- 1
    max_ci_width <- ci_width_cap_multiplier * y_iqr
    band_ribbon <- band_ribbon %>%
      dplyr::mutate(ci_width = upr - lwr) %>%
      dplyr::filter(is.finite(ci_width), ci_width <= max_ci_width) %>%
      dplyr::select(-ci_width)
    if (nrow(band_ribbon) < min_ci_points_to_draw) {
      ## If trimming removes almost everything, keep no ribbon rather than drawing an ugly polygon.
      band_ribbon <- band0[0, ]
    }
  }

  ## Clip ribbon limits to the observed-data panel so grey bands do not overflow the axes.
  if (nrow(band_ribbon) > 0) {
    band_ribbon <- band_ribbon %>%
      dplyr::mutate(lwr = pmax(lwr, ylim[1]), upr = pmin(upr, ylim[2]))
  }

  ## The fitted line is also clipped by the panel. Keeping clip = "on" prevents the curve
  ## from extending into the margins, which happened in the LNRR C -> chemical sensitivity panel.
  band_line <- band0 %>%
    dplyr::filter(x >= xlim[1], x <= xlim[2])
  if (nrow(band_line) == 0) return(NULL)

  ann_x <- xlim[1] + 0.035 * diff(xlim)
  ann_y <- ylim[2] - 0.005 * diff(ylim)

  p <- ggplot2::ggplot(df, ggplot2::aes(x = zX, y = zY)) +
    ggplot2::geom_hline(yintercept = 0, linetype = "dashed", linewidth = 0.55, color = "grey78") +
    ggplot2::geom_vline(xintercept = 0, linetype = "dashed", linewidth = 0.55, color = "grey78")

  if (nrow(band_ribbon) > 0) {
    p <- p + ggplot2::geom_ribbon(
      data = band_ribbon,
      ggplot2::aes(x = x, ymin = lwr, ymax = upr),
      inherit.aes = FALSE,
      fill = "grey72",
      alpha = 0.30
    )
  }

  p +
    ggplot2::geom_line(
      data = band_line,
      ggplot2::aes(x = x, y = fit),
      inherit.aes = FALSE,
      linewidth = 1.15,
      linetype = line_type,
      color = "#1F5FBF"
    ) +
    ggplot2::geom_point(
      ggplot2::aes(size = point_weight),
      shape = 21,
      fill = "grey35",
      color = "black",
      alpha = 0.72,
      stroke = 0.25
    ) +
    ggplot2::scale_size_continuous(range = point_size_range, guide = "none") +
    ggplot2::annotate("text", x = ann_x, y = ann_y, label = r2_label, parse = TRUE,
                      hjust = 0, vjust = 1, size = 5.5, fontface = "bold") +
    ggplot2::coord_cartesian(xlim = xlim, ylim = ylim, clip = ifelse(isTRUE(clip_panel), "on", "off")) +
    ggplot2::labs(x = labs$x, y = labs$y) +
    ggplot2::theme_classic(base_size = base_font_size) +
    ggplot2::theme(
      plot.title = ggplot2::element_blank(),
      plot.subtitle = ggplot2::element_blank(),
      axis.title = ggplot2::element_text(size = base_font_size + 1, face = "bold"),
      axis.text = ggplot2::element_text(size = base_font_size - 1, color = "black"),
      axis.line = ggplot2::element_line(linewidth = 0.9, color = "black"),
      axis.ticks = ggplot2::element_line(linewidth = 0.9, color = "black"),
      plot.margin = ggplot2::margin(8, 12, 8, 12)
    )
}

# 3) Read and prepare raw data
# =========================
infile <- input_file
raw <- read_input_table(infile, sheet_name)

name_map <- c(
  "StudyID" = "StudyID", "Study_Id" = "StudyID", "Number" = "StudyID",
  "SystemID" = "SystemID", "System_Id" = "SystemID",
  "ComparisonType" = "ComparisonType", "ComparsionType" = "ComparisonType", "Comparison" = "ComparisonType",
  "InvasiveSpecies" = "InvasiveSpecies", "InvasiveSpec" = "InvasiveSpecies",
  "ComparatorSpecies" = "ComparatorSpecies", "ComparatorSpec" = "ComparatorSpecies",
  "Trait" = "Trait", "TraitCategory" = "TraitCategory", "TraitForm" = "TraitForm",
  "Treatment" = "Treatment", "Tissue" = "Tissue", "ExperimentType" = "ExperimentType",
  "PairDesign" = "PairDesign", "DataLevel" = "DataLevel", "PairingLevel" = "PairingLevel",
  "UseForRegression" = "UseForRegression",
  "Mc" = "Mc", "SDc" = "SDc", "Nc" = "Nc",
  "Mi" = "Mi", "SDi" = "SDi", "Sdi" = "SDi", "Ni" = "Ni",
  "Unit" = "Unit", "FigureTable" = "FigureTable", "FigTable" = "FigureTable",
  "Notes" = "Notes", "Note" = "Notes"
)
for (old in names(name_map)) {
  if (old %in% names(raw) && !(name_map[[old]] %in% names(raw))) {
    names(raw)[names(raw) == old] <- name_map[[old]]
  }
}

required <- c("StudyID", "SystemID", "ComparisonType", "InvasiveSpecies", "ComparatorSpecies",
              "Trait", "TraitCategory", "Treatment", "Tissue", "ExperimentType", "PairDesign",
              "DataLevel", "PairingLevel", "UseForRegression", "Mc", "SDc", "Nc", "Mi", "SDi", "Ni")
missing_cols <- setdiff(required, names(raw))
if (length(missing_cols) > 0) {
  stop("Missing required columns: ", paste(missing_cols, collapse = ", "))
}
if (!("TraitForm" %in% names(raw))) raw$TraitForm <- NA_character_
if (!("Unit" %in% names(raw))) raw$Unit <- NA_character_
if (!("FigureTable" %in% names(raw))) raw$FigureTable <- NA_character_
if (!("Notes" %in% names(raw))) raw$Notes <- NA_character_

raw_clean <- raw %>%
  dplyr::mutate(
    RowID = dplyr::row_number(),
    StudyID = clean_key(.data$StudyID),
    SystemID = clean_key(.data$SystemID),
    ComparisonType = std_text(.data$ComparisonType),
    InvasiveSpecies = as.character(.data$InvasiveSpecies),
    ComparatorSpecies = as.character(.data$ComparatorSpecies),
    Trait = std_trait(.data$Trait),
    TraitCategory = std_text(.data$TraitCategory),
    TraitForm = std_text(.data$TraitForm),
    Treatment = std_text(.data$Treatment),
    Tissue = std_text(.data$Tissue),
    ExperimentType = std_text(.data$ExperimentType),
    PairDesign = std_text(.data$PairDesign),
    DataLevel = std_text(.data$DataLevel),
    PairingLevel = std_text(.data$PairingLevel),
    UseForRegression = std_text(.data$UseForRegression),
    Mc = as.numeric(.data$Mc), SDc = as.numeric(.data$SDc), Nc = as.numeric(.data$Nc),
    Mi = as.numeric(.data$Mi), SDi = as.numeric(.data$SDi), Ni = as.numeric(.data$Ni),
    Unit = as.character(.data$Unit),
    FigureTable = as.character(.data$FigureTable),
    Notes = as.character(.data$Notes)
  ) %>%
  dplyr::mutate(
    TraitCategory = dplyr::if_else(force_lignin_to_physical & Trait == "Lignin", "physical_defense", TraitCategory),
    d_value = Mi - Mc,
    Var_d = var_delta(SDc, Nc, SDi, Ni),
    valid_diff = is.finite(Mi) & is.finite(Mc) & is.finite(SDi) & is.finite(SDc) &
      is.finite(Ni) & is.finite(Nc) & Ni > 0 & Nc > 0 & SDi >= 0 & SDc >= 0 &
      is.finite(d_value) & is.finite(Var_d) & Var_d >= 0,
    lnRR_value = dplyr::if_else(Mi > 0 & Mc > 0, log(Mi / Mc), NA_real_),
    Var_lnRR = var_lnrr(Mc, SDc, Nc, Mi, SDi, Ni),
    valid_lnrr = is.finite(lnRR_value) & is.finite(Var_lnRR) & Var_lnRR >= 0 &
      Mi > 0 & Mc > 0 & Ni > 0 & Nc > 0 & SDi >= 0 & SDc >= 0
  )


get_axis_labels <- function(comparison_type, x_trait, y_category, metric) {
 list(x=nutrient_axis_label(x_trait), y=paste0("lnRR-", ifelse(y_category == "physical_defense", "Physical defense", "Chemical defense")))
}

plan <- manifest[manifest$Figure == "F2", ]
base <- make_analysis_base("invasive_vs_native", "LNRR")
assert_system_study_keys(base)
output <- vector("list", nrow(plan))
for (i in seq_len(nrow(plan))) {
 spec <- plan[i, ]
 message("Fitting ", spec$Panel, " (", spec$Dataset, ")")
 row <- tibble::tibble(analysis_id=spec$AnalysisID, x_trait=spec$XTrait, y_category=spec$YCategory)
 dat_raw <- build_fit_data(base, row, "invasive_vs_native", "LNRR")
 dat <- if (spec$Dataset == "IQR_removed") iqr_filter_2d(dat_raw, k=iqr_k) else dat_raw
 validate_fit_input(dat, spec)
 dat$dataset <- spec$Dataset
 res <- run_WLS(dat, "quadratic")
 plot <- build_plot(dat, res, "invasive_vs_native", row, spec$Dataset, "LNRR")
 stopifnot(!is.null(plot), all(is.finite(stats::coef(res$fit))))
 summary <- model_summary_row(spec, dat, res$fit, res$R2)
 output[[i]] <- panel_result(spec, dat_raw, dat, res$fit, res$band, plot, summary)
}
output

}


# ---- Embedded f3_module.R ----
# Statistical functions retained from the supplied F2andF3.R.
# Scope: manuscript-selected invasive-vs-native panels only.
run_f3 <- function(input_file, sheet_name, manifest) {
use_for_regression_keep <- "main"
pairing_level_keep <- c("exact", "baseline")
data_level_keep <- c("species_pair", "treatment_level", "population_pair")
tissue_keep <- "leaf"
force_lignin_to_physical <- TRUE
diff_standardization_mode <- "global_signed_norm"
standardize_lnrr <- FALSE
lnrr_standardization_mode <- "global_signed_norm"
cap_wls_weights <- FALSE
weight_cap_quantile <- 0.99
iqr_k <- 1.5
eps <- 1e-12
base_font_size <- 18
clip_panel <- TRUE
trim_unstable_ci_edges <- TRUE
min_ci_points_to_draw <- 12
min_n_linear <- 3
min_n_quadratic <- 6

## Model-selection thresholds for deciding whether a fitted relationship is usable.
## IMPORTANT:
##   The two comparison types have different replication realities.
##   invasive_vs_native usually has more studies, so StudyID >= 3 is retained.
##   introduced_vs_native_range is rare, so StudyID >= 2 is allowed, but results are
##   handled by predefined comparison-specific replication criteria.
## These historical rules are retained to audit the manuscript model forms.
## The eight relationships and RAW/IQR variants are fixed in the manifest;
## they are not described as an independently prespecified discovery analysis.
selection_alpha <- 0.05
selection_min_r2 <- 0.10
selection_min_n <- 10
selection_min_studies_invasive_vs_native <- 3
selection_min_studies_introduced_vs_native_range <- 2
selection_min_systems <- 5
selection_min_delta_aic_for_quadratic <- 2
selection_min_delta_r2_for_quadratic <- 0.05
prefer_quadratic_for_structural_total <- TRUE

## Optional supplementary tier.
## If TRUE, models that are close but not strict can be exported with
## SelectionStatus = "selected_supplementary". I recommend keeping this TRUE only for
## supplementary model inspection; such results should be reported with the criteria used.
allow_supplementary_output <- TRUE
selection_alpha_supplementary <- 0.10
selection_min_r2_supplementary <- 0.08

base_font_size <- 18
point_size_range <- c(3.0, 8.2)
eps <- 1e-12

## Plot safety controls. These do not change model fitting or reported statistics.
clip_panel <- TRUE
trim_unstable_ci_edges <- TRUE
ci_width_cap_multiplier <- 2.2
min_ci_points_to_draw <- 12
ci_trim_quantile <- 0.90
ribbon_alpha <- 0.24

## Structural material trait list.
## These are used for the C -> total structural material defense analysis.
structural_material_traits <- c("Cellulose", "Hemicellulose", "Lignin", "Wax", "Resin", "Latex")

structural_material_traits <- c("Cellulose", "Hemicellulose", "Lignin", "Wax", "Resin", "Latex")
clean_key <- function(x) {
  x <- as.character(x)
  x <- stringr::str_trim(x)
  x <- stringr::str_replace_all(x, "\\s+", "_")
  x <- stringr::str_replace_all(x, "[^A-Za-z0-9_\\-\\.]+", "-")
  x <- stringr::str_replace_all(x, "-+", "-")
  x
}

std_text <- function(x) {
  x <- as.character(x)
  x <- stringr::str_trim(x)
  x <- stringr::str_replace_all(x, "\\s+", "_")
  tolower(x)
}

std_trait <- function(x) {
  x0 <- stringr::str_trim(as.character(x))
  x1 <- tolower(x0)
  x1 <- stringr::str_replace_all(x1, "[_\\-]+", " ")
  x1 <- stringr::str_squish(x1)
  dplyr::case_when(
    x1 %in% c("c", "carbon", "leaf c", "plant c", "total c") ~ "C",
    x1 %in% c("n", "nitrogen", "leaf n", "plant n", "total n") ~ "N",
    x1 %in% c("p", "phosphorus", "phosphorous", "leaf p", "plant p", "total p") ~ "P",
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

fmt_num <- function(x, d = 3) {
  if (length(x) == 0 || is.na(x) || !is.finite(x)) return("NA")
  formatC(as.numeric(x), digits = d, format = "f")
}

var_delta <- function(SDc, Nc, SDi, Ni) {
  (SDi^2) / (Ni + eps) + (SDc^2) / (Nc + eps)
}

var_lnrr <- function(Mc, SDc, Nc, Mi, SDi, Ni) {
  (SDi^2) / ((Ni + eps) * (Mi^2 + eps)) +
    (SDc^2) / ((Nc + eps) * (Mc^2 + eps))
}

weighted_R2 <- function(y, yhat, w, clamp_zero = TRUE) {
  ok <- is.finite(y) & is.finite(yhat) & is.finite(w) & w > 0
  y <- y[ok]; yhat <- yhat[ok]; w <- w[ok]
  if (length(y) < 3) return(NA_real_)
  ybar <- sum(w * y) / sum(w)
  sst <- sum(w * (y - ybar)^2)
  sse <- sum(w * (y - yhat)^2)
  if (!is.finite(sst) || sst <= 0) return(NA_real_)
  r2 <- 1 - sse / sst
  if (!is.finite(r2)) return(NA_real_)
  if (isTRUE(clamp_zero)) r2 <- max(0, r2)
  r2
}

signed_norm_vec <- function(x, var = NULL, eps2 = 1e-8) {
  x <- as.numeric(x)
  if (is.null(var)) var <- rep(NA_real_, length(x))
  var <- as.numeric(var)
  z <- rep(NA_real_, length(x))
  dz <- rep(NA_real_, length(x))

  global_sd <- stats::sd(x[is.finite(x)], na.rm = TRUE)
  global_mu <- mean(x[is.finite(x)], na.rm = TRUE)
  if (!is.finite(global_sd) || global_sd <= eps2) global_sd <- 1
  if (!is.finite(global_mu)) global_mu <- 0

  pos <- x[is.finite(x) & x > 0]
  neg <- x[is.finite(x) & x < 0]
  mu_pos <- if (length(pos) >= 2) mean(pos) else global_mu
  sd_pos <- if (length(pos) >= 2) stats::sd(pos) else global_sd
  mu_neg <- if (length(neg) >= 2) mean(neg) else global_mu
  sd_neg <- if (length(neg) >= 2) stats::sd(neg) else global_sd
  if (!is.finite(sd_pos) || sd_pos <= eps2) sd_pos <- global_sd
  if (!is.finite(sd_neg) || sd_neg <= eps2) sd_neg <- global_sd

  for (i in seq_along(x)) {
    if (!is.finite(x[i])) next
    if (x[i] > 0) {
      z[i] <- (x[i] - mu_pos) / sd_pos
      dz[i] <- 1 / sd_pos
    } else if (x[i] < 0) {
      z[i] <- (x[i] - mu_neg) / sd_neg
      dz[i] <- 1 / sd_neg
    } else {
      z[i] <- 0
      dz[i] <- 1 / global_sd
    }
  }
  list(z = z, dz = dz, var_z = (dz^2) * var)
}

standardize_diff <- function(df) {
  tx <- signed_norm_vec(df$X_value_raw, df$Var_X_raw)
  df$zX <- tx$z
  df$Var_zX <- tx$var_z

  ty_global <- signed_norm_vec(df$Y_value_raw, df$Var_Y_raw)
  df$zY <- ty_global$z
  df$vi <- ty_global$var_z
  df$Y_standardization <- "global_signed_norm"

  if (diff_standardization_mode == "within_trait_signed_norm") {
    spl <- split(seq_len(nrow(df)), df$YTrait)
    for (ii in spl) {
      if (length(ii) >= 3) {
        ty <- signed_norm_vec(df$Y_value_raw[ii], df$Var_Y_raw[ii])
        df$zY[ii] <- ty$z
        df$vi[ii] <- ty$var_z
        df$Y_standardization[ii] <- "within_trait_signed_norm"
      }
    }
  }
  df
}

standardize_lnrr_if_needed <- function(df) {
  if (!isTRUE(standardize_lnrr)) {
    df$zX <- df$X_value_raw
    df$zY <- df$Y_value_raw
    df$vi <- df$Var_Y_raw
    df$Var_zX <- df$Var_X_raw
    df$Y_standardization <- "raw_lnRR_no_standardization"
    return(df)
  }

  tx <- signed_norm_vec(df$X_value_raw, df$Var_X_raw)
  df$zX <- tx$z
  df$Var_zX <- tx$var_z

  ty_global <- signed_norm_vec(df$Y_value_raw, df$Var_Y_raw)
  df$zY <- ty_global$z
  df$vi <- ty_global$var_z
  df$Y_standardization <- "lnRR_global_signed_norm"

  if (lnrr_standardization_mode == "within_trait_signed_norm") {
    spl <- split(seq_len(nrow(df)), df$YTrait)
    for (ii in spl) {
      if (length(ii) >= 3) {
        ty <- signed_norm_vec(df$Y_value_raw[ii], df$Var_Y_raw[ii])
        df$zY[ii] <- ty$z
        df$vi[ii] <- ty$var_z
        df$Y_standardization[ii] <- "lnRR_within_trait_signed_norm"
      }
    }
  }
  df
}

make_xgrid <- function(x, n = 240) {
  xr <- range(x, na.rm = TRUE)
  if (!all(is.finite(xr)) || diff(xr) == 0) return(seq(-1, 1, length.out = n))
  seq(xr[1], xr[2], length.out = n)
}

iqr_filter_2d <- function(df, xcol = "zX", ycol = "zY", k = 1.5) {
  if (nrow(df) < 4) return(df)
  qx <- stats::quantile(df[[xcol]], probs = c(0.25, 0.75), na.rm = TRUE)
  qy <- stats::quantile(df[[ycol]], probs = c(0.25, 0.75), na.rm = TRUE)
  iqx <- qx[2] - qx[1]
  iqy <- qy[2] - qy[1]
  lx <- qx[1] - k * iqx
  ux <- qx[2] + k * iqx
  ly <- qy[1] - k * iqy
  uy <- qy[2] + k * iqy
  df %>% dplyr::filter(.data[[xcol]] >= lx, .data[[xcol]] <= ux,
                        .data[[ycol]] >= ly, .data[[ycol]] <= uy)
}

rescale_point_sizes <- function(w, size_range = point_size_range) {
  w <- as.numeric(w)
  w[!is.finite(w)] <- NA_real_
  out <- rep(mean(size_range), length(w))
  ok <- is.finite(w)
  if (!any(ok)) return(out)
  base <- sqrt(pmax(w[ok], eps))
  if (length(unique(round(base, 10))) <= 1) {
    out[ok] <- mean(size_range)
    return(out)
  }
  out[ok] <- scales::rescale(base, to = size_range, from = range(base, na.rm = TRUE))
  out <- pmax(size_range[1], pmin(size_range[2], out))
  out
}

safe_file_label <- function(x) {
  x <- stringr::str_replace_all(x, "[^A-Za-z0-9_.-]+", "_")
  x <- stringr::str_replace_all(x, "_+", "_")
  stringr::str_replace_all(x, "(^_|_$)", "")
}

plot_file_label <- function(comparison_type, x_trait, y_label, metric) {
  prefix <- if (metric == "LNRR") {
    ifelse(comparison_type == "introduced_vs_native_range", "LNRR_I.range-N.range", "LNRR_INV-NAT")
  } else {
    ifelse(comparison_type == "introduced_vs_native_range", "Cha_I.range-N.range", "Diff_INV-NAT")
  }
  paste(prefix, paste0("for_", x_trait, "_vs_", safe_file_label(y_label)), sep = "_")
}

make_analysis_base <- function(comparison_type, metric) {
  valid_col <- if (metric == "LNRR") "valid_lnrr" else "valid_diff"
  dat <- raw_clean %>%
    dplyr::filter(
      .data[[valid_col]],
      ComparisonType == std_text(comparison_type),
      UseForRegression %in% std_text(use_for_regression_keep),
      PairingLevel %in% std_text(pairing_level_keep),
      DataLevel %in% std_text(data_level_keep)
    )
  if (length(tissue_keep) > 0) dat <- dat %>% dplyr::filter(Tissue %in% std_text(tissue_keep))
  dat
}

filter_y_by_target <- function(analysis_base, target_row) {
  target_type <- target_row$TargetType
  if (target_type == "structural_material_total") {
    analysis_base %>%
      dplyr::filter(
        TraitCategory == "physical_defense",
        TraitForm == "material",
        Trait %in% structural_material_traits,
        !(Trait %in% c("C", "N", "P"))
      )
  } else if (target_type == "specific_trait") {
    ydat <- analysis_base %>% dplyr::filter(Trait == target_row$YTrait)
    if (!is.na(target_row$YCategory) && nzchar(target_row$YCategory)) {
      ydat <- ydat %>% dplyr::filter(TraitCategory == target_row$YCategory)
    }
    ydat %>% dplyr::filter(!(Trait %in% c("C", "N", "P")))
  } else {
    stop("Unknown TargetType: ", target_type)
  }
}

build_fit_data <- function(analysis_base, target_row, comparison_type, metric) {
  x_trait <- target_row$XTrait
  value_col <- if (metric == "LNRR") "lnRR_value" else "d_value"
  var_col   <- if (metric == "LNRR") "Var_lnRR" else "Var_d"

  xdat <- analysis_base %>%
    dplyr::filter(Trait == x_trait) %>%
    dplyr::mutate(wX = 1 / pmax(.data[[var_col]], eps)) %>%
    dplyr::group_by(SystemID) %>%
    dplyr::summarise(
      StudyID = dplyr::first(StudyID),
      ComparisonType = dplyr::first(ComparisonType),
      InvasiveSpecies = dplyr::first(InvasiveSpecies),
      ComparatorSpecies = dplyr::first(ComparatorSpecies),
      Treatment = dplyr::first(Treatment),
      Tissue = dplyr::first(Tissue),
      ExperimentType = dplyr::first(ExperimentType),
      PairDesign = dplyr::first(PairDesign),
      DataLevel = dplyr::first(DataLevel),
      PairingLevel = dplyr::first(PairingLevel),
      XTrait = dplyr::first(Trait),
      X_value_raw = stats::weighted.mean(.data[[value_col]], w = wX, na.rm = TRUE),
      Var_X_raw = 1 / sum(wX, na.rm = TRUE),
      X_n_rows = dplyr::n(),
      X_RowIDs = paste(RowID, collapse = ";"),
      .groups = "drop"
    )

  yraw <- filter_y_by_target(analysis_base, target_row)
  ydat <- yraw %>%
    dplyr::transmute(
      SystemID, StudyID, Y_RowID = RowID,
      YTrait = Trait,
      YTraitCategory = TraitCategory,
      YTraitForm = TraitForm,
      YUnit = Unit,
      YFigureTable = FigureTable,
      Y_value_raw = .data[[value_col]],
      Var_Y_raw = pmax(.data[[var_col]], eps),
      Mi_y = Mi, Mc_y = Mc, SDi_y = SDi, SDc_y = SDc, Ni_y = Ni, Nc_y = Nc,
      Notes
    )

  fit <- ydat %>%
    ## Pair nutrients and defense traits primarily by SystemID.
    ## SystemID is the user-defined biological comparison unit.
    dplyr::inner_join(xdat, by = "SystemID", suffix = c("_Y", "_X")) %>%
    dplyr::mutate(
      StudyID = dplyr::coalesce(.data$StudyID_Y, .data$StudyID_X),
      AnalysisOrder = target_row$AnalysisOrder,
      AnalysisID = target_row$AnalysisID,
      TargetType = target_row$TargetType,
      comparison_type = comparison_type,
      metric = metric,
      x_trait = x_trait,
      y_label = target_row$YLabel,
      y_target_trait = target_row$YTrait
    ) %>%
    dplyr::select(-dplyr::any_of(c("StudyID_Y", "StudyID_X")))

  if (nrow(fit) == 0) return(fit)

  if (metric == "LNRR") {
    fit <- standardize_lnrr_if_needed(fit)
  } else {
    fit <- standardize_diff(fit)
  }

  fit <- fit %>%
    dplyr::filter(is.finite(zX), is.finite(zY), is.finite(vi), vi > 0) %>%
    dplyr::mutate(
      weight_raw = 1 / pmax(vi, eps),
      w_for_model = weight_raw,
      RegressionID = paste(SystemID, metric, x_trait, YTrait, Y_RowID, sep = "__")
    )

  if (isTRUE(cap_wls_weights) && nrow(fit) >= 4) {
    cap <- stats::quantile(fit$w_for_model, probs = weight_cap_quantile, na.rm = TRUE)
    if (is.finite(cap) && cap > 0) fit$w_for_model <- pmin(fit$w_for_model, cap)
  }
  fit
}

run_WLS_model <- function(df, model_form = c("linear", "quadratic")) {
  model_form <- match.arg(model_form)
  min_n <- if (model_form == "quadratic") min_n_quadratic else min_n_linear
  if (nrow(df) < min_n) stop(paste0("too few rows for WLS ", model_form))
  if (length(unique(df$zX[is.finite(df$zX)])) < ifelse(model_form == "quadratic", 3, 2)) {
    stop(paste0("too few unique x values for WLS ", model_form))
  }

  if (model_form == "quadratic") {
    fit <- stats::lm(zY ~ zX + I(zX^2), data = df, weights = w_for_model)
  } else {
    fit <- stats::lm(zY ~ zX, data = df, weights = w_for_model)
  }

  s <- summary(fit)
  b <- stats::coef(fit)
  b0 <- unname(b[1])
  b1 <- if ("zX" %in% names(b)) unname(b["zX"]) else NA_real_
  b2 <- if ("I(zX^2)" %in% names(b)) unname(b["I(zX^2)"]) else NA_real_
  p1 <- tryCatch(as.numeric(s$coefficients["zX", 4]), error = function(e) NA_real_)
  p2 <- tryCatch(as.numeric(s$coefficients["I(zX^2)", 4]), error = function(e) NA_real_)
  r2 <- weighted_R2(df$zY, stats::fitted(fit), df$w_for_model, clamp_zero = TRUE)
  list(fit = fit, model_form = model_form, b0 = b0, b1 = b1, b2 = b2,
       p_linear = p1, p_quadratic = p2, R2 = r2, AIC = stats::AIC(fit), error = "")
}

model_row_from_result <- function(res, dat, target_row, cmp, metric, dataset_tag) {
  tibble::tibble(
    AnalysisOrder = target_row$AnalysisOrder,
    Metric = metric,
    ComparisonType = cmp,
    AnalysisID = target_row$AnalysisID,
    TargetType = target_row$TargetType,
    PlotLabel = plot_file_label(cmp, target_row$XTrait, target_row$YLabel, metric),
    XTrait = target_row$XTrait,
    YLabel = target_row$YLabel,
    YTargetTrait = target_row$YTrait,
    Dataset = dataset_tag,
    Method = "WLS",
    ModelForm = res$model_form,
    n_rows = nrow(dat),
    n_studies = dplyr::n_distinct(dat$StudyID),
    n_systems = dplyr::n_distinct(dat$SystemID),
    n_response_traits = dplyr::n_distinct(dat$YTrait),
    response_traits = if (nrow(dat) > 0) paste(sort(unique(dat$YTrait)), collapse = "; ") else NA_character_,
    Intercept = res$b0,
    Slope_linear = res$b1,
    Slope_quadratic = res$b2,
    P_linear = res$p_linear,
    P_quadratic = res$p_quadratic,
    R2 = res$R2,
    AIC = res$AIC,
    Error = res$error
  )
}

fit_candidate_models <- function(dat, target_row, cmp, metric, dataset_tag) {
  out <- list()
  for (mf in c("linear", "quadratic")) {
    res <- tryCatch(
      run_WLS_model(dat, mf),
      error = function(e) list(fit = NULL, model_form = mf, b0 = NA_real_, b1 = NA_real_, b2 = NA_real_,
                               p_linear = NA_real_, p_quadratic = NA_real_, R2 = NA_real_,
                               AIC = NA_real_, error = e$message)
    )
    out[[mf]] <- list(result = res, summary = model_row_from_result(res, dat, target_row, cmp, metric, dataset_tag))
  }
  out
}

min_studies_required <- function(comparison_type) {
  cmp <- std_text(comparison_type)
  if (identical(cmp, "introduced_vs_native_range")) {
    return(selection_min_studies_introduced_vs_native_range)
  }
  selection_min_studies_invasive_vs_native
}

selection_tier_label <- function(row, comparison_type) {
  ## The minimum StudyID threshold is predefined separately for each comparison type.
  ## If a model passes that prespecified threshold and the other criteria, it is treated
  ## as a selected model without additional labels in output file names.
  if (!is.finite(row$n_studies)) return("not_selected")
  "selected"
}

passes_common_quality <- function(row, comparison_type, min_r2 = selection_min_r2) {
  if (!is.na(row$Error) && nzchar(row$Error)) return(FALSE)
  if (!is.finite(row$R2) || row$R2 < min_r2) return(FALSE)
  if (!is.finite(row$n_rows) || row$n_rows < selection_min_n) return(FALSE)
  if (!is.finite(row$n_studies) || row$n_studies < min_studies_required(comparison_type)) return(FALSE)
  if (!is.finite(row$n_systems) || row$n_systems < selection_min_systems) return(FALSE)
  TRUE
}

passes_model_term <- function(row, model_form, alpha_value) {
  if (identical(model_form, "linear")) {
    return(is.finite(row$P_linear) && row$P_linear <= alpha_value)
  }
  is.finite(row$P_quadratic) && row$P_quadratic <= alpha_value
}

make_selection_decision <- function(linear_row, quad_row, target_type, comparison_type) {
  min_studies_now <- min_studies_required(comparison_type)

  lin_ok <- passes_common_quality(linear_row, comparison_type, selection_min_r2) &&
    passes_model_term(linear_row, "linear", selection_alpha)
  quad_ok_base <- passes_common_quality(quad_row, comparison_type, selection_min_r2) &&
    passes_model_term(quad_row, "quadratic", selection_alpha)

  delta_aic <- if (is.finite(linear_row$AIC) && is.finite(quad_row$AIC)) linear_row$AIC - quad_row$AIC else NA_real_
  delta_r2  <- if (is.finite(linear_row$R2) && is.finite(quad_row$R2)) quad_row$R2 - linear_row$R2 else NA_real_

  quad_clearly_better <- quad_ok_base &&
    is.finite(delta_aic) && is.finite(delta_r2) &&
    delta_aic >= selection_min_delta_aic_for_quadratic &&
    delta_r2 >= selection_min_delta_r2_for_quadratic

  is_structural_total <- identical(target_type, "structural_material_total")

  status_if_selected <- selection_tier_label(quad_row, comparison_type)
  status_if_selected_linear <- selection_tier_label(linear_row, comparison_type)

  if (is_structural_total && isTRUE(prefer_quadratic_for_structural_total) && quad_ok_base) {
    return(tibble::tibble(
      SelectedModelForm = "quadratic",
      SelectionStatus = status_if_selected,
      SelectionReason = paste0("structural material total: quadratic passed quality criteria; ",
                               "pooled multiple structural material traits; minimum StudyID for this comparison = ",
                               min_studies_now),
      LinearPass = lin_ok,
      QuadraticPass = quad_ok_base,
      DeltaAIC_linear_minus_quadratic = delta_aic,
      DeltaR2_quadratic_minus_linear = delta_r2,
      MinStudiesRequired = min_studies_now
    ))
  }

  if (quad_clearly_better) {
    return(tibble::tibble(
      SelectedModelForm = "quadratic",
      SelectionStatus = status_if_selected,
      SelectionReason = paste0("quadratic selected: P_quadratic <= ", selection_alpha,
                               ", R2 >= ", selection_min_r2,
                               ", AIC improved by >= ", selection_min_delta_aic_for_quadratic,
                               ", R2 improved by >= ", selection_min_delta_r2_for_quadratic,
                               ", and StudyID >= ", min_studies_now),
      LinearPass = lin_ok,
      QuadraticPass = quad_ok_base,
      DeltaAIC_linear_minus_quadratic = delta_aic,
      DeltaR2_quadratic_minus_linear = delta_r2,
      MinStudiesRequired = min_studies_now
    ))
  }

  if (lin_ok) {
    return(tibble::tibble(
      SelectedModelForm = "linear",
      SelectionStatus = status_if_selected_linear,
      SelectionReason = paste0("linear selected by parsimony: linear passed quality criteria and quadratic was not clearly better; ",
                               "minimum StudyID for this comparison = ", min_studies_now),
      LinearPass = lin_ok,
      QuadraticPass = quad_ok_base,
      DeltaAIC_linear_minus_quadratic = delta_aic,
      DeltaR2_quadratic_minus_linear = delta_r2,
      MinStudiesRequired = min_studies_now
    ))
  }

  if (quad_ok_base && !lin_ok) {
    return(tibble::tibble(
      SelectedModelForm = "quadratic",
      SelectionStatus = status_if_selected,
      SelectionReason = paste0("quadratic selected: linear did not pass, but quadratic term passed quality criteria; ",
                               "minimum StudyID for this comparison = ", min_studies_now),
      LinearPass = lin_ok,
      QuadraticPass = quad_ok_base,
      DeltaAIC_linear_minus_quadratic = delta_aic,
      DeltaR2_quadratic_minus_linear = delta_r2,
      MinStudiesRequired = min_studies_now
    ))
  }

  ## Optional supplementary tier: less strict p/R2 thresholds, still with replication filters.
  ## This allows the user to inspect borderline relationships without presenting them as
  ## fully supported main results.
  if (isTRUE(allow_supplementary_output)) {
    lin_exp <- passes_common_quality(linear_row, comparison_type, selection_min_r2_supplementary) &&
      passes_model_term(linear_row, "linear", selection_alpha_supplementary)
    quad_exp <- passes_common_quality(quad_row, comparison_type, selection_min_r2_supplementary) &&
      passes_model_term(quad_row, "quadratic", selection_alpha_supplementary)

    if (quad_exp || lin_exp) {
      choose_quad <- FALSE
      if (quad_exp && lin_exp) {
        choose_quad <- is.finite(delta_aic) && is.finite(delta_r2) &&
          delta_aic >= selection_min_delta_aic_for_quadratic &&
          delta_r2 >= selection_min_delta_r2_for_quadratic
      } else if (quad_exp && !lin_exp) {
        choose_quad <- TRUE
      }

      chosen_form <- if (choose_quad) "quadratic" else "linear"
      return(tibble::tibble(
        SelectedModelForm = chosen_form,
        SelectionStatus = "selected_supplementary",
        SelectionReason = paste0("supplementary selected: secondary criteria met (P <= ",
                                 selection_alpha_supplementary, ", R2 >= ",
                                 selection_min_r2_supplementary, ", StudyID >= ",
                                 min_studies_now, "); reported as supplementary evidence"),
        LinearPass = lin_exp,
        QuadraticPass = quad_exp,
        DeltaAIC_linear_minus_quadratic = delta_aic,
        DeltaR2_quadratic_minus_linear = delta_r2,
        MinStudiesRequired = min_studies_now
      ))
    }
  }

  ## Explain the first major reason for rejection so the user can audit results.
  fail_parts <- c()
  if (linear_row$n_rows < selection_min_n || quad_row$n_rows < selection_min_n) fail_parts <- c(fail_parts, paste0("n < ", selection_min_n))
  if (linear_row$n_studies < min_studies_now || quad_row$n_studies < min_studies_now) fail_parts <- c(fail_parts, paste0("StudyID < ", min_studies_now))
  if (linear_row$n_systems < selection_min_systems || quad_row$n_systems < selection_min_systems) fail_parts <- c(fail_parts, paste0("SystemID < ", selection_min_systems))
  if ((is.finite(linear_row$R2) && linear_row$R2 < selection_min_r2) && (is.finite(quad_row$R2) && quad_row$R2 < selection_min_r2)) fail_parts <- c(fail_parts, paste0("R2 < ", selection_min_r2))
  if (!(is.finite(linear_row$P_linear) && linear_row$P_linear <= selection_alpha) && !(is.finite(quad_row$P_quadratic) && quad_row$P_quadratic <= selection_alpha)) fail_parts <- c(fail_parts, paste0("no significant linear/quadratic term at alpha = ", selection_alpha))
  if (length(fail_parts) == 0) fail_parts <- "did not pass predefined model-selection criteria"

  tibble::tibble(
    SelectedModelForm = NA_character_,
    SelectionStatus = "not_selected",
    SelectionReason = paste(unique(fail_parts), collapse = "; "),
    LinearPass = lin_ok,
    QuadraticPass = quad_ok_base,
    DeltaAIC_linear_minus_quadratic = delta_aic,
    DeltaR2_quadratic_minus_linear = delta_r2,
    MinStudiesRequired = min_studies_now
  )
}

make_prediction_band <- function(fit, df) {
  xg <- make_xgrid(df$zX, n = 240)
  newdat <- data.frame(zX = xg)
  pr <- stats::predict(fit, newdata = newdat, se.fit = TRUE)
  tval <- stats::qt(0.975, df = max(1, fit$df.residual))
  data.frame(
    x = xg,
    fit = as.numeric(pr$fit),
    lwr = as.numeric(pr$fit - tval * pr$se.fit),
    upr = as.numeric(pr$fit + tval * pr$se.fit)
  )
}

build_plot <- function(df, model_res, comparison_type, target_row, dataset_tag, metric, model_form) {
  if (nrow(df) < min_n_linear || !is.finite(model_res$R2)) return(NULL)
  band <- make_prediction_band(model_res$fit, df)
  if (nrow(band) == 0) return(NULL)

  axis_labs <- get_axis_labels(comparison_type, target_row$XTrait, target_row$YLabel, metric)
  r2_label <- paste0("bold(R)^2 == bold('", fmt_num(model_res$R2, 3), "')")
  df <- df %>% dplyr::mutate(
    point_weight = w_for_model,
    point_size = rescale_point_sizes(w_for_model, point_size_range)
  )

  xr <- range(df$zX, na.rm = TRUE)
  if (!all(is.finite(xr)) || diff(xr) == 0) xr <- xr + c(-1, 1)
  xpad <- 0.08 * diff(xr)
  if (!is.finite(xpad) || xpad <= 0) xpad <- 0.15
  xlim <- c(xr[1] - xpad, xr[2] + xpad)

  band_finite <- band %>%
    dplyr::filter(is.finite(x), is.finite(fit), is.finite(lwr), is.finite(upr)) %>%
    dplyr::filter(x >= xlim[1], x <= xlim[2])

  y_ref <- c(df$zY, band_finite$fit)
  yr <- range(y_ref[is.finite(y_ref)], na.rm = TRUE)
  if (!all(is.finite(yr)) || diff(yr) == 0) yr <- yr + c(-1, 1)
  ypad <- 0.14 * diff(yr)
  if (!is.finite(ypad) || ypad <= 0) ypad <- 0.15
  ylim <- c(yr[1] - ypad, yr[2] + ypad)
  ylim <- range(c(ylim, 0), na.rm = TRUE)
  if (!all(is.finite(ylim)) || diff(ylim) == 0) ylim <- c(-1, 1)

  band_plot <- band_finite
  draw_ribbon <- FALSE
  if (nrow(band_plot) >= min_ci_points_to_draw) {
    observed_span <- diff(range(df$zY, na.rm = TRUE))
    if (!is.finite(observed_span) || observed_span <= 0) observed_span <- diff(ylim)
    if (!is.finite(observed_span) || observed_span <= 0) observed_span <- 1
    band_plot <- band_plot %>% dplyr::mutate(ci_width = upr - lwr)
    width_cap_1 <- ci_width_cap_multiplier * observed_span
    width_cap_2 <- suppressWarnings(stats::quantile(band_plot$ci_width, probs = ci_trim_quantile, na.rm = TRUE))
    width_cap <- min(width_cap_1, width_cap_2, na.rm = TRUE)
    if (!is.finite(width_cap) || width_cap <= 0) width_cap <- width_cap_1
    band_trimmed <- band_plot %>% dplyr::filter(ci_width <= width_cap)
    if (nrow(band_trimmed) >= min_ci_points_to_draw) band_plot <- band_trimmed

    band_plot <- band_plot %>%
      dplyr::mutate(
        lwr_plot = pmax(lwr, ylim[1]),
        upr_plot = pmin(upr, ylim[2])
      ) %>%
      dplyr::filter(is.finite(lwr_plot), is.finite(upr_plot), upr_plot >= lwr_plot) %>%
      dplyr::select(-ci_width)
    draw_ribbon <- nrow(band_plot) >= min_ci_points_to_draw
  }

  line_plot <- band_finite %>% dplyr::mutate(fit_plot = fit) %>% dplyr::filter(is.finite(fit_plot))

  p <- ggplot2::ggplot(df, ggplot2::aes(x = zX, y = zY)) +
    ggplot2::geom_hline(yintercept = 0, linetype = "dashed", linewidth = 0.45, color = "grey78") +
    ggplot2::geom_vline(xintercept = 0, linetype = "dashed", linewidth = 0.45, color = "grey78")

  if (draw_ribbon) {
    p <- p + ggplot2::geom_ribbon(
      data = band_plot,
      ggplot2::aes(x = x, ymin = lwr_plot, ymax = upr_plot),
      inherit.aes = FALSE, fill = "grey70", alpha = ribbon_alpha
    )
  }

  p <- p +
    ggplot2::geom_line(
      data = line_plot,
      ggplot2::aes(x = x, y = fit_plot), inherit.aes = FALSE,
      color = "#2C6BD8", linewidth = 1.08, linetype = "dashed"
    ) +
    ggplot2::geom_point(
      ggplot2::aes(size = point_size), shape = 21,
      fill = "grey45", color = "black", stroke = 0.22, alpha = 0.78
    ) +
    ggplot2::annotate(
      "text", x = xlim[1] + 0.07 * diff(xlim), y = ylim[2] - 0.025 * diff(ylim),
      label = r2_label, parse = TRUE, hjust = 0, vjust = 1, size = 5.0, fontface = "bold"
    ) +
    ggplot2::scale_size_identity(guide = "none") +
    ggplot2::coord_cartesian(xlim = xlim, ylim = ylim, expand = FALSE, clip = "on") +
    ggplot2::labs(x = axis_labs$x, y = axis_labs$y) +
    ggplot2::theme_classic(base_size = base_font_size) +
    ggplot2::theme(
      text = ggplot2::element_text(color = "black"),
      axis.title = ggplot2::element_text(face = "bold", size = base_font_size + 2),
      axis.text = ggplot2::element_text(color = "black", size = base_font_size),
      axis.line = ggplot2::element_line(linewidth = 0.8, color = "black"),
      axis.ticks = ggplot2::element_line(linewidth = 0.8, color = "black"),
      axis.ticks.length = grid::unit(0.18, "cm"),
      plot.margin = ggplot2::margin(8, 10, 8, 10)
    )
  p
}

parse_row_ids <- function(x) {
  x <- as.character(x)
  x <- x[!is.na(x) & nzchar(x)]
  if (length(x) == 0) return(integer(0))
  ids <- unlist(strsplit(x, ";", fixed = TRUE))
  ids <- trimws(ids)
  ids <- suppressWarnings(as.integer(ids))
  ids[is.finite(ids)]
}

pool_lnrr_one_trait <- function(dat_trait) {
  dat_trait <- dat_trait %>%
    dplyr::filter(is.finite(lnRR_value), is.finite(Var_lnRR), Var_lnRR > 0)
  k <- nrow(dat_trait)
  if (k == 0) {
    return(tibble::tibble(
      Estimate = NA_real_, SE = NA_real_, CI_lower = NA_real_, CI_upper = NA_real_,
      P_value = NA_real_, tau2 = NA_real_, I2 = NA_real_, k = 0,
      n_studies = 0, n_systems = 0, Model = "no_valid_LNRR"
    ))
  }
  if (k == 1) {
    est <- dat_trait$lnRR_value[1]
    se <- sqrt(dat_trait$Var_lnRR[1])
    return(tibble::tibble(
      Estimate = est, SE = se,
      CI_lower = est - 1.96 * se,
      CI_upper = est + 1.96 * se,
      P_value = 2 * stats::pnorm(-abs(est / se)),
      tau2 = NA_real_, I2 = NA_real_, k = k,
      n_studies = dplyr::n_distinct(dat_trait$StudyID),
      n_systems = dplyr::n_distinct(dat_trait$SystemID),
      Model = "single_effect"
    ))
  }

  res <- tryCatch(
    metafor::rma.uni(yi = lnRR_value, vi = Var_lnRR, data = dat_trait, method = "REML"),
    error = function(e) NULL
  )
  if (!is.null(res)) {
    return(tibble::tibble(
      Estimate = as.numeric(res$b[1]),
      SE = as.numeric(res$se[1]),
      CI_lower = as.numeric(res$ci.lb[1]),
      CI_upper = as.numeric(res$ci.ub[1]),
      P_value = as.numeric(res$pval[1]),
      tau2 = as.numeric(res$tau2),
      I2 = as.numeric(res$I2),
      k = k,
      n_studies = dplyr::n_distinct(dat_trait$StudyID),
      n_systems = dplyr::n_distinct(dat_trait$SystemID),
      Model = "REML_random_effects"
    ))
  }

  ## Fallback fixed-effect inverse-variance estimate if REML fails.
  w <- 1 / pmax(dat_trait$Var_lnRR, eps)
  est <- sum(w * dat_trait$lnRR_value, na.rm = TRUE) / sum(w, na.rm = TRUE)
  se <- sqrt(1 / sum(w, na.rm = TRUE))
  tibble::tibble(
    Estimate = est, SE = se,
    CI_lower = est - 1.96 * se,
    CI_upper = est + 1.96 * se,
    P_value = 2 * stats::pnorm(-abs(est / se)),
    tau2 = NA_real_, I2 = NA_real_, k = k,
    n_studies = dplyr::n_distinct(dat_trait$StudyID),
    n_systems = dplyr::n_distinct(dat_trait$SystemID),
    Model = "fixed_effect_fallback"
  )
}

make_companion_forest <- function(reg_dat, target_row, comparison_type, dataset_tag, metric) {
  if (nrow(reg_dat) == 0) return(NULL)

  x_ids <- parse_row_ids(reg_dat$X_RowIDs)
  y_ids <- suppressWarnings(as.integer(reg_dat$Y_RowID))
  y_ids <- y_ids[is.finite(y_ids)]
  ids <- unique(c(x_ids, y_ids))
  if (length(ids) == 0) return(NULL)

  forest_rows <- raw_clean %>%
    dplyr::filter(RowID %in% ids, valid_lnrr) %>%
    dplyr::mutate(
      Trait = std_trait(Trait),
      ForestRole = dplyr::if_else(Trait == target_row$XTrait, "nutrient", "defense")
    )

  if (nrow(forest_rows) == 0) return(NULL)

  wanted_traits <- unique(c(target_row$XTrait, reg_dat$YTrait))
  forest_rows <- forest_rows %>% dplyr::filter(Trait %in% wanted_traits)
  if (nrow(forest_rows) == 0) return(NULL)

  forest_summary <- forest_rows %>%
    dplyr::group_by(Trait) %>%
    dplyr::group_modify(~ pool_lnrr_one_trait(.x)) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(
      ComparisonType = comparison_type,
      Dataset = dataset_tag,
      MetricUsedForForest = "LNRR",
      AnalysisOrder = target_row$AnalysisOrder,
      AnalysisID = target_row$AnalysisID,
      XTrait = target_row$XTrait,
      YLabel = target_row$YLabel,
      YTargetTrait = target_row$YTrait,
      Role = dplyr::if_else(Trait == target_row$XTrait, "nutrient", "defense")
    ) %>%
    dplyr::filter(is.finite(Estimate), is.finite(CI_lower), is.finite(CI_upper))

  if (nrow(forest_summary) == 0) return(NULL)

  trait_order <- c(target_row$XTrait, sort(setdiff(unique(forest_summary$Trait), target_row$XTrait)))
  trait_n <- forest_summary$k[match(trait_order, as.character(forest_summary$Trait))]
  trait_levels <- rev(paste0(trait_order, "
(n=", trait_n, ")"))

  forest_summary <- forest_summary %>%
    dplyr::mutate(
      n = k,
      SignifStar = dplyr::case_when(
        is.finite(P_value) & P_value < 0.001 ~ "***",
        is.finite(P_value) & P_value < 0.01  ~ "**",
        is.finite(P_value) & P_value < 0.05  ~ "*",
        TRUE ~ ""
      ),
      Trait = factor(Trait, levels = rev(trait_order)),
      TraitLabel = paste0(as.character(Trait), "
(n=", n, ")"),
      TraitLabel = factor(TraitLabel, levels = trait_levels)
    )

  ## Display range for the companion forest plot.
  ## Full estimates and 95% CIs are shown. No horizontal CI clipping is applied.
  ## A small right-side star zone is added so significance stars do not overlap CI bars.
  xvals <- c(forest_summary$CI_lower, forest_summary$CI_upper, forest_summary$Estimate, 0)
  xr <- range(xvals[is.finite(xvals)], na.rm = TRUE)
  if (!all(is.finite(xr)) || diff(xr) == 0) xr <- c(-1, 1)
  xspan <- diff(xr)
  if (!is.finite(xspan) || xspan <= 0) xspan <- 1
  xpad_left  <- 0.16 * xspan
  xpad_right <- 0.34 * xspan
  xlim <- c(xr[1] - xpad_left, xr[2] + xpad_right)
  star_x <- xr[2] + 0.16 * xspan

  forest_summary <- forest_summary %>%
    dplyr::mutate(
      CI_lower_plot = CI_lower,
      CI_upper_plot = CI_upper,
      cap_h = 0.045,
      y_num = as.numeric(TraitLabel),
      StarX = star_x
    )

  p <- ggplot2::ggplot(forest_summary, ggplot2::aes(y = TraitLabel, x = Estimate)) +
    ggplot2::geom_vline(xintercept = 0, linetype = "dashed", linewidth = 0.50, color = "grey72") +
    ggplot2::geom_segment(ggplot2::aes(x = CI_lower_plot, xend = CI_upper_plot, yend = TraitLabel),
                          linewidth = 0.78, color = "black") +
    ggplot2::geom_segment(ggplot2::aes(x = CI_lower_plot, xend = CI_lower_plot,
                                       y = y_num - cap_h, yend = y_num + cap_h),
                          inherit.aes = FALSE, linewidth = 0.78, color = "black") +
    ggplot2::geom_segment(ggplot2::aes(x = CI_upper_plot, xend = CI_upper_plot,
                                       y = y_num - cap_h, yend = y_num + cap_h),
                          inherit.aes = FALSE, linewidth = 0.78, color = "black") +
    ggplot2::geom_point(shape = 21, size = 2.8, stroke = 0.35, fill = "grey45", color = "black") +
    ggplot2::geom_text(ggplot2::aes(x = StarX, label = SignifStar),
                       color = "#D7191C", fontface = "bold", size = 4.8, hjust = 0.5, vjust = 0.45) +
    ggplot2::scale_x_continuous(breaks = scales::pretty_breaks(n = 3), labels = scales::label_number(accuracy = 0.01), guide = ggplot2::guide_axis(check.overlap = TRUE)) +
    ggplot2::coord_cartesian(xlim = xlim, clip = "on") +
    ggplot2::scale_y_discrete(labels = forest_trait_labels) +
    ggplot2::labs(x = "Effect size (lnRR)", y = NULL) +
    ggplot2::theme_classic(base_size = base_font_size) +
    ggplot2::theme(
      text = ggplot2::element_text(color = "black"),
      axis.title.x = ggplot2::element_text(face = "bold", size = base_font_size + 1),
      axis.text.x = ggplot2::element_text(color = "black", size = base_font_size - 1),
      axis.text.y = ggplot2::element_text(color = "black", size = base_font_size - 3, face = "bold", lineheight = 0.82),
      axis.line = ggplot2::element_line(linewidth = 0.8, color = "black"),
      axis.ticks = ggplot2::element_line(linewidth = 0.8, color = "black"),
      axis.ticks.length = grid::unit(0.14, "cm"),
      plot.margin = ggplot2::margin(8, 2, 8, 4)
    )

  list(plot = p, summary = forest_summary, n_forest_rows = nrow(forest_summary))
}

make_combined_panel <- function(forest_plot, regression_plot) {
  if (is.null(regression_plot)) return(NULL)
  if (is.null(forest_plot)) return(regression_plot)
  forest_plot + regression_plot + patchwork::plot_layout(widths = c(0.95, 2.50))
}

# 4) Read and prepare raw data
# =========================
infile <- input_file
raw <- read_input_table(infile, sheet_name)

name_map <- c(
  "StudyID" = "StudyID", "Study_Id" = "StudyID", "Number" = "StudyID",
  "SystemID" = "SystemID", "System_Id" = "SystemID",
  "ComparisonType" = "ComparisonType", "ComparsionType" = "ComparisonType", "Comparison" = "ComparisonType",
  "InvasiveSpecies" = "InvasiveSpecies", "InvasiveSpec" = "InvasiveSpecies",
  "ComparatorSpecies" = "ComparatorSpecies", "ComparatorSpec" = "ComparatorSpecies",
  "Trait" = "Trait", "TraitCategory" = "TraitCategory", "TraitForm" = "TraitForm",
  "Treatment" = "Treatment", "Tissue" = "Tissue", "ExperimentType" = "ExperimentType",
  "PairDesign" = "PairDesign", "DataLevel" = "DataLevel", "PairingLevel" = "PairingLevel",
  "UseForRegression" = "UseForRegression",
  "Mc" = "Mc", "SDc" = "SDc", "Nc" = "Nc",
  "Mi" = "Mi", "SDi" = "SDi", "Sdi" = "SDi", "Ni" = "Ni",
  "Unit" = "Unit", "FigureTable" = "FigureTable", "FigTable" = "FigureTable",
  "Notes" = "Notes", "Note" = "Notes"
)
for (old in names(name_map)) {
  if (old %in% names(raw) && !(name_map[[old]] %in% names(raw))) {
    names(raw)[names(raw) == old] <- name_map[[old]]
  }
}

required <- c("StudyID", "SystemID", "ComparisonType", "InvasiveSpecies", "ComparatorSpecies",
              "Trait", "TraitCategory", "Treatment", "Tissue", "ExperimentType", "PairDesign",
              "DataLevel", "PairingLevel", "UseForRegression", "Mc", "SDc", "Nc", "Mi", "SDi", "Ni")
missing_cols <- setdiff(required, names(raw))
if (length(missing_cols) > 0) stop("Missing required columns: ", paste(missing_cols, collapse = ", "))
if (!("TraitForm" %in% names(raw))) raw$TraitForm <- NA_character_
if (!("Unit" %in% names(raw))) raw$Unit <- NA_character_
if (!("FigureTable" %in% names(raw))) raw$FigureTable <- NA_character_
if (!("Notes" %in% names(raw))) raw$Notes <- NA_character_

raw_clean <- raw %>%
  dplyr::mutate(
    RowID = dplyr::row_number(),
    StudyID = clean_key(.data$StudyID),
    SystemID = clean_key(.data$SystemID),
    ComparisonType = std_text(.data$ComparisonType),
    InvasiveSpecies = as.character(.data$InvasiveSpecies),
    ComparatorSpecies = as.character(.data$ComparatorSpecies),
    Trait = std_trait(.data$Trait),
    TraitCategory = std_text(.data$TraitCategory),
    TraitForm = std_text(.data$TraitForm),
    Treatment = std_text(.data$Treatment),
    Tissue = std_text(.data$Tissue),
    ExperimentType = std_text(.data$ExperimentType),
    PairDesign = std_text(.data$PairDesign),
    DataLevel = std_text(.data$DataLevel),
    PairingLevel = std_text(.data$PairingLevel),
    UseForRegression = std_text(.data$UseForRegression),
    Mc = as.numeric(.data$Mc), SDc = as.numeric(.data$SDc), Nc = as.numeric(.data$Nc),
    Mi = as.numeric(.data$Mi), SDi = as.numeric(.data$SDi), Ni = as.numeric(.data$Ni),
    Unit = as.character(.data$Unit),
    FigureTable = as.character(.data$FigureTable),
    Notes = as.character(.data$Notes)
  ) %>%
  dplyr::mutate(
    TraitCategory = dplyr::if_else(force_lignin_to_physical & Trait == "Lignin", "physical_defense", TraitCategory),
    TraitForm = dplyr::case_when(
      Trait %in% c("C", "N", "P") ~ "material",
      Trait %in% structural_material_traits ~ "material",
      Trait %in% c("Phenolics", "Tannins", "Flavonoids", "Terpenoids", "Saponins", "Glycosides", "Alkaloids", "Glucosinolates", "Proteinaceous toxins") ~ "material",
      TRUE ~ TraitForm
    ),
    d_value = Mi - Mc,
    Var_d = var_delta(SDc, Nc, SDi, Ni),
    valid_diff = is.finite(Mi) & is.finite(Mc) & is.finite(SDi) & is.finite(SDc) &
      is.finite(Ni) & is.finite(Nc) & Ni > 0 & Nc > 0 & SDi >= 0 & SDc >= 0 &
      is.finite(d_value) & is.finite(Var_d) & Var_d >= 0,
    lnRR_value = dplyr::if_else(Mi > 0 & Mc > 0, log(Mi / Mc), NA_real_),
    Var_lnRR = var_lnrr(Mc, SDc, Nc, Mi, SDi, Ni),
    valid_lnrr = is.finite(lnRR_value) & is.finite(Var_lnRR) & Var_lnRR >= 0 &
      Mi > 0 & Mc > 0 & Ni > 0 & Nc > 0 & SDi >= 0 & SDc >= 0
  )

# =========================
# 5) Build analysis datasets

get_axis_labels <- function(comparison_type, x_trait, y_label, metric) {
 list(x=nutrient_axis_label(x_trait), y=paste0("lnRR-", y_label))
}

min_effects_for_bias_tests <- 10
run_begg_ranktest <- TRUE
run_trimfill_for_funnel <- TRUE

make_funnel_title <- function(metric, target_row, forest_trait) {
  prefix <- if (identical(metric, "DIFF")) "Cha" else "LNRR"
  paste0(prefix, " ", target_row$XTrait, " VS. ", target_row$YLabel, "-", forest_trait)
}

make_funnel_data_one_trait <- function(reg_dat, target_row, forest_trait, metric) {
  x_ids <- parse_row_ids(reg_dat$X_RowIDs)
  y_ids <- suppressWarnings(as.integer(reg_dat$Y_RowID))
  y_ids <- y_ids[is.finite(y_ids)]
  ids <- unique(c(x_ids, y_ids))
  if (length(ids) == 0) return(NULL)

  fdat <- raw_clean %>%
    dplyr::filter(RowID %in% ids, valid_lnrr) %>%
    dplyr::mutate(
      Trait = std_trait(Trait),
      SE_lnRR = sqrt(Var_lnRR)
    ) %>%
    dplyr::filter(Trait == forest_trait,
                  is.finite(lnRR_value), is.finite(Var_lnRR), Var_lnRR > 0,
                  is.finite(SE_lnRR), SE_lnRR > 0)

  if (nrow(fdat) < 3) return(NULL)

  model_obj <- tryCatch(
    metafor::rma.uni(yi = lnRR_value, vi = Var_lnRR, data = fdat, method = "REML"),
    error = function(e) NULL
  )
  if (is.null(model_obj)) {
    model_obj <- tryCatch(
      metafor::rma.uni(yi = lnRR_value, vi = Var_lnRR, data = fdat, method = "FE"),
      error = function(e) NULL
    )
  }
  if (is.null(model_obj)) return(NULL)

  pooled <- pool_lnrr_one_trait(fdat)

  egger_p <- NA_real_
  begg_p <- NA_real_
  trimfill_k0 <- NA_real_

  if (nrow(fdat) >= min_effects_for_bias_tests) {
    egger_p <- tryCatch(as.numeric(metafor::regtest(model_obj, model = "rma")$pval), error = function(e) NA_real_)
    if (isTRUE(run_begg_ranktest)) {
      begg_p <- tryCatch(as.numeric(metafor::ranktest(model_obj)$pval), error = function(e) NA_real_)
    }
    if (isTRUE(run_trimfill_for_funnel)) {
      tf <- tryCatch(metafor::trimfill(model_obj), error = function(e) NULL)
      if (!is.null(tf)) trimfill_k0 <- suppressWarnings(as.numeric(tf$k0))
    }
  }

  list(
    data = fdat,
    model = model_obj,
    pooled = pooled,
    title = make_funnel_title(metric, target_row, forest_trait),
    egger_p = egger_p,
    begg_p = begg_p,
    trimfill_k0 = trimfill_k0,
    n_effect_sizes = nrow(fdat),
    k_studies = dplyr::n_distinct(fdat$StudyID)
  )
}

plan <- manifest[manifest$Figure == "F3", ]
base <- make_analysis_base("invasive_vs_native", "LNRR")
assert_system_study_keys(base)
output <- vector("list", nrow(plan))
for (i in seq_len(nrow(plan))) {
 spec <- plan[i, ]
 message("Fitting ", spec$Panel, " (", spec$Dataset, ")")
 row <- tibble::tibble(AnalysisOrder=spec$AnalysisOrder, AnalysisID=spec$AnalysisID,
   XTrait=spec$XTrait, TargetType="specific_trait", YTrait=spec$YTrait,
   YCategory=spec$YCategory, YLabel=spec$YTrait)
 dat_raw <- build_fit_data(base, row, "invasive_vs_native", "LNRR")
 dat <- if (spec$Dataset == "IQR_removed") iqr_filter_2d(dat_raw, k=iqr_k) else dat_raw
 validate_fit_input(dat, spec)
 dat$dataset <- spec$Dataset
 cand <- fit_candidate_models(dat, row, "invasive_vs_native", "LNRR", spec$Dataset)
 decision <- make_selection_decision(cand$linear$summary, cand$quadratic$summary, "specific_trait", "invasive_vs_native")
 if (is.na(decision$SelectedModelForm) || decision$SelectedModelForm != spec$ModelForm) {
   stop(spec$Panel, ": the original selection rule no longer reproduces the manuscript model. Review the input data; no substitute panel was exported.")
 }
 res <- cand[[spec$ModelForm]]$result
 plot <- build_plot(dat, res, "invasive_vs_native", row, spec$Dataset, "LNRR", spec$ModelForm)
 forest <- make_companion_forest(dat, row, "invasive_vs_native", spec$Dataset, "LNRR")
 stopifnot(!is.null(plot), !is.null(forest), nrow(forest$summary)==2L)
 summary <- model_summary_row(spec, dat, res$fit, res$R2)
 summary$OriginalSelectionStatus <- decision$SelectionStatus
 summary$OriginalSelectionReason <- decision$SelectionReason
 forest$summary$Panel <- spec$Panel
 forest$summary$SelectedModelForm <- spec$ModelForm
 decision$Panel <- spec$Panel
 output[[i]] <- panel_result(spec, dat_raw, dat, res$fit, make_prediction_band(res$fit,dat),
   make_combined_panel(forest$plot, plot), summary)
 output[[i]]$forest_summary <- forest$summary
 ids <- unique(c(parse_row_ids(dat$X_RowIDs), dat$Y_RowID))
 output[[i]]$forest_data <- raw_clean[raw_clean$RowID %in% ids & raw_clean$valid_lnrr, ]
 output[[i]]$forest_data$Panel <- spec$Panel
 output[[i]]$forest_data$SourceExcelRow <- output[[i]]$forest_data$RowID + input_info$header_row
 output[[i]]$selection <- decision
 output[[i]]$funnels <- list()
 for(trait in c(spec$XTrait,spec$YTrait)) {
   funnel_warnings <- character()
   obj <- withCallingHandlers(make_funnel_data_one_trait(dat,row,trait,"LNRR"),
     warning=function(w) { funnel_warnings <<- c(funnel_warnings,conditionMessage(w)) })
   if(is.null(obj)) stop(spec$Panel, ": cannot construct the funnel for ",trait)
   obj$warnings <- unique(funnel_warnings)
   expected_ids <- output[[i]]$forest_data$RowID[
     output[[i]]$forest_data$Trait==trait & is.finite(output[[i]]$forest_data$Var_lnRR) & output[[i]]$forest_data$Var_lnRR>0]
   if(!setequal(obj$data$RowID,expected_ids)) stop(spec$Panel,": funnel/forest source rows differ.")
   fs <- forest$summary[as.character(forest$summary$Trait)==trait,]
   for(field in c("Estimate","SE","CI_lower","CI_upper","P_value","tau2","I2","k","n_studies","n_systems")) {
     if(!isTRUE(all.equal(as.numeric(obj$pooled[[field]]),as.numeric(fs[[field]]),tolerance=1e-10)))
       stop(spec$Panel,": funnel/forest statistic differs: ",field)
   }
   output[[i]]$funnels[[trait]] <- obj
 }
}
output

}

if(!file.exists(shared_input_file)) stop("Cannot find input workbook: ",shared_input_file)
shared_input_file <- normalizePath(shared_input_file,winslash="/",mustWork=TRUE)
if(!is.null(output_root_override)) output_root <- output_root_override
ensure_output_dir(output_root)
# Direct output folders. Re-running replaces the same named output files.
run_dir <- normalizePath(output_root,winslash="/",mustWork=TRUE)
manifest <- tibble::tribble(
 ~Panel, ~Figure, ~AnalysisID, ~XTrait, ~YCategory, ~YTrait, ~Dataset, ~ModelForm, ~AnalysisOrder,
 "F2_A","F2","C_physical_total","C","physical_defense",NA_character_,"RAW","quadratic",1L,
 "F2_B","F2","C_chemical_total","C","chemical_defense",NA_character_,"IQR_removed","quadratic",2L,
 "F2_C","F2","N_chemical_total","N","chemical_defense",NA_character_,"RAW","quadratic",3L,
 "F2_D","F2","P_chemical_total","P","chemical_defense",NA_character_,"IQR_removed","quadratic",4L,
 "F3_A","F3","C_vs_Phenolics","C",NA_character_,"Phenolics","RAW","linear",8L,
 "F3_B","F3","C_vs_Lignin","C",NA_character_,"Lignin","RAW","linear",4L,
 "F3_C","F3","N_vs_Phenolics","N","chemical_defense","Phenolics","IQR_removed","linear",17L,
 "F3_D","F3","N_vs_Tannins","N","chemical_defense","Tannins","RAW","quadratic",18L
)
current_stage <- "output preflight"
captured_warnings <- character()
started_at <- format(Sys.time(),"%Y-%m-%d %H:%M:%S %Z")
message("Input: ",shared_input_file)
message("Output folder: ",run_dir)
info_path <- file.path(run_dir,"run_info.txt")
expected_main <- c(file.path("F2",c("F2.pdf",paste0("F2_",LETTERS[1:4],".pdf"))),
                   file.path("F3",c("F3.pdf",paste0("F3_",LETTERS[1:4],".pdf"))))
plan3 <- manifest[manifest$Figure=="F3",]
funnel_names <- unlist(lapply(seq_len(nrow(plan3)),function(i)
 paste0(plan3$Panel[i],"_",c(plan3$XTrait[i],plan3$YTrait[i]),"_funnel.pdf")))
expected_pdf <- c(expected_main,file.path("F3_funnel",c(funnel_names,"SF5.pdf")))
expected_png <- sub("[.]pdf$",".png",expected_pdf)
expected_tables <- c("F2/F2_results.xlsx","F2/F2_model_summary.csv",
 "F3/F3_results.xlsx","F3/F3_model_summary.csv",
 "F3_funnel/F3_funnel_results.xlsx","F3_funnel/F3_funnel_summary.csv",
 "F3_funnel/F3_funnel_diagnostics.txt","F2_F3_model_summary.csv","panel_manifest.csv")
tryCatch(withCallingHandlers({
 check_existing_output(info_path)
 safe_write_lines(c("RUNNING",paste0("Started: ",started_at),paste0("Input: ",shared_input_file),
   "Existing named results will be replaced. Use this folder only after this file reports COMPLETE."),info_path)
 probe_output_dir(run_dir,35L)
 for(fig in c("F2","F3","F3_funnel")) probe_output_dir(file.path(run_dir,fig),45L)
 for(path in c(expected_pdf,expected_png,expected_tables)) check_existing_output(file.path(run_dir,path))
 current_stage <- "input inspection"
 input_info <- inspect_input(shared_input_file,shared_sheet_name)
 current_stage <- "F2 fitting"
 f2 <- run_f2(shared_input_file,shared_sheet_name,manifest)
 current_stage <- "F3 fitting, forest and funnel diagnostics"
 f3 <- run_f3(shared_input_file,shared_sheet_name,manifest)
 current_stage <- "F2 figure and table export"
 summary2 <- export_figure(f2,"F2",run_dir)
 current_stage <- "F3 figure and table export"
 summary3 <- export_figure(f3,"F3",run_dir)
 current_stage <- "F3 individual funnels and SF5 assembly"
 funnel_summary <- export_f3_funnels(f3,run_dir)
 current_stage <- "final verification"
 if(!all(file.exists(file.path(run_dir,c(expected_pdf,expected_png,expected_tables[1:7])))))
   stop("One or more expected figure/table exports are missing.")
 if(!identical(unname(tools::md5sum(shared_input_file)),input_info$md5)) stop("Input file changed during this run.")
 manifest$InputFile <- basename(shared_input_file)
 manifest$InputMD5 <- input_info$md5
 manifest$HeaderRow <- input_info$header_row
 manifest$PlotFile <- paste0(manifest$Figure,"/",manifest$Panel,".pdf")
 safe_write_csv(manifest,file.path(run_dir,"panel_manifest.csv"),row.names=FALSE,na="",fileEncoding="UTF-8")
 safe_write_csv(dplyr::bind_rows(summary2,summary3),file.path(run_dir,"F2_F3_model_summary.csv"),row.names=FALSE,na="",fileEncoding="UTF-8")
 safe_write_lines(c("COMPLETE",paste0("Started: ",started_at),
   paste0("Finished: ",format(Sys.time(),"%Y-%m-%d %H:%M:%S %Z")),
   paste0("Input: ",shared_input_file),paste0("Input location method: ",located$method),
   paste0("Input MD5: ",input_info$md5),paste0("Header row: ",input_info$header_row),
   paste0("Imported data rows: ",input_info$data_rows),
   "Input workbook was read only.",
   "Exports: F2/F3 eight panels + two composites; eight F3 trait funnels + one SF5 assembly.",
   "19 PDFs and 19 matching PNGs. Stable filenames are replaced on each rerun.",
   "lnRR=log(invasive/native); no signed or z-score standardization.",
   "WLS weight=1/max(response lnRR variance,1e-12); weight capping OFF.",
   "F2 join: SystemID+StudyID; F3 join: SystemID (original separate rules).",
   "IQR: inclusive 1.5 IQR bounds on both axes, after pairing; F2_B,F2_D,F3_C.",
   "IQR_Audit preserves row membership for reproducibility.",
   "Original regression CI display trimming retained; Prediction_Data contains full mean-curve confidence limits.",
   "F3: left forest stars test pooled lnRR; right regressions display R-squared only.",
   "Funnels use the same positive-variance original RowIDs as the corresponding forest trait.",
   "SF5 is a 3x3 arrangement of eight separate funnels, in order F3A X/Y, F3B X/Y, F3C X/Y, F3D X/Y.",
   "Egger/Begg P values are funnel diagnostics. Trimfill k0 is reported without adding imputed points to plots or forests.",
   "Begg ties warnings, when present, indicate that the returned P value uses the original approximate calculation.",
   "Funnel contour key: original white p>=0.10; grey90 0.05<=p<0.10; grey80 p<0.05. The p=0.01 contour line remains, with the same grey80 on both sides.",
   "", "Warnings captured:",
   if(length(captured_warnings)) captured_warnings else "No warnings captured.",
   "", "Session information:", capture.output(sessionInfo())),info_path)
 message("Completed F2, F3 and SF5. Output folder: ",run_dir)
},warning=function(w) {
 captured_warnings <<- c(captured_warnings,paste0("[",current_stage,"] ",conditionMessage(w)))
}),error=function(e) {
 original <- conditionMessage(e)
 diagnostic <- c("FAILED",paste0("Started: ",started_at),paste0("Stage: ",current_stage),
   paste0("Original error: ",original),
   "This attempt did not complete. Files in this folder may be from different attempts; rerun successfully before using them.",
   captured_warnings)
 tryCatch(safe_write_lines(diagnostic,info_path),error=function(io) message("Could not save diagnostic: ",conditionMessage(io)))
 stop(original,"\nStage: ",current_stage,"\nOutput folder: ",run_dir,call.=FALSE)
})
})
