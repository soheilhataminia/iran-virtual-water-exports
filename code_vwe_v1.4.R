# =============================================================================
# IRAN CROP-BASED EXPORTS AND VIRTUAL WATER: REPRODUCIBLE ANALYSIS
# Version 1.4 | Input schema: IRAN_VWT_RAW_1.0 | R >= 4.2
# Retains F05 top fifteen water shares and F17 product-group shares.
# Version 1.4: accurate omit-2023 window labels, export-unit-value terminology,
# and identity-checked aggregation of within into unit-value/footprint effects.
# Startup fix: select the workbook first; retry package installation through
# alternative repositories; request Windows binaries without a compiler prompt.
#
# RStudio: open this script and click Source. Select the supplied raw .xlsx file.
# Console: source(file.choose(), encoding = "UTF-8")  # first select THIS script
# Command line: Rscript Iran_Virtual_Water_Analysis_v1_4.R path/to/raw.xlsx
# Optional: options(vwt.input = "path/to/raw.xlsx", vwt.output = "new/directory")
#
# All transformations occur below. The input workbook is never overwritten.
# First use requires internet to install missing analysis packages. Later runs are
# offline. No original result tables, target coefficients or target signs are used.
#
# ESTIMAND: gross FAOSTAT crop-based exports and their estimated virtual-water exports
# (green+blue water) using CWASI supply-side coefficients. Imported water can be
# embodied in re-exports. This is NOT domestic blue-water depletion, scarcity-
# weighted welfare, profit, or a causal response to a water-policy intervention.
#
# PRIMARY DIRECTION matches the manuscript: log(V) on log(X). The reverse
# regression is separately labelled and is NOT the reciprocal of this slope.
# The shared quantity in X = unit_value * Q and V = footprint * Q matters for the
# log(V)-log(X) association. Unit-value/footprint and unrestricted Q/UV models below
# diagnose that identity; none supplies an instrument or causal identification.
# =============================================================================

VWT_CONFIG <- list(
  start_year = 1986L, end_year = 2023L, base_year = 2020L,
  excluded_aggregate_codes = 30L, min_core_years = 30L,
  auto_install = TRUE, seed = 20260906L, bootstrap_replications = 999L,
  wild_bootstrap_replications = 999L, block_length = 3L,
  run_leave_one_product_out = TRUE,
  run_ppml = TRUE, run_ihs = TRUE, run_bootstrap = TRUE,
  run_wild_bootstrap = TRUE, run_stationarity_diagnostics = TRUE,
  coefficient_sensitivity = c(0.10, 0.20), figure_dpi = 400L,
  nthreads = 1L
)

vwt_run <- function(config = VWT_CONFIG) {
  if (getRversion() < "4.2.0") stop("Use R 4.2 or newer.")
  # Ask for the workbook BEFORE any package check or download, so an install
  # failure cannot conceal the input-selection step. This uses only base R.
  message("[Startup 1/2] Select the raw Excel workbook.")
  args <- commandArgs(trailingOnly = TRUE)
  input <- getOption("vwt.input", NULL)
  if (is.null(input) && length(args) && file.exists(args[1L])) input <- args[1L]
  if (is.null(input)) {
    message("Select Iran_Virtual_Water_Raw_Extract.xlsx in the file-selection window.")
    input <- tryCatch(file.choose(),error=function(e)stop("Workbook selection was cancelled or failed. Run vwt_run() again and select the .xlsx file.",call.=FALSE))
  }
  input <- normalizePath(input, mustWork = TRUE)
  if (!grepl("\\.xlsx$", input, ignore.case = TRUE)) stop("Select the raw .xlsx workbook, not the R script.")
  message("Workbook selected: ",input)
  message("[Startup 2/2] Checking analysis packages; first-time downloads may take a few minutes.")
  needed <- c("data.table", "readxl", "fixest", "ggplot2", "writexl", "sandwich")
  minimum <- c(data.table="1.14.0", readxl="1.4.0", fixest="0.14.2", ggplot2="3.5.0", writexl="1.4.0", sandwich="3.0.0")
  ready <- function(p) tryCatch(packageVersion(p)>=package_version(minimum[[p]]),error=function(e)FALSE)
  stale_loaded <- needed[vapply(needed,function(p)p%in%loadedNamespaces() && getNamespaceVersion(p)<package_version(minimum[[p]]),logical(1))]
  if(length(stale_loaded))stop("Restart R before updating these already-loaded packages: ",paste(stale_loaded,collapse=", "))
  missing <- needed[!vapply(needed, ready, logical(1))]
  if (length(missing)) {
    if (!isTRUE(config$auto_install)) stop("Install these CRAN packages first: ", paste(missing, collapse = ", "))
    message("Installing missing analysis packages: ", paste(missing, collapse = ", "))
    old_download_options <- options(timeout = max(60, getOption("timeout", 60)))
    on.exit(options(old_download_options),add=TRUE)
    # The maintainer publishes Windows R 4.6 binaries, including fixest 0.14.2:
    # https://lrberge.r-universe.dev/fixest
    # Required dependencies are available via the accompanying CRAN mirror.
    repository_attempts <- list(
      c("https://lrberge.r-universe.dev", "https://cran.r-universe.dev"),
      "https://packagemanager.posit.co/cran/latest",
      "https://cloud.r-project.org"
    )
    install_type <- if(.Platform$OS.type=="windows")"binary" else getOption("pkgType","source")
    for(repo in repository_attempts){
      missing <- needed[!vapply(needed,ready,logical(1))]
      if(!length(missing))break
      message("Trying package repository: ",paste(repo,collapse=" ; "))
      tryCatch(install.packages(missing,repos=repo,dependencies=NA,type=install_type),
               error=function(e)message("Installation attempt failed: ",conditionMessage(e)))
    }
  }
  missing <- needed[!vapply(needed, ready, logical(1))]
  if (length(missing)) stop("Package installation incomplete: ", paste(missing, collapse = ", "),
    ". Your workbook was selected, but analysis has not started. Check the download messages above. An inaccessible PACKAGES index does not establish R-version incompatibility. Restore access to a repository, then rerun this script.",call.=FALSE)
  suppressPackageStartupMessages(library(data.table))
  suppressPackageStartupMessages(library(ggplot2))
  set.seed(config$seed)
  data.table::setDTthreads(config$nthreads)
  fixest::setFixest_nthreads(config$nthreads)
  stamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  out <- getOption("vwt.output", file.path(dirname(input), paste0("Iran_VWT_Results_", stamp)))
  if (dir.exists(out) && length(list.files(out, all.files = FALSE))) stop("Output directory is not empty. Choose a new vwt.output or leave it unset.")
  dir.create(out, recursive = TRUE, showWarnings = FALSE)
  out <- normalizePath(out)
  dirs <- setNames(file.path(out, c("tables", "figures", "data", "models", "audit")), c("tables", "figures", "data", "models", "audit"))
  invisible(lapply(dirs, dir.create, showWarnings = FALSE))
  logcon <- file(file.path(out, "run_log.txt"), open = "wt", encoding = "UTF-8")
  sink(logcon, split = TRUE)
  on.exit({sink(); close(logcon)}, add = TRUE)
  cat("INPUT:", input, "\nOUTPUT:", out, "\nSTART:", as.character(Sys.time()), "\n")
  cat("Workbook MD5:", unname(tools::md5sum(input)), "\n")
  dput(config, file = file.path(dirs["audit"], "analysis_config.R"))
  warnings_log <- data.table(stage = character(), message = character())
  record_warning <- function(stage, msg) {
    warnings_log <<- rbind(warnings_log, data.table(stage = stage, message = as.character(msg)))
    cat("[NOTE]", stage, ":", msg, "\n")
  }
  safe <- function(stage, expr) {
    tryCatch(withCallingHandlers(expr, warning = function(w) {
      record_warning(stage, conditionMessage(w)); invokeRestart("muffleWarning")
    }), error = function(e) {record_warning(stage, paste("NOT ESTIMABLE:", conditionMessage(e))); NULL})
  }
  tables <- list(); table_notes <- list(); model_store <- list(); model_rows <- list()
  figure_notes <- data.table(file = character(), caption = character())
  figure_caption_catalog <- c(
    F04_concentration="Herfindahl concentration of annual water and export-value shares in the eligible basket; source data: T09.",
    F05_top_products="Fifteen products with the highest full-period virtual-water export shares, based on pooled totals; source data: T08.",
    F06_bennet_components="Exact changes in the value-to-water ratio split into common-product within and composition terms and separate coverage entry/exit contributions. Coverage changes need not be actual market entry or exit; source data: T11.",
    F07_product_contributions="Fifteen largest absolute product contributions to the endpoint ratio change, sorted by signed net contribution; source data: T12, observed mode, 1986-1990 to 2021-2023.",
    F08_chained_bennet="Adjacent-year Bennet contributions, with common-product and coverage components separated. Annual and long-horizon attribution need not coincide; source data: T13.",
    F10_IHS_units="Water is the outcome and export value is the regressor. Both are divided by the stated scale before asinh transformation. Two-way clustered 95% intervals; these coefficients are not elasticities. Source data: T28.",
    F11_source_quantities="Positive FAOSTAT and CWASI-implied export quantities for matched product-years. Dashed line denotes agreement. Implied CWASI quantity is VWT divided by its supply coefficient; source data: audit/quantity_comparison_all.csv.",
    F12_source_ratio_comparison="Ratios calculated with identical export-value observations, using either FAOSTAT quantity times the supply footprint or original CWASI water. The difference reflects trade-quantity inconsistency, not causal effects; source data: T30.",
    F13_residual_diagnostics="Residuals versus fitted log water in the primary product/year fixed-effects model; source data: data/main_model_residuals.csv. The scatterplot does not by itself establish valid inference.",
    F14_fixed_basket="Annual ratios for the time-balanced product basket and the varying eligible basket. Product inclusion is fixed in the balanced series, while water shares still vary; source data: T31.",
    F15_product_influence="Each point omits one product and recomputes the primary log-water-on-log-value slope and endpoint ratio change. Dashed lines mark full-sample estimates; source data: T32.",
    F16_pistachio_unit_value_footprint="Pooled period real export unit values and quantity-weighted supply footprints for pistachios on separate axes. Values are accounting components, not causal productivity estimates; source data: T10, code 223."
  )
  table_save <- function(name, x, note = "") {
    if (is.null(x)) return(invisible(NULL))
    x <- as.data.table(x)
    tables[[name]] <<- as.data.frame(x)
    table_notes[[name]] <<- note
    fwrite(x, file.path(dirs["tables"], paste0(name, ".csv")), na = "NA", bom = TRUE)
    invisible(x)
  }
  valid_png <- function(path) {
    if(!file.exists(path) || file.info(path)$size < 100L)return(FALSE)
    con<-file(path,"rb");on.exit(close(con))
    sig<-readBin(con,"raw",8L)
    seek(con,where=file.info(path)$size-12L,origin="start")
    tail<-readBin(con,"raw",12L)
    identical(sig,as.raw(c(137,80,78,71,13,10,26,10))) &&
      identical(tail,as.raw(c(0,0,0,0,73,69,78,68,174,66,96,130)))
  }
  valid_pdf <- function(path) {
    if(!file.exists(path) || file.info(path)$size < 100L)return(FALSE)
    con<-file(path,"rb");on.exit(close(con))
    sig<-rawToChar(readBin(con,"raw",5L))
    seek(con,where=max(0,file.info(path)$size-64L),origin="start")
    end_bytes<-readBin(con,"raw",64L)
    while(length(end_bytes) && as.integer(tail(end_bytes,1L))%in%c(9L,10L,13L,32L))end_bytes<-head(end_bytes,-1L)
    identical(sig,"%PDF-") && identical(tail(end_bytes,5L),charToRaw("%%EOF"))
  }
  figure_save <- function(name, plot, width = 8, height = 5, caption = "") {
    if(!nzchar(caption))caption <- unname(figure_caption_catalog[name])
    if(length(caption)!=1L || is.na(caption) || !nzchar(caption))stop("Missing figure caption: ",name)
    plot <- plot + theme_bw(base_size = 11, base_family = "sans") +
      theme(panel.grid.minor = element_blank(), panel.grid.major = element_line(colour = "#E6E6E6", linewidth = 0.25),
            legend.position = "bottom", plot.title = element_blank(), plot.subtitle = element_blank(), plot.caption = element_blank())
    # Exact plotted data are saved separately wherever the plot is created.
    pdf_path <- file.path(dirs["figures"], paste0(name, ".pdf"))
    png_path <- file.path(dirs["figures"], paste0(name, ".png"))
    render <- function(path,kind) {
      # Own and explicitly close only this graphics device. Avoid reopening a
      # previous file device when this script is sourced inside another session.
      if(kind=="pdf")grDevices::pdf(file=path,width=width,height=height,useDingbats=FALSE)
      else grDevices::png(filename=path,width=width,height=height,units="in",res=config$figure_dpi,bg="white")
      device_id<-grDevices::dev.cur()
      on.exit({if(device_id%in%grDevices::dev.list())invisible(grDevices::dev.off(device_id))},add=TRUE)
      print(plot)
      invisible(NULL)
    }
    pdf_tmp<-tempfile(pattern="vwt_figure_",fileext=".pdf")
    png_tmp<-tempfile(pattern="vwt_figure_",fileext=".png")
    on.exit(unlink(c(pdf_tmp,png_tmp)),add=TRUE)
    render(pdf_tmp,"pdf")
    if(!valid_pdf(pdf_tmp))stop("PDF rendering failed: ",name)
    if(!file.copy(pdf_tmp,pdf_path,overwrite=TRUE))stop("Could not finalize PDF: ",name)
    # Validate the PNG signature and terminal IEND block before publishing the
    # file. One fresh render is allowed after an incomplete device write.
    for(attempt in 1:2){
      render(png_tmp,"png")
      if(valid_png(png_tmp))break
      record_warning(paste("Figure",name),paste("Incomplete PNG write; attempt",attempt))
    }
    if(!valid_png(png_tmp))stop("PNG rendering failed after retry: ",name)
    # Copy validated bytes so final outputs are independent of any temporary
    # file retained internally by the graphics device or its finalizer.
    if(!file.copy(png_tmp,png_path,overwrite=TRUE))stop("Could not finalize PNG: ",name)
    if(any(!file.exists(c(pdf_path,png_path))) || any(file.info(c(pdf_path,png_path))$size < 100L))
      stop("Figure export failed: ", name, ". Inspect graphics-device availability and rerun in a new output directory.")
    figure_notes <<- rbind(figure_notes, data.table(file = name, caption = caption))
    invisible(plot)
  }
  num <- function(x) suppressWarnings(as.numeric(if(is.factor(x))as.character(x) else x))
  safe_sum <- function(x) if (all(is.na(x))) NA_real_ else sum(x, na.rm = TRUE)
  period <- function(y) as.character(cut(y, c(1985,1990,2000,2010,2020,2023), labels = c("1986-1990","1991-2000","2001-2010","2011-2020","2021-2023")))
  assert_unique <- function(x, keys, label) {
    dup <- x[duplicated(x, by = keys) | duplicated(x, by = keys, fromLast = TRUE)]
    if (nrow(dup)) {fwrite(dup, file.path(dirs["audit"], paste0("DUPLICATES_", label, ".csv"))); stop("Duplicate source keys in ", label, "; inspect the audit file.")}
  }
  sheet <- function(n) as.data.table(readxl::read_excel(input, sheet = n, col_types = "text", .name_repair = "minimal"))
  required_sheets <- c("README","FAOSTAT_raw","CWASI_supply_raw","CWASI_production_raw","CWASI_VWT_raw","CWASI_items","CWASI_countries","CPI_raw","Source_manifest")
  absent <- setdiff(required_sheets, readxl::excel_sheets(input))
  if (length(absent)) stop("Wrong workbook schema. Missing sheets: ", paste(absent, collapse = ", "))
  guide <- sheet("README")
  if (!any(grepl("IRAN_VWT_RAW_1.0", unlist(guide), fixed = TRUE))) stop("Unsupported input schema.")
  fwrite(sheet("Source_manifest"), file.path(dirs["audit"], "source_manifest.csv"))
  fwrite(guide, file.path(dirs["audit"], "input_readme.csv"))

  # 1. SOURCE READING AND EXPLICIT, NON-DESTRUCTIVE TRANSFORMATION --------------
  cat("\n[1/9] Reading raw source sheets and building the product-year grid.\n")
  f <- sheet("FAOSTAT_raw")
  required_f <- c("Area Code", "Item Code", "Element Code", "Year", "Unit", "Value", "Flag", "Note", "Item")
  if (!all(required_f %in% names(f))) stop("FAOSTAT source headers are incomplete.")
  f[, `:=`(code = num(get("Item Code")), Year = num(Year), element = num(get("Element Code")), value = num(Value))]
  if (any(num(f[["Area Code"]]) != 102, na.rm = TRUE)) stop("FAOSTAT_raw contains a country other than Iran.")
  f <- f[Year >= config$start_year & Year <= config$end_year]
  assert_unique(f, c("code","Year","element"), "FAOSTAT")
  if (f[element == 5910, any(Unit != "t")]) stop("Export quantity must be in tonnes (t).")
  if (f[element == 5922, any(!Unit %in% c("1000 USD","1000 US$"))]) stop("Unexpected export-value unit; do not silently rescale.")
  q <- f[element == 5910, .(code, Year, Q = value, q_flag = Flag, q_note = Note, q_record = TRUE)]
  x <- f[element == 5922, .(code, Year, X_1000USD = value, x_flag = Flag, x_note = Note, fao_name = Item, x_record = TRUE)]
  coefficient <- function(n, value_name) {
    a <- sheet(n)
    yc <- names(a)[grepl("^[0-9]{4}$", names(a))]
    b <- melt(a, id.vars = "Item_code", measure.vars = yc, variable.name = "Year", value.name = value_name)
    setnames(b, "Item_code", "code")
    b[, `:=`(code = num(code), Year = num(Year))]
    b[, (value_name) := num(get(value_name))]
    assert_unique(b, c("code","Year"), n)
    b[Year >= config$start_year & Year <= config$end_year]
  }
  ws <- coefficient("CWASI_supply_raw", "W_supply")
  wp <- coefficient("CWASI_production_raw", "W_production")
  if(!nrow(ws)||!nrow(wp))stop("Coefficient years failed to decode; inspect the four-digit source year headers.")
  vr <- sheet("CWASI_VWT_raw")
  partners <- sprintf("P%03d", 1:255)
  if (!all(partners %in% names(vr))) stop("All 255 raw destination columns must be present.")
  mat <- as.matrix(vr[, lapply(.SD, num), .SDcols = partners])
  # No na.rm=TRUE: an unknown matrix row cannot become a zero trade flow.
  v <- data.table(code = num(vr$Item_code), Year = num(vr$Year),
                  V_CWASI = rowSums(mat, na.rm = FALSE),
                  missing_destinations = rowSums(is.na(mat)), source_file = vr$Source_file)
  if (any(vr$Exporter_FAO_code != "102")) stop("Unexpected exporter in CWASI raw rows.")
  assert_unique(v, c("code","Year"), "CWASI_VWT")
  rm(vr, mat); invisible(gc())
  item_raw <- sheet("CWASI_items")
  items <- item_raw[, .(code = num(get("FAO code")), item = get("Commodity name"), category = Category, primary = num(Primary))]
  assert_unique(items, "code", "CWASI_items")
  universe <- sort(unique(v$code))
  grid <- CJ(code = universe, Year = seq.int(config$start_year, config$end_year))
  d <- Reduce(function(a,b) merge(a,b, by = c("code","Year"), all.x = TRUE, sort = FALSE), list(grid, q, x, ws, wp, v))
  d <- merge(d, items, by = "code", all.x = TRUE, sort = FALSE)
  if (anyNA(d$item) || anyNA(d$category)) stop("An active product lacks source metadata; do not invent a category.")
  cpi_input <- sheet("CPI_raw")
  if(anyNA(cpi_input$Country_ISO3)||any(cpi_input$Country_ISO3!="USA")||
     anyNA(cpi_input$Indicator)||any(cpi_input$Indicator!="FP.CPI.TOTL"))stop("CPI must be USA FP.CPI.TOTL.")
  cpi <- cpi_input[, .(Year = num(Year), CPI = num(US_CPI_2010_100))]
  cpi_provenance <- if("Source_file"%in%names(cpi_input))paste("CPI: World Bank WDI; USA FP.CPI.TOTL; uploaded source:",paste(unique(cpi_input$Source_file),collapse="; "),
    "; publisher updated",paste(unique(cpi_input$Publisher_updated),collapse="; ")) else "CPI: World Bank WDI, USA FP.CPI.TOTL; see the input README and source manifest."
  fwrite(cpi_input,file.path(dirs["audit"],"cpi_source_extract.csv"),na="NA")
  assert_unique(cpi, "Year", "CPI")
  cpi0 <- cpi[Year == config$base_year, CPI]
  if (length(cpi0) != 1L || !is.finite(cpi0) || cpi0 <= 0) stop("Missing/invalid base-year CPI.")
  d <- merge(d, cpi, by = "Year", all.x = TRUE, sort = FALSE)
  if (any(!is.finite(d$CPI) | d$CPI <= 0)) stop("CPI must be positive and complete for all study years.")
  d[is.na(q_record), q_record := FALSE]; d[is.na(x_record), x_record := FALSE]
  d[, `:=`(X = X_1000USD * 1000 * cpi0 / CPI,
           V = Q * W_supply,
           Q_CWASI_implied = V_CWASI / W_supply,
           excluded_aggregate = code %in% config$excluded_aggregate_codes,
           period = period(Year), trend = Year - config$start_year)]
  d[, exclusion := fcase(excluded_aggregate, "aggregate_overlap", !q_record, "quantity_record_absent", !x_record, "value_record_absent",
                        !is.finite(Q) | !is.finite(X), "missing_numeric_trade", Q < 0 | X < 0, "negative_trade",
                        Q == 0 & X > 0, "positive_value_zero_quantity", Q == 0 & X == 0, "observed_no_trade",
                        !is.finite(W_supply) | W_supply <= 0, "missing_or_nonpositive_footprint", default = "eligible")]
  d[, `:=`(eligible = exclusion == "eligible", matched_original = V_CWASI > 0 & x_record & is.finite(X) & X >= 0 & !excluded_aggregate)]
  d[is.na(matched_original), matched_original := FALSE]
  d[, quantity_disagreement := fifelse(Q > 0 & Q_CWASI_implied > 0, abs(Q - Q_CWASI_implied) / pmax(Q, Q_CWASI_implied), NA_real_)]
  setorder(d, code, Year)
  fwrite(d, file.path(dirs["data"], "constructed_full_panel.csv"), na = "NA", bom = TRUE)
  base <- copy(d[eligible == TRUE])
  if (nrow(base) < 100 || uniqueN(base$code) < 10) stop("Insufficient eligible observations.")
  base[, `:=`(lnV = log(V), lnQ = log(Q), lnW = log(W_supply),
              lnX = fifelse(X > 0, log(pmax(X, .Machine$double.xmin)), NA_real_),
              X_million = X / 1e6, V_million = V / 1e6, export_unit_value = X / Q, e = X / V)]
  base[, `:=`(lnUV = fifelse(export_unit_value > 0, log(pmax(export_unit_value, .Machine$double.xmin)), NA_real_),
              lnE = fifelse(e > 0, log(pmax(e, .Machine$double.xmin)), NA_real_))]
  pos <- base[X > 0]
  stopifnot(max(abs(pos$lnX - pos$lnQ - pos$lnUV)) < 1e-9,
            max(abs(pos$lnV - pos$lnQ - pos$lnW)) < 1e-9)
  fwrite(base, file.path(dirs["data"], "analysis_panel.csv"), na = "NA", bom = TRUE)
  table_save("T01_sample_flow", d[, .(observations = .N, products = uniqueN(code), reported_real_value = safe_sum(X)), by = exclusion],
             "Sequential mutually exclusive exclusions; missing trade records are never changed to zero. Raw input remains intact.")
  coverage <- d[, .(universe = .N, quantity_records = sum(q_record), value_records = sum(x_record),
                    positive_quantity = sum(Q > 0, na.rm = TRUE), eligible_products = sum(eligible),
                    positive_value_eligible = sum(eligible & X > 0, na.rm = TRUE),
                    retained_value_share = sum(X[eligible], na.rm = TRUE) / sum(X[x_record & !excluded_aggregate & X >= 0], na.rm = TRUE)), by = Year]
  table_save("T02_coverage", coverage, "Coverage share uses non-overlapping study-universe export values, not all Iranian agricultural exports.")
  table_save("T03_source_flags", base[, .(observations = .N, X = sum(X), V = sum(V)), by = .(q_flag, x_flag)], "A/E/I/X are retained source quality flags, not proof that one source is true.")
  table_save("T04_source_disagreement", d[Q > 0 & Q_CWASI_implied > 0 & !excluded_aggregate,
    .(n = .N, median_relative_difference = median(quantity_disagreement), fraction_over_10pct = mean(quantity_disagreement > .1)), by = period],
    "Relative difference = abs(Q_FAO-Q_CWASI_implied)/max(Q_FAO,Q_CWASI_implied). CWASI quantity is inferred from VWT/uWFs.")
  fwrite(d[eligible == FALSE], file.path(dirs["audit"], "excluded_observations.csv"), na = "NA")
  fwrite(d[is.finite(quantity_disagreement)][order(-quantity_disagreement)], file.path(dirs["audit"], "quantity_comparison_all.csv"), na = "NA")
  fwrite(d[code %in% c(27,28,29,30,31,32)], file.path(dirs["audit"], "rice_audit.csv"), na = "NA")

  # 2. DESCRIPTION: POOLED RATIOS, NEVER UNWEIGHTED MEANS OF PRODUCT RATIOS -----
  cat("\n[2/9] Descriptive statistics, coverage, concentration and unit-value/footprint accounting.\n")
  aggregate_metrics <- function(a, bycols) {
    z <- a[, .(observations = .N, products = uniqueN(code), X = sum(X), V = sum(V), Q = sum(Q)), by = bycols]
    z[, E := X / V]; z[]
  }
  annual <- aggregate_metrics(base, "Year"); setorder(annual, Year)
  periods <- aggregate_metrics(base, "period"); setorder(periods, period)
  groups <- aggregate_metrics(base, c("period","category"))
  products <- base[, .(item = first(item), category = first(category), primary = first(primary), years = .N,
                       X = sum(X), V = sum(V), Q = sum(Q)), by = code]
  products[, `:=`(E = X / V, water_share = V / sum(V), value_share = X / sum(X))]
  setorder(products, -V)
  top3 <- head(products$code, 3); top10 <- head(products$code, 10)
  balanced <- base[, .N, by = code][N == config$end_year - config$start_year + 1L, code]
  core <- base[, .N, by = code][N >= config$min_core_years, code]
  concentration <- base[, {ww <- V/sum(V); xx <- X/sum(X);
    .(water_HHI = sum(ww^2), value_HHI = sum(xx^2), water_top10_share = sum(head(sort(ww, decreasing = TRUE),10)),
      value_top10_share = sum(head(sort(xx, decreasing = TRUE),10)))}, by = Year]
  unit_value_water <- base[, .(X = sum(X), V = sum(V), Q = sum(Q), item = first(item), category = first(category)), by = .(code,period)]
  unit_value_water[, `:=`(real_export_unit_value_per_t = X/Q, weighted_footprint_per_t = V/Q, E = X/V)]
  table_save("T05_annual", annual, "X in 2020 USD; V in m3; E=sum(X)/sum(V). Q across products is an arithmetic total, not a constant-quality output index.")
  table_save("T06_periods", periods, "Pooled period ratios; periods differ in length. Do not compare unannualized totals as growth rates.")
  table_save("T07_product_groups", groups, "Category metadata are copied from CWASI.")
  table_save("T08_product_ranking", products, "Full-period water ranking determines the pre-specified top-3/top-10 exclusion sensitivities; no p-value selection.")
  table_save("T09_concentration", concentration)
  table_save("T10_unit_value_water", unit_value_water, "E=(X/Q)/(V/Q). X/Q is a real export unit value (2020 USD/tonne), not an observed transaction price; it can reflect product quality and shipment composition. The period footprint is quantity-weighted and is not a pure technological-efficiency measure.")
  figure_save("F01_annual_value_water", ggplot(melt(annual[, .(Year, `Real exports (billion 2020 USD)` = X/1e9, `Virtual water exports (billion m3)` = V/1e9)], id.vars = "Year"), aes(Year,value))+
    geom_line(colour = "#176B87", linewidth = .8)+facet_wrap(~variable, scales = "free_y", ncol = 1)+labs(x = NULL,y = NULL), height = 6,
    caption = "Annual real export value and harmonized virtual water exports on separate axes/panels; selected coverage only.")
  figure_save("F02_value_per_water", ggplot(annual,aes(Year,E))+geom_line(colour="#176B87",linewidth=.8)+geom_point(size=1.4)+labs(x=NULL,y="2020 USD per m3"), caption="Ratio of annual sums, based on FAOSTAT quantities and CWASI supply coefficients.")
  figure_save("F03_coverage", ggplot(melt(coverage[,.(Year,quantity_records,value_records,eligible_products)],id.vars="Year"),aes(Year,value,colour=variable))+geom_line(linewidth=.75)+labs(x=NULL,y="Number of products",colour=NULL), caption="Record availability and eligible coverage are distinct from the existence of an economic trade flow.")
  figure_save("F04_concentration", ggplot(melt(concentration[,.(Year,water_HHI,value_HHI)],id.vars="Year"),aes(Year,value,colour=variable))+geom_line(linewidth=.8)+labs(x=NULL,y="Herfindahl index",colour=NULL))
  figure_save("F05_top_products", ggplot(head(products,15),aes(reorder(item,water_share),water_share*100))+geom_col(fill="#176B87")+coord_flip()+labs(x=NULL,y="Share of virtual water exports (%)"),height=6)

  # 3. EXACT BENNET DECOMPOSITION WITH SEPARATE COVERAGE ENTRY/EXIT -------------
  cat("\n[3/9] Exact Bennet accounting and coverage-sensitive decompositions.\n")
  decompose <- function(a, p0, p1, mode = "observed", timecol = "period") {
    z <- a[get(timecol) %in% c(p0,p1), .(X=sum(X),V=sum(V),item=first(item)), by=c("code",timecol)]
    z0 <- z[get(timecol)==p0,.(code,item0=item,X0=X,V0=V)]
    z1 <- z[get(timecol)==p1,.(code,item1=item,X1=X,V1=V)]
    j <- merge(z0,z1,by="code",all=TRUE)
    if (!nrow(j) || !nrow(z0) || !nrow(z1)) stop("Both comparison periods need eligible data.")
    j[, status:=fcase(is.na(V0),"coverage_entry",is.na(V1),"coverage_exit",default="common")]
    if(mode=="common_renormalized")j<-j[status=="common"]
    j[is.na(X0),`:=`(X0=0,V0=0)];j[is.na(X1),`:=`(X1=0,V1=0)]
    j[, `:=`(w0=V0/sum(V0),w1=V1/sum(V1),e0=fifelse(V0>0,X0/V0,NA_real_),e1=fifelse(V1>0,X1/V1,NA_real_))]
    j[, `:=`(within=0,composition=0,entry=0,exit=0)]
    j[status=="common",`:=`(within=.5*(w0+w1)*(e1-e0),composition=.5*(e0+e1)*(w1-w0))]
    j[status=="coverage_entry",entry:=w1*e1]
    j[status=="coverage_exit",exit:=-w0*e0]
    j[, net:=within+composition+entry+exit]
    j[, item:=fifelse(!is.na(item0),item0,item1)]
    E0<-sum(j$X0)/sum(j$V0);E1<-sum(j$X1)/sum(j$V1)
    residual<-E1-E0-sum(j$net)
    if(abs(residual)>1e-9*max(1,abs(E0),abs(E1)))stop("Bennet identity failed.")
    s<-data.table(start=as.character(p0),end=as.character(p1),mode=mode,E0=E0,E1=E1,change=E1-E0,
                  within=sum(j$within),composition=sum(j$composition),coverage_entry=sum(j$entry),coverage_exit=sum(j$exit),
                  common_products=sum(j$status=="common"),entry_products=sum(j$status=="coverage_entry"),exit_products=sum(j$status=="coverage_exit"),identity_error=residual)
    list(summary=s,detail=j)
  }
  pp<-sort(unique(base$period));pairs<-c(lapply(seq_len(length(pp)-1L),function(i)c(pp[i],pp[i+1])),list(c(pp[1],tail(pp,1))))
  decs<-list();details<-list()
  for(i in seq_along(pairs)){
    for(md in c("observed","common_renormalized")){
      zz<-decompose(base,pairs[[i]][1],pairs[[i]][2],md)
      decs[[length(decs)+1]]<-zz$summary
      details[[length(details)+1]]<-cbind(data.table(start=pairs[[i]][1],end=pairs[[i]][2],mode=md),zz$detail)
    }
  }
  dec<-rbindlist(decs);det<-rbindlist(details)
  table_save("T11_bennet",dec,"Within/composition use common products; coverage entry/exit are separate. Absence can reflect unavailable data, not true economic entry/exit. common_renormalized holds the endpoint basket fixed.")
  table_save("T12_bennet_products",det,"Exact product contributions. Undefined endpoint ratios are never arbitrarily set to zero for the main decomposition.")
  endpoint<-copy(det[start==pp[1]&end==tail(pp,1)&mode=="observed"])
  endpoint[,absnet:=abs(net)];setorder(endpoint,-absnet)
  figure_save("F06_bennet_components",ggplot(melt(dec[mode=="observed"],id.vars=c("start","end"),measure.vars=c("within","composition","coverage_entry","coverage_exit")),aes(paste(start,end,sep=" to "),value,fill=variable))+
    geom_col(position="dodge")+geom_hline(yintercept=0,linewidth=.3)+coord_flip()+labs(x=NULL,y="Contribution (2020 USD per m3)",fill=NULL),height=6)
  figure_save("F07_product_contributions",ggplot(head(endpoint,15),aes(reorder(item,net),net))+geom_col(fill="#176B87")+geom_hline(yintercept=0,linewidth=.3)+coord_flip()+labs(x=NULL,y="Net endpoint contribution (2020 USD per m3)"),height=6)
  annual_decs<-rbindlist(lapply(seq.int(config$start_year,config$end_year-1L),function(y)decompose(base,y,y+1,timecol="Year")$summary))
  table_save("T13_annual_bennet",annual_decs,"Adjacent-year accounting; component path depends on comparison horizon. Endpoint and chained components need not agree.")
  figure_save("F08_chained_bennet",ggplot(melt(annual_decs,id.vars="end",measure.vars=c("within","composition","coverage_entry","coverage_exit")),aes(as.numeric(end),value,colour=variable))+geom_line(linewidth=.65)+geom_hline(yintercept=0,linewidth=.3)+labs(x=NULL,y="Annual contribution (2020 USD per m3)",colour=NULL))

  # 4. FIXED-EFFECTS MODELS AND FINITE-CLUSTER INFERENCE ------------------------
  cat("\n[4/9] Fixed-effects models, PPML, first differences and diagnostic specifications.\n")
  # Full fixed-effect degrees counted; each cluster component gets its own G/(G-1).
  ssc <- fixest::ssc(K.adj=TRUE,K.fixef="full",G.adj=TRUE,G.df="conventional",t.df="min",K.exact=TRUE)
  fit <- function(label, fml, a, estimator="OLS", null_term="lnX", null_value=1, family="main", offset=NULL) {
    a<-copy(a)
    if(nrow(a)<30||uniqueN(a$code)<3||uniqueN(a$Year)<3)stop("Insufficient observations/products/years.")
    m<-if(estimator=="PPML")fixest::fepois(fml,data=a,offset=offset,ssc=ssc,glm.iter=200,glm.tol=1e-9,fixef.tol=1e-8,fixef.rm="perfect_fit",data.save=TRUE,warn=TRUE,notes=TRUE) else
      fixest::feols(fml,data=a,ssc=ssc,fixef.tol=1e-9,fixef.rm="none",data.save=TRUE,warn=TRUE,notes=TRUE)
    if(estimator=="PPML"&&!is.null(m$convStatus)&&!isTRUE(m$convStatus))stop("PPML did not converge; do not interpret this model.")
    model_store[[label]]<<-m
    used<-a[fixest::obs(m)]
    dfs<-c(product=uniqueN(used$code)-1L,year=uniqueN(used$Year)-1L,twoway=min(uniqueN(used$code),uniqueN(used$Year))-1L)
    for(vc in c("twoway","product","year")) {
      vv<-safe(paste(label,vc),stats::vcov(m,vcov=switch(vc,twoway=~code+Year,product=~code,year=~Year),ssc=ssc,vcov_fix=FALSE))
      if(is.null(vv))next
      b<-coef(m);ss<-sqrt(pmax(diag(vv),0));ss[diag(vv)<=0]<-NA_real_
      dd<-dfs[vc];crit<-qt(.975,dd)
      row<-data.table(model=label,outcome=all.vars(fml)[1],formula=paste(deparse(fml),collapse=" "),estimator=estimator,family=family,vcov=vc,term=names(b),estimate=unname(b),se=unname(ss),df=unname(dd),
                      lower=unname(b-crit*ss),upper=unname(b+crit*ss),p_zero=unname(2*pt(-abs(b/ss),dd)),
                      observations=nobs(m),products=uniqueN(used$code),years=uniqueN(used$Year),input_observations=nrow(a),
                      dropped_by_estimator=nrow(a)-nobs(m))
      row[,`:=`(null_value=fifelse(term==null_term,null_value,NA_real_),p_null=NA_real_)]
      row[is.finite(null_value),p_null:=2*pt(-abs((estimate-null_value)/se),df)]
      model_rows[[length(model_rows)+1L]]<<-row
    }
    m
  }
  main<-fit("M01_log_water_value",lnV~lnX|code+Year,pos)
  unrestricted<-fit("M02_quantity_unit_value",lnV~lnQ+lnUV|code+Year,pos,null_term="",family="mechanism_diagnostic")
  unit_value_model<-fit("M03_footprint_export_unit_value",lnW~lnUV|code+Year,pos,null_term="lnUV",null_value=0,family="mechanism_diagnostic")
  reverse<-fit("D_reverse_log_value_water",lnX~lnV|code+Year,pos,null_term="lnV",family="reverse_direction_diagnostic")
  diff<-copy(pos);setorder(diff,code,Year)
  diff[,`:=`(prev_year=shift(Year),dlnX=lnX-shift(lnX),dlnV=lnV-shift(lnV),dlnUV=lnUV-shift(lnUV),dlnW=lnW-shift(lnW)),by=code]
  diff<-diff[Year-prev_year==1L]
  fit("M04_first_difference",dlnV~dlnX|Year,diff,null_term="dlnX",family="robustness")
  safe("Product trends",fit("M05_product_trends",lnV~lnX|code[trend]+Year,pos,family="robustness"))
  if(config$run_ppml){
    # In the manuscript direction log(X) requires positive export value. There
    # are no observed zero-V outcomes in this positive-Q analysis sample.
    safe("PPML",fit("M06_PPML_water",V_million~lnX|code+Year,pos,estimator="PPML",family="robustness"))
    safe("PPML water intensity",fit("M07_PPML_water_intensity",V_million~lnUV|code+Year,pos,estimator="PPML",null_term="lnUV",null_value=0,offset=~lnQ,family="mechanism_diagnostic"))
    # Reverse-direction PPML can retain the recorded X=0, Q>0 observations.
    safe("Reverse PPML",fit("D_reverse_PPML_value",X_million~lnV|code+Year,base,estimator="PPML",null_term="lnV",family="reverse_direction_diagnostic"))
  }
  # Restriction implied by log(V) on log(X): coefficients of log(Q) and log(P) equal.
  ur_v<-vcov(unrestricted,vcov=~code+Year,ssc=ssc,vcov_fix=FALSE)
  delta<-coef(unrestricted)["lnQ"]-coef(unrestricted)["lnUV"]
  se_delta<-sqrt(ur_v["lnQ","lnQ"]+ur_v["lnUV","lnUV"]-2*ur_v["lnQ","lnUV"])
  table_save("T14_shared_quantity_test",data.table(restriction="beta_logQ = beta_logP",difference=unname(delta),se=unname(se_delta),df=uniqueN(pos$Year)-1,
    p_value=unname(2*pt(-abs(delta/se_delta),uniqueN(pos$Year)-1))),"Tests the restriction imposed by replacing logQ and logP with their sum logX in a logV outcome model. Association only; not a causal test.")
  cov_rows<-list()
  for(L in 1:4){
    vv<-safe(paste("Driscoll-Kraay",L),fixest::vcov_DK(main,time=~Year,lag=L,ssc=ssc,vcov_fix=FALSE))
    if(!is.null(vv)&&vv["lnX","lnX"]>0){
      b<-coef(main)["lnX"];se<-sqrt(vv["lnX","lnX"]);df<-uniqueN(pos$Year)-1
      cov_rows[[length(cov_rows)+1]]<-data.table(vcov="Driscoll-Kraay",lag=L,estimate=unname(b),se=unname(se),df=df,lower=unname(b-qt(.975,df)*se),upper=unname(b+qt(.975,df)*se),p_beta_one=unname(2*pt(-abs((b-1)/se),df)))
    }
  }
  table_save("T15_DK_inference",rbindlist(cov_rows),"Time-HAC sensitivity on the main model; T=38 limits asymptotic reliability. Main inference is two-way clustered with conservative t degrees of freedom.")

  # 5. PREDECLARED SENSITIVITIES: SAMPLE, FOOTPRINT, SOURCE, SCALE --------------
  cat("\n[5/9] Sample, coefficient, source and transformation sensitivities.\n")
  sample_sets<-list(full=base,positive_value=base[X>0],exclude_pistachio=base[code!=223],
    exclude_top3=base[!code%in%top3],exclude_top10=base[!code%in%top10],
    primary_only=base[primary==1],processed_only=base[primary==0],
    balanced=base[code%in%balanced],core_30_years=base[code%in%core],
    official_both=base[q_flag=="A"&x_flag=="A"],post_1990=base[Year>=1991],pre_2020=base[Year<=2019],
    omit_2023=base[Year<=2022],minimum_one_tonne=base[Q>=1])
  sensitivity_decs<-list()
  for(nm in names(sample_sets)){
    a<-sample_sets[[nm]]
    zz<-if(all(c(pp[1],tail(pp,1))%in%a$period))safe(paste("Endpoint",nm),decompose(a,pp[1],tail(pp,1))) else NULL
    if(!is.null(zz)){
      # The same period group can have a shorter observed window after trimming.
      # Preserve nominal groups and disclose the actual years for every sample.
      actual_window<-function(p){
        yy<-sort(unique(a[period==p,Year]))
        list(label=if(length(yy)==1L)as.character(yy) else paste0(min(yy),"-",max(yy)),
             years=paste(yy,collapse=","),n=length(yy))
      }
      w_start<-actual_window(pp[1]);w_end<-actual_window(tail(pp,1))
      zz$summary[,`:=`(nominal_start=start,nominal_end=end,
        start=w_start$label,end=w_end$label,start_years=w_start$years,
        end_years=w_end$years,start_n_years=w_start$n,end_n_years=w_end$n)]
      sensitivity_decs[[length(sensitivity_decs)+1]]<-cbind(data.table(sample=nm),zz$summary)
    }
    if(!nm%in%c("full","positive_value"))safe(paste("Sample model",nm),fit(paste0("S_",nm),lnV~lnX|code+Year,a[X>0],family="sample_robustness"))
  }
  # Fixed product-year sample for all source comparisons.
  common_source<-copy(base[is.finite(V_CWASI)&V_CWASI>0])
  old_source<-copy(common_source);old_source[,`:=`(V=V_CWASI,lnV=log(V_CWASI))]
  for(nm in c("same_sample_harmonized","same_sample_CWASI")){
    a<-if(nm=="same_sample_harmonized")common_source else old_source
    zz<-decompose(a,pp[1],tail(pp,1));sensitivity_decs[[length(sensitivity_decs)+1]]<-cbind(data.table(sample=nm),zz$summary)
    fit(paste0("S_",nm),lnV~lnX|code+Year,a[X>0],family="source_diagnostic")
  }
  # Supply/production comparison uses identical primary-product observations.
  prod_same<-copy(base[primary==1&is.finite(W_production)&W_production>0])
  prod_alt<-copy(prod_same);prod_alt[,`:=`(V=Q*W_production,lnV=log(Q*W_production))]
  for(nm in c("primary_supply_same_sample","primary_production_same_sample")){
    a<-if(nm=="primary_supply_same_sample")prod_same else prod_alt
    zz<-safe(nm,decompose(a,pp[1],tail(pp,1)))
    if(!is.null(zz))sensitivity_decs[[length(sensitivity_decs)+1]]<-cbind(data.table(sample=nm),zz$summary)
    safe(nm,fit(paste0("S_",nm),lnV~lnX|code+Year,a[X>0],family="footprint_robustness"))
  }
  # Hold each product's footprint fixed at its observed time mean. This changes
  # the estimand and is a measurement sensitivity, not a preferable estimate.
  constant<-copy(base);constant[,W_fixed:=mean(W_supply),by=code];constant[,`:=`(V=Q*W_fixed,lnV=log(Q*W_fixed))]
  sensitivity_decs[[length(sensitivity_decs)+1]]<-cbind(data.table(sample="time_constant_footprint"),decompose(constant,pp[1],tail(pp,1))$summary)
  fit("S_time_constant_footprint",lnV~lnX|code+Year,constant[X>0],family="footprint_robustness")
  # Tail removal is reported separately, never used to select the baseline.
  bounds<-quantile(pos$lnE,c(.01,.99),na.rm=TRUE,type=7)
  trimmed<-pos[lnE>=bounds[1]&lnE<=bounds[2]]
  safe("Tail sensitivity",fit("S_trim_1_99_log_ratio",lnV~lnX|code+Year,trimmed,family="sample_robustness"))
  table_save("T16_decomposition_sensitivity",rbindlist(sensitivity_decs,fill=TRUE),"Sample-sensitivity start/end labels reflect actual observed years; nominal groups and exact year lists are retained separately. omit_2023 therefore ends in 2021-2022. Same-source comparisons hold observations fixed. Production footprint applies only to primary products and assumes exported goods share domestic production coefficients. Post-1990 and pre-2020 samples have no full endpoint comparison and are used for regression sensitivity only.")
  if(config$run_ihs){
    for(sc in c(1,1e3,1e6,1e9)){
      a<-copy(base);a[,`:=`(ihsX=asinh(X/sc),ihsV=asinh(V/sc))]
      fit(paste0("IHS_scale_",format(sc,scientific=FALSE)),ihsV~ihsX|code+Year,a,null_term="",family="IHS_scale_diagnostic")
    }
  }
  # Common nominal-year deflator is absorbed by year FE in log models.
  nominal<-copy(pos);nominal[,lnX:=log(X_1000USD*1000)]
  nom<-fit("S_nominal_log_values",lnV~lnX|code+Year,nominal,family="algebraic_check")
  if(abs(coef(nom)["lnX"]-coef(main)["lnX"])>1e-7)stop("Nominal/real log slope identity failed.")
  # Endpoint-window sensitivity reduces reliance on a single chosen starting era.
  windows<-list(c(1986,1990,2021,2023),c(1986,1988,2021,2023),c(1986,1990,2019,2023),c(1991,1995,2019,2023))
  window_rows<-lapply(windows,function(w){
    a<-copy(base[(Year>=w[1]&Year<=w[2])|(Year>=w[3]&Year<=w[4])]);a[,window:=fifelse(Year<=w[2],"early","late")]
    zz<-decompose(a,"early","late",timecol="window")$summary
    cbind(data.table(early=paste(w[1:2],collapse="-"),late=paste(w[3:4],collapse="-")),zz)
  })
  table_save("T17_endpoint_windows",rbindlist(window_rows),"Alternative, explicitly reported windows; not selected to obtain a desired sign.")
  loo<-list()
  if(config$run_leave_one_product_out){
    for(cc in products$code){
      a<-base[code!=cc];b<-pos[code!=cc]
      zz<-decompose(a,pp[1],tail(pp,1))$summary
      mm<-safe(paste("Leave out",cc),fixest::feols(lnV~lnX|code+Year,data=b,fixef.rm="none",fixef.tol=1e-9,notes=FALSE))
      loo[[length(loo)+1]]<-cbind(data.table(omitted_code=cc,omitted_item=products[code==cc,item],log_slope=if(is.null(mm))NA_real_ else unname(coef(mm)["lnX"])),zz)
    }
    table_save("T18_leave_one_product_out",rbindlist(loo),"Deterministic influence analysis, not a collection of independent hypothesis tests.")
  }

  # 6. MODEL DIAGNOSTICS AND CONDITIONAL INFERENCE CHECKS ----------------------
  cat("\n[6/9] Dependence, functional form, annual stationarity and bootstrap diagnostics.\n")
  residual_data<-copy(pos[fixest::obs(main)])
  residual_data[,`:=`(residual=as.numeric(residuals(main)),fitted=as.numeric(fitted(main)))]
  setorder(residual_data,code,Year)
  residual_data[,`:=`(lag_residual=shift(residual),lag_year=shift(Year)),by=code]
  fwrite(residual_data,file.path(dirs["data"],"main_model_residuals.csv"),na="NA",bom=TRUE)
  ardata<-residual_data[Year-lag_year==1]
  arfit<-fit("D_residual_serial",residual~lag_residual|code+Year,ardata,null_term="lag_residual",null_value=0,family="diagnostic")
  # RESET-type diagnostic with clustered inference; generated fitted term means
  # this is a specification diagnostic, not a new structural coefficient.
  resetdata<-copy(residual_data);resetdata[,fitted_sq:=(fitted-mean(fitted))^2]
  fit("D_functional_form",lnV~lnX+fitted_sq|code+Year,resetdata,null_term="fitted_sq",null_value=0,family="diagnostic")
  hetero<-copy(residual_data);hetero[,residual_sq:=residual^2]
  fit("D_variance_pattern",residual_sq~lnX|code+Year,hetero,null_term="lnX",null_value=0,family="diagnostic")
  wide<-dcast(residual_data,Year~code,value.var="residual")
  rr<-as.matrix(wide[,-1]);np<-ncol(rr);cdsum<-0;correlations<-numeric()
  for(i in seq_len(np-1L))for(j in seq.int(i+1L,np)){
    ok<-is.finite(rr[,i])&is.finite(rr[,j]);nt<-sum(ok)
    if(nt>=5&&sd(rr[ok,i])>0&&sd(rr[ok,j])>0){
      rho<-cor(rr[ok,i],rr[ok,j]);cdsum<-cdsum+sqrt(nt)*rho;correlations<-c(correlations,rho)
    }
  }
  cd<-if(length(correlations))cdsum/sqrt(length(correlations)) else NA_real_
  table_save("T19_cross_section_dependence",data.table(usable_pairs=length(correlations),mean_pair_correlation=mean(correlations),mean_absolute_correlation=mean(abs(correlations)),CD_statistic=cd,normal_reference_p=2*pnorm(-abs(cd))),
    "Exploratory Pesaran-style standardized sum using pair-specific overlap (minimum 5 years). Time FE and unbalanced coverage affect the reference distribution; not a formal decision gate. Two-way/DK inference is reported regardless of the test.")
  # Exact algebra checks on the common sample: real vs nominal; within identity.
  table_save("T20_algebra_checks",data.table(check=c("log_X_equals_logP_plus_logQ","log_V_equals_logW_plus_logQ","real_nominal_log_FE_slope"),
    max_error=c(max(abs(pos$lnX-pos$lnUV-pos$lnQ)),max(abs(pos$lnV-pos$lnW-pos$lnQ)),abs(coef(main)["lnX"]-coef(nom)["lnX"]))),"These are construction checks, not empirical evidence of an economic mechanism.")
  if(config$run_wild_bootstrap){
    # Restricted wild-cluster bootstrap-t using Frisch-Waugh-Lovell projection.
    # One clustering dimension at a time, supplementary to main TWO-WAY SEs.
    dd<-as.data.frame(pos)
    Z<-model.matrix(~factor(code)+factor(Year),dd);qrZ<-qr(Z)
    xv<-qr.resid(qrZ,dd$lnX);den<-sum(xv^2);bhat<-unname(coef(main)["lnX"])
    wild_one<-function(groups,null=1,B=config$wild_bootstrap_replications){
      g<-match(groups,unique(groups));G<-max(g);k<-qrZ$rank+1L;n<-length(xv)
      adj<-G/(G-1)*(n-1)/(n-k)
      sefun<-function(u)sqrt(adj*sum(rowsum(xv*u,g,reorder=FALSE)^2))/den
      u0<-qr.resid(qrZ,dd$lnV-null*dd$lnX)
      u<-qr.resid(qrZ,dd$lnV)-bhat*xv
      t0<-(bhat-null)/sefun(u)
      ts<-rep(NA_real_,B)
      for(b in seq_len(B)){
        w<-sample(c(-1,1),G,replace=TRUE)
        ym<-qr.resid(qrZ,u0*w[g]);db<-sum(xv*ym)/den
        ub<-ym-db*xv;se<-sefun(ub)
        if(is.finite(se)&&se>0)ts[b]<-db/se
      }
      good<-is.finite(ts)
      data.table(null=null,estimate=bhat,t_stat=t0,replications=B,successful=sum(good),p_bootstrap=(1+sum(abs(ts[good])>=abs(t0)))/(1+sum(good)),clusters=G)
    }
    wr<-rbindlist(lapply(c("product","year"),function(g)cbind(data.table(cluster=g),wild_one(if(g=="product")dd$code else dd$Year))))
    table_save("T21_wild_cluster",wr,"Restricted Rademacher wild-cluster bootstrap-t for H0: beta=1. Product and year versions are separate checks; neither is a two-way wild bootstrap. Main inference remains two-way clustered.")
  }
  # Time-series supplement: log levels with a trend and consecutive log changes.
  ann<-copy(annual);ann[,`:=`(lnX=log(X),lnV=log(V),lnE=log(E),trend=Year-config$start_year)]
  ann[,`:=`(dlnX=lnX-shift(lnX),dlnV=lnV-shift(lnV),dlnE=lnE-shift(lnE))]
  ann[,lag_dlnV:=shift(dlnV)]
  ann_models<-list(levels_with_trend=lm(lnV~lnX+trend,data=ann),first_differences=lm(dlnV~dlnX,data=ann),
                   dynamic_growth=lm(dlnV~dlnX+lag_dlnV,data=ann))
  ann_rows<-list()
  for(nm in names(ann_models))for(L in c(1,2,3,4)){
    m<-ann_models[[nm]];vv<-sandwich::NeweyWest(m,lag=L,prewhite=FALSE,adjust=TRUE);b<-coef(m);se<-sqrt(diag(vv));df<-df.residual(m)
    ann_rows[[length(ann_rows)+1]]<-data.table(model=nm,lag=L,term=names(b),estimate=unname(b),se=unname(se),lower=unname(b-qt(.975,df)*se),upper=unname(b+qt(.975,df)*se),p_zero=unname(2*pt(-abs(b/se),df)),observations=nobs(m))
  }
  table_save("T22_annual_HAC",rbindlist(ann_rows),"Small-T descriptive supplement. Changing product coverage and potential nonstationarity limit levels interpretation; no cointegration or causal claim is inferred.")
  if(config$run_stationarity_diagnostics){
    # ADF statistic with intercept and deterministic trend, fixed lag order.
    # Reference p-values are simulated under an IID Gaussian random-walk null,
    # NOT the ordinary OLS t distribution. Report lag sensitivity and assumptions.
    adf_t<-function(y,L){
      dy<-diff(y);idx<-seq.int(L+1L,length(dy))
      A<-cbind(1,idx,y[idx])
      if(L>0)for(j in seq_len(L))A<-cbind(A,dy[idx-j])
      fit0<-lm.fit(A,dy[idx]);n<-length(idx);k<-fit0$rank
      if(k<ncol(A))return(NA_real_)
      se<-sqrt(sum(fit0$residuals^2)/(n-k)*diag(chol2inv(chol(crossprod(A)))))
      unname(fit0$coefficients[3]/se[3])
    }
    sr<-list();B<-max(999L,config$bootstrap_replications)
    for(L in 0:2){
      sim<-replicate(B,adf_t(cumsum(rnorm(nrow(ann))),L));sim<-sim[is.finite(sim)]
      for(vv in c("lnX","lnV","lnE")){
        tt<-adf_t(ann[[vv]],L)
        sr[[length(sr)+1]]<-data.table(series=vv,ADF_lags=L,statistic=tt,critical_5pct=unname(quantile(sim,.05)),p_simulated=(1+sum(sim<=tt))/(length(sim)+1),replications=length(sim))
      }
    }
    table_save("T23_annual_ADF",rbindlist(sr),"ADF includes intercept+trend, lag orders 0/1/2. Monte Carlo reference assumes IID Gaussian innovations; heteroskedasticity/dependence and T=38 limit reliability. Failure to reject is not proof of a unit root. This is not a panel cointegration test.")
  }

  # 7. DECOMPOSITION RESAMPLING AND FOOTPRINT SCENARIOS ------------------------
  cat("\n[7/9] Conditional temporal resampling and footprint scenarios.\n")
  if(config$run_bootstrap){
    draw_blocks<-function(yy,L){
      n<-length(yy);starts<-sample(seq_len(n),ceiling(n/L),replace=TRUE)
      ii<-unlist(lapply(starts,function(s)((s-1L+seq.int(0,L-1L))%%n)+1L))
      yy[ii[seq_len(n)]]
    }
    # Resample years jointly across all products within each endpoint period.
    # This preserves contemporaneous comovement; very short endpoint windows
    # make these conditional sensitivity intervals, not source-error uncertainty.
    bp<-base[period%in%c(pp[1],tail(pp,1))]
    br<-vector("list",config$bootstrap_replications)
    for(b in seq_len(config$bootstrap_replications)){
      samples<-lapply(c(pp[1],tail(pp,1)),function(pr){
        a<-bp[period==pr];yy<-sort(unique(a$Year));sel<-draw_blocks(yy,min(config$block_length,max(1L,length(yy)-1L)))
        w<-data.table(Year=sel)[,.N,by=Year];aa<-merge(a,w,by="Year",all=FALSE)
        aa[,`:=`(X=X*N,V=V*N)];aa
      })
      zz<-decompose(rbindlist(samples),pp[1],tail(pp,1))$summary
      br[[b]]<-zz[,.(change,within,composition,coverage_entry,coverage_exit)]
    }
    boot<-rbindlist(br);boot[,replication:=.I]
    table_save("T24_bennet_bootstrap_draws",boot,"Temporal resampling sensitivity conditional on supplied source values; not a calibrated measure of water-footprint estimation uncertainty.")
    bs<-rbindlist(lapply(setdiff(names(boot),"replication"),function(nm)data.table(component=nm,median=median(boot[[nm]]),lower=unname(quantile(boot[[nm]],.025)),upper=unname(quantile(boot[[nm]],.975)),replications=nrow(boot))))
    table_save("T25_bennet_resampling",bs,"95% percentile resampling intervals; circular year blocks within each endpoint period, shared across products. Block length is min(configured length, period length minus one). The last period has only three years; interpret cautiously.")
  }
  ep<-decompose(base,pp[1],tail(pp,1))$summary
  scen<-rbindlist(lapply(config$coefficient_sensitivity,function(u)data.table(relative_footprint_deviation=u,
    baseline_change=ep$change,lower_change=ep$E1/(1+u)-ep$E0/(1-u),upper_change=ep$E1/(1-u)-ep$E0/(1+u))))
  table_save("T26_footprint_scenarios",scen,"Deterministic bounds if each endpoint aggregate water estimate can vary independently by the stated percentage. These are assumptions, not source-provided uncertainty intervals.")
  # Unit value versus inverse-footprint Shapley/Bennet accounting at product level.
  pw0<-unit_value_water[period==pp[1],.(code,UV0=real_export_unit_value_per_t,W0=weighted_footprint_per_t,E0=E)]
  pw1<-unit_value_water[period==tail(pp,1),.(code,UV1=real_export_unit_value_per_t,W1=weighted_footprint_per_t,E1=E)]
  pw<-merge(pw0,pw1,by="code")
  pw[,`:=`(unit_value_contribution=.5*(1/W0+1/W1)*(UV1-UV0),footprint_contribution=.5*(UV0+UV1)*(1/W1-1/W0))]
  pw[,identity_error:=E1-E0-unit_value_contribution-footprint_contribution]
  stopifnot(max(abs(pw$identity_error))<1e-8)
  pw<-merge(pw,items[,.(code,item,category)],by="code")
  table_save("T27_unit_value_footprint_split",pw,"Exact decomposition of endpoint product ratio UV/W into unit-value and inverse-footprint contributions. Common products only; no causal interpretation.")

  # Aggregate using the ORIGINAL full-basket mean water weights. Do not
  # renormalize common products: that would target a different within term.
  within_reference<-decompose(base,pp[1],tail(pp,1),mode="observed")
  common_weights<-within_reference$detail[status=="common",.(code,w0,w1,within)]
  stopifnot(setequal(pw$code,common_weights$code),!anyDuplicated(pw$code))
  weighted_split<-merge(pw,common_weights,by="code",all=FALSE)
  weighted_split[,mean_water_share:=.5*(w0+w1)]
  weighted_split[,`:=`(
    within_unit_value=mean_water_share*unit_value_contribution,
    within_footprint=mean_water_share*footprint_contribution)]
  weighted_split[,within_identity_error:=within-within_unit_value-within_footprint]
  tol<-1e-9*max(1,abs(within_reference$summary$within))
  if(any(!is.finite(weighted_split$within_identity_error))||
     any(abs(weighted_split$within_identity_error)>tol))stop("Product weighted-within identity failed.")
  within_split<-weighted_split[,.(start=pp[1],end=tail(pp,1),mode="observed",
    common_products=.N,within=sum(within),unit_value_effect=sum(within_unit_value),
    footprint_effect=sum(within_footprint))]
  within_split[,`:=`(component_sum=unit_value_effect+footprint_effect,
    reference_within=within_reference$summary$within)]
  within_split[,`:=`(identity_error=within-component_sum,
    reference_error=within-reference_within)]
  if(any(!is.finite(unlist(within_split[,.(within,component_sum,identity_error,reference_error)])))||
     abs(within_split$identity_error)>tol||abs(within_split$reference_error)>tol)
    stop("Aggregate within split does not reconcile with observed Bennet within.")
  table_save("T37_within_split_products",weighted_split,"Common-product unit-value and inverse-footprint contributions multiplied by (w0+w1)/2 from full-basket water shares. Unit: 2020 USD/m3. Weights are not renormalized to the common basket; coverage contributions remain separate.")
  table_save("T38_within_split_aggregate",within_split,"Exact endpoint identity: observed Bennet within = unit_value_effect + footprint_effect. Symmetric two-factor decomposition of UV*(1/W), aggregated with original Bennet mean water shares. Both identity_error and reference_error are checked. Accounting attribution, not causal effects or pure technological change.")

  # 8. ALL MODEL TABLES, FIGURES AND MACHINE-READABLE RESULTS -------------------
  cat("\n[8/9] Exporting complete tables, figures and reproducibility records.\n")
  models<-rbindlist(model_rows,fill=TRUE)
  models[,p_null_Holm:=NA_real_]
  models[vcov=="twoway"&family=="sample_robustness"&is.finite(p_null),p_null_Holm:=p.adjust(p_null,method="holm")]
  table_save("T28_all_model_coefficients",models,"All coefficients are exported. Two-way clustering is primary; product/year alternatives are supplementary. Holm adjustment applies to the family of sample-sensitivity H0: beta=1 tests. IHS coefficients are not elasticities and receive no beta=1 test.")
  main_table<-models[vcov=="twoway"&grepl("^M0",model),
    .(model,outcome,estimator,term,estimate,se,lower,upper,p_zero,p_null,observations,products,years)]
  table_save("T29_main_models",main_table,"Primary direction matches the manuscript: water outcome, export-value regressor. Log regressors require positive values, also in the main-direction PPML. Reverse PPML in T28 includes recorded zero export values with positive Q. No model identifies an extensive margin from absent records.")
  figure_save("F09_model_coefficients",ggplot(models[vcov=="twoway"&term%in%c("lnX","dlnX")&!family%in%c("diagnostic","algebraic_check")],aes(estimate,reorder(model,estimate)))+
    geom_vline(xintercept=1,linetype=2,colour="#777777")+geom_errorbar(aes(xmin=lower,xmax=upper),orientation="y",width=.15)+geom_point(colour="#176B87",size=2)+labs(x="Export-value slope in water model, with 95% interval",y=NULL),height=9,width=10,
    caption="Water is the dependent variable. Two-way clustered intervals; OLS, differences, PPML and sample sensitivities have distinct estimands. A slope near one can reflect shared export quantity. Source data: T28.")
  if(config$run_ihs){
    ihs_plot<-copy(models[vcov=="twoway"&family=="IHS_scale_diagnostic"])
    ihs_plot[,scale:=factor(sub("IHS_scale_","",model),levels=c("1","1000","1000000","1000000000"),labels=c("1","1,000","1,000,000","1,000,000,000"))]
    figure_save("F10_IHS_units",ggplot(ihs_plot,aes(scale,estimate))+geom_errorbar(aes(ymin=lower,ymax=upper),width=.15)+geom_point(size=2,colour="#176B87")+labs(x="Divide both variables by",y="IHS coefficient (not an elasticity)"))
  }
  figure_save("F11_source_quantities",ggplot(d[Q>0&Q_CWASI_implied>0&!excluded_aggregate],aes(Q,Q_CWASI_implied))+geom_point(alpha=.35,size=.8,colour="#176B87")+geom_abline(slope=1,intercept=0,linetype=2)+scale_x_log10()+scale_y_log10()+labs(x="FAOSTAT export quantity (tonnes, log scale)",y="CWASI-implied quantity (tonnes, log scale)"))
  sp<-rbind(aggregate_metrics(common_source,"Year")[,source:="Harmonized water"],aggregate_metrics(old_source,"Year")[,source:="Original CWASI water"])
  table_save("T30_source_annual_same_sample",sp,"Annual source comparison on identical observations and export values.")
  figure_save("F12_source_ratio_comparison",ggplot(sp,aes(Year,E,colour=source))+geom_line(linewidth=.8)+labs(x=NULL,y="2020 USD per m3",colour=NULL))
  figure_save("F13_residual_diagnostics",ggplot(residual_data,aes(fitted,residual))+geom_point(alpha=.25,size=.7,colour="#176B87")+geom_hline(yintercept=0,linetype=2)+labs(x="Fitted log virtual water exports",y="Log-model residual"))
  # Balanced-basket trend is a transparent composition control, not predicted FE.
  if(length(balanced)){
    fixed<-aggregate_metrics(base[code%in%balanced],"Year");fixed[,basket:="Balanced products"]
    varying<-copy(annual);varying[,basket:="Available eligible products"]
    ff<-rbind(varying,fixed,fill=TRUE);table_save("T31_fixed_basket_trend",ff)
    figure_save("F14_fixed_basket",ggplot(ff,aes(Year,E,colour=basket))+geom_line(linewidth=.8)+labs(x=NULL,y="2020 USD per m3",colour=NULL))
  }
  if(config$run_leave_one_product_out){
    lr<-rbindlist(loo);table_save("T32_influence_plot_data",lr[,.(omitted_code,omitted_item,change,log_slope)])
    figure_save("F15_product_influence",ggplot(lr,aes(log_slope,change))+geom_point(colour="#176B87",alpha=.65)+geom_hline(yintercept=ep$change,linetype=2)+geom_vline(xintercept=coef(main)["lnX"],linetype=2)+labs(x="Export-value slope in water model after omission",y="Endpoint ratio change after omission"))
  }
  pist<-unit_value_water[code==223]
  if(nrow(pist))figure_save("F16_pistachio_unit_value_footprint",ggplot(melt(pist[,.(period,`Real export unit value (2020 USD per tonne)`=real_export_unit_value_per_t,`Supply footprint (m3 per tonne)`=weighted_footprint_per_t)],id.vars="period"),aes(period,value,group=1))+
    geom_line(colour="#176B87",linewidth=.7)+geom_point(size=2)+facet_wrap(~variable,scales="free_y",ncol=1)+labs(x=NULL,y=NULL),height=6)
  # Requested additional figure: pooled full-period product-group shares.
  group_shares<-base[,.(X=sum(X),V=sum(V),observations=.N,products=uniqueN(code)),by=category]
  group_shares[,`:=`(water_share_pct=100*V/sum(V),real_export_share_pct=100*X/sum(X))]
  setorder(group_shares,-real_export_share_pct,category)
  stopifnot(abs(sum(group_shares$water_share_pct)-100)<1e-9,
            abs(sum(group_shares$real_export_share_pct)-100)<1e-9)
  table_save("T36_group_shares_full_period",group_shares,"1986-2023 pooled shares: sums over the eligible study basket, not averages of annual shares or shares of all Iranian exports. Categories follow CWASI metadata.")
  gp<-melt(group_shares,id.vars="category",measure.vars=c("water_share_pct","real_export_share_pct"))
  gp[,category:=factor(category,levels=rev(group_shares$category))]
  gp[,variable:=factor(variable,levels=c("water_share_pct","real_export_share_pct"),labels=c("Virtual water exports","Real export value"))]
  figure_save("F17_product_group_shares",ggplot(gp,aes(category,value,fill=variable))+
    geom_col(position=position_dodge(width=.75),width=.65)+coord_flip()+
    scale_fill_manual(values=c("Virtual water exports"="#176B87","Real export value"="#D3A250"))+
    labs(x=NULL,y="Share of study-basket total (%)",fill=NULL),width=8,height=4.5,
    caption="Full-period shares of six CWASI product groups in virtual water exports and real export value, 1986-2023. Each measure sums to 100% over the eligible study basket. Source data: T36.")
  chart_note<-data.frame(Field=c("Period","Coverage","Prices","CPI source","Axes"),
    Description=c("1986-2023","Eligible study basket; same sample as the analysis panel.",
    "Annual FAOSTAT 1000 USD * 1000 * CPI2020/CPIyear, then summed.",cpi_provenance,
    "Linear axes. Group chart: grouped horizontal bars."))
  writexl::write_xlsx(list(Chart_data=as.data.frame(group_shares[,.(Group=category,Water_share_pct=water_share_pct,Real_export_share_pct=real_export_share_pct)]),
    Source_totals=as.data.frame(group_shares),Notes=chart_note),file.path(out,"Figure_17_Product_Group_Shares.xlsx"))
  table_save("T33_figure_captions",figure_notes)
  table_save("T34_warnings",warnings_log,"Every non-estimable optional specification and intercepted warning is recorded; absence of an optional result is not interpreted as support for a claim.")
  fit_rows<-rbindlist(lapply(names(model_store),function(nm){
    mm<-model_store[[nm]]
    stat<-function(type)tryCatch(as.numeric(fixest::fitstat(mm,type,simplify=TRUE)),error=function(e)NA_real_)
    data.table(model=nm,observations=nobs(mm),parameters_including_FE=mm$nparams,
               R2=if(mm$method=="feols")stat("r2") else NA_real_,
               within_R2=if(mm$method=="feols")stat("wr2") else NA_real_,
               pseudo_R2=if(mm$method!="feols")stat("pr2") else NA_real_,
               AIC=AIC(mm),BIC=BIC(mm))
  }),fill=TRUE)
  table_save("T35_model_fit",fit_rows,"R2 and within-R2 are for OLS; pseudo-R2 is for PPML. Information criteria across different outcomes, transformations or samples are not comparable. Coefficient tables identify the outcome, formula, sample and estimator.")

  # Self-contained HTML and LaTeX tables; no Word/TeX installation required.
  escape_html<-function(x){x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);gsub(">","&gt;",x,fixed=TRUE)}
  fmt<-function(x){if(is.numeric(x))ifelse(is.na(x),"NA",sprintf("%.6g",x)) else ifelse(is.na(x),"NA",as.character(x))}
  html_table<-function(x){
    heads<-paste0("<th>",escape_html(names(x)),"</th>",collapse="")
    cols<-lapply(x,fmt)
    rows<-if(nrow(x))vapply(seq_len(nrow(x)),function(i)paste0("<tr>",paste0("<td>",escape_html(vapply(cols,`[`,character(1),i)),"</td>",collapse=""),"</tr>"),character(1)) else ""
    paste0("<table><thead><tr>",heads,"</tr></thead><tbody>",paste(rows,collapse="\n"),"</tbody></table>")
  }
  latex_escape<-function(x){
    x<-gsub("\\","\\textbackslash{}",x,fixed=TRUE)
    for(ch in c("&","%","$","#","_"))x<-gsub(ch,paste0("\\",ch),x,fixed=TRUE)
    x
  }
  index_sections<-character()
  for(nm in names(tables)){
    tb<-tables[[nm]];note<-table_notes[[nm]]
    h<-paste0("<!doctype html><html><head><meta charset='utf-8'><style>body{font-family:Arial;margin:30px;color:#16324f}table{border-collapse:collapse;font-size:12px}th,td{padding:7px 10px;text-align:right;border-bottom:1px solid #ddd}th{border-top:2px solid #16324f;border-bottom:2px solid #16324f}p{max-width:1100px}</style></head><body><h2>",nm,"</h2><p>",escape_html(note),"</p>",html_table(tb),"</body></html>")
    writeLines(h,file.path(dirs["tables"],paste0(nm,".html")),useBytes=TRUE)
    # Large audit tables remain CSV/XLSX/HTML; compact tables also have booktabs.
    if(nrow(tb)<=300){
      vals<-lapply(tb,fmt)
      tx<-c("% Requires \\usepackage{booktabs}; large-width tables may need landscape.",paste0("% ",note),paste0("\\begin{tabular}{",paste(rep("r",ncol(tb)),collapse=""),"}"),"\\toprule",
            paste0(paste(latex_escape(names(tb)),collapse=" & ")," \\\\"),"\\midrule")
      if(nrow(tb))tx<-c(tx,vapply(seq_len(nrow(tb)),function(i)paste0(paste(latex_escape(vapply(vals,`[`,character(1),i)),collapse=" & ")," \\\\"),character(1)))
      writeLines(c(tx,"\\bottomrule","\\end{tabular}"),file.path(dirs["tables"],paste0(nm,".tex")),useBytes=TRUE)
    }
    index_sections<-c(index_sections,paste0("<li><a href='tables/",nm,".html'>",nm,"</a>: ",escape_html(note),"</li>"))
  }
  # Excel results workbook is a separate OUTPUT, never the raw workbook.
  xlsx_tables<-tables
  names(xlsx_tables)<-substr(names(xlsx_tables),1,31)
  names(xlsx_tables)<-make.unique(names(xlsx_tables),sep="_")
  writexl::write_xlsx(xlsx_tables,file.path(out,"All_Results.xlsx"))
  saveRDS(model_store,file.path(dirs["models"],"estimated_models.rds"))
  saveRDS(list(full_panel=d,analysis_panel=base,annual=annual,periods=periods),file.path(dirs["data"],"analysis_data.rds"))
  package_versions<-data.table(package=needed,version=vapply(needed,function(p)as.character(packageVersion(p)),character(1)))
  fwrite(package_versions,file.path(dirs["audit"],"package_versions.csv"))
  capture.output(sessionInfo(),file=file.path(dirs["audit"],"sessionInfo.txt"))
  citations<-c(
    "CWASI data: https://zenodo.org/records/18964006",
    "Tamea et al. (2021). Virtual water trade and water footprint of agricultural goods. https://doi.org/10.5194/essd-13-2025-2021",
    "FAOSTAT Crops and livestock products trade: https://www.fao.org/faostat/en/#data/TCL",
    paste(cpi_provenance,"https://data.worldbank.org/indicator/FP.CPI.TOTL?locations=US"),
    "fixest estimation and inference: https://lrberge.github.io/fixest/",
    "fixest small-sample correction: https://lrberge.github.io/fixest/reference/ssc.html",
    "fixest time-HAC estimators: https://lrberge.github.io/fixest/reference/vcov_hac.html",
    "Chen and Roth (2024), Logs with Zeros? Some Problems and Solutions. QJE 139(2), 891-936. https://www.jonathandroth.com/assets/files/LogUniqueHOD0_Draft_Accepted.pdf",
    "Cameron, Gelbach and Miller (2011), Robust inference with multiway clustering. https://doi.org/10.1198/jbes.2010.07136",
    "Santos Silva and Tenreyro (2006), The Log of Gravity. https://doi.org/10.1162/rest.88.4.641",
    "Driscoll and Kraay (1998), Consistent covariance matrix estimation with spatially dependent panel data. https://doi.org/10.1162/003465398557825"
  )
  writeLines(citations,file.path(out,"References_and_sources.txt"),useBytes=TRUE)
  mainrow<-models[model=="M01_log_water_value"&vcov=="twoway"&term=="lnX"]
  report<-c(
    "IRAN VIRTUAL WATER: ANALYSIS AND INTERPRETATION NOTES",
    paste("Input:",basename(input)),paste("Eligible observations:",nrow(base),"; products:",uniqueN(base$code)),
    paste("Log-positive observations:",nrow(pos),"; products:",uniqueN(pos$code)),
    sprintf("Endpoint E: %.6f to %.6f; change %.6f (2020 USD/m3).",ep$E0,ep$E1,ep$change),
    sprintf("Log-water-on-log-value FE slope: %.6f; two-way SE %.6f; 95%% interval [%.6f, %.6f]; p(H0 beta=1)=%.6f.",mainrow$estimate,mainrow$se,mainrow$lower,mainrow$upper,mainrow$p_null),
    "",
    "Measurement: V=FAOSTAT export quantity * CWASI supply footprint, X=FAOSTAT export value deflated using CPI. Sources are not silently mixed at different trade quantities.",
    "Raw is relative to the study. CWASI inputs are already modeled estimates. CPI retains its published 2010=100 base in the input; conversion to constant 2020 USD occurs only in this script.",
    cpi_provenance,
    sprintf("Within split (%s to %s): unit-value %.9f + footprint %.9f = within %.9f; identity error %.3g.",within_split$start,within_split$end,within_split$unit_value_effect,within_split$footprint_effect,within_split$within,within_split$identity_error),
    "X/Q is real export unit value, not an observed transaction price. Its changes can reflect prices, quality and within-code shipment composition. The footprint component is an accounting contribution and does not identify technological or causal water savings.",
    "Rice aggregate 30 is excluded; component products remain individually identified. Coefficients and raw records outside the analysis sample remain in the workbook.",
    "No missing FAOSTAT record is converted to zero. Coverage entry/exit is distinct from actual trade entry/exit; the endpoint intersection sensitivity renormalizes the common basket.",
    "Supply water includes green and blue water and can originate outside Iran. No claim about domestic groundwater depletion, local water scarcity, profit, or welfare follows from this ratio.",
    "Primary regression: log V on log X, matching the manuscript question. The reverse log X on log V is separately labelled and is not its reciprocal. Both contain log Q, so neither identifies causal water productivity. Consult the quantity/unit-value restriction and footprint-on-unit-value models.",
    "Two-way clustered t inference is primary; year clusters number only 38. Product/year-only and Driscoll-Kraay checks do not eliminate all finite-sample uncertainty.",
    "The one-dimensional wild bootstraps are supplementary and do not replicate two-way wild inference. IHS coefficients depend on units and are not elasticities.",
    "Main-direction PPML models water using log export value and requires positive values. Reverse-direction PPML models export value and retains observed X=0 with positive Q/W. IHS also retains these zeros. None includes missing records or fabricated non-trade values. Check estimator removals and convergence warnings.",
    "No IV, sanctions causal effect, dynamic panel GMM, or panel cointegration claim is made: the supplied variables do not provide a justified identification strategy for these estimands.",
    "Annual levels, first differences, dynamic growth and ADF simulations are supplementary. Varying coverage, few years and assumptions on innovations must be acknowledged. Earlier macro controls are not carried forward from undocumented derived series; they require independently documented source data.",
    "Resampling intervals are conditional temporal sensitivities, not uncertainty in source footprints. +/-10/20% footprint scenarios are assumed bounds, not measured error distributions.",
    "The shared annual CPI deflator is absorbed by year FE in log models; nominal/real slope agreement is an algebraic check, not independent empirical robustness.",
    "All model/sample changes are disclosed. The specification set is fixed in this script but is not a preregistration. Do not choose a preferred specification because a coefficient is significant or agrees with an earlier narrative.",
    "Figures are white-background PDF vectors and 400-dpi PNGs, with external captions and underlying tabular data. Tables are CSV, HTML, compact LaTeX booktabs and a combined Excel results workbook.",
    paste("Recorded optional-stage warnings:",nrow(warnings_log)),
    "Inspect T34_warnings and run_log.txt before using results. Automatic execution and diagnostic testing cannot guarantee journal acceptance."
  )
  writeLines(report,file.path(out,"Analysis_Notes.txt"),useBytes=TRUE)
  fig_sections<-vapply(seq_len(nrow(figure_notes)),function(i)paste0("<h3>",figure_notes$file[i],"</h3><img style='max-width:100%' src='figures/",figure_notes$file[i],".png'><p>",escape_html(figure_notes$caption[i]),"</p>"),character(1))
  writeLines(paste0("<!doctype html><html><head><meta charset='utf-8'><title>Iran virtual-water analysis</title><style>body{font:15px Arial;margin:35px;max-width:1200px;color:#16324f}li{margin:9px 0}pre{white-space:pre-wrap}</style></head><body><h1>Iran virtual-water analysis</h1><pre>",escape_html(paste(report,collapse="\n")),"</pre><h2>Tables</h2><ul>",paste(index_sections,collapse="\n"),"</ul><h2>Figures</h2>",paste(fig_sections,collapse="\n"),"</body></html>"),file.path(out,"Results_Report.html"),useBytes=TRUE)
  files<-setdiff(list.files(out,recursive=TRUE,full.names=TRUE),file.path(out,"run_log.txt"))
  if(any(file.info(files)$size==0L))stop("One or more output files are empty; inspect the output directory before using results.")
  png_files<-list.files(dirs["figures"],pattern="\\.png$",full.names=TRUE)
  if(!all(vapply(png_files,valid_png,logical(1))))stop("A PNG is incomplete at final validation; do not use the incomplete output.")
  pdf_files<-list.files(dirs["figures"],pattern="\\.pdf$",full.names=TRUE)
  if(!all(vapply(pdf_files,valid_pdf,logical(1))))stop("A PDF is incomplete at final validation; do not use the incomplete output.")
  inventory<-data.table(file=substring(files,nchar(out)+2L),bytes=file.info(files)$size,MD5=unname(tools::md5sum(files)))
  fwrite(inventory,file.path(out,"Output_inventory.csv"))
  cat("\n[9/9] COMPLETED.\n",paste(report[1:6],collapse="\n"),"\nOpen Results_Report.html or All_Results.xlsx in:\n",out,"\n")
  invisible(list(output_directory=out,main_model=main,main_table=main_table,endpoint=ep,warnings=warnings_log))
}

# Set options(vwt.no_autorun=TRUE) before sourcing only for code inspection/tests.
if (!isTRUE(getOption("vwt.no_autorun", FALSE))) VWT_RESULT <- vwt_run()
