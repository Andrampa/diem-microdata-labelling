# ============================================================
# DIEM microdata labelling script (R version)
#   
#   This script takes a DIEM household survey microdata file (CSV or Excel) and produces
#   a labelled version of the dataset by automatically applying value labels from the
#   official DIEM codebook.
#
# - Detects DIEM microdata infrastructure (v1/v2) from header
# - Downloads the correct DIEM codebook workbook if missing
# - Applies value labels using one sheet per field (code/label columns)
# - Handles derived yes/no fields listed in "derived_fields" sheet
# - Produces diagnostics: candidate fields, skipped reasons,
#   low match-rate counts, overall match rate, top unmapped codes per field
# - Saves output next to input with "_LABELLED" appended
#
# Supported microdata input: .csv, .xlsx, .xls
# Supported output:
#   - .csv -> .csv
#   - .xlsx -> .xlsx
#   - .xls -> .xlsx (R does not reliably write .xls; this avoids corrupt output)
# ============================================================

# ---------------------------
# Packages (auto-install)
# ---------------------------
needed_pkgs <- c("readr", "readxl", "writexl", "stringr")
to_install <- needed_pkgs[!vapply(needed_pkgs, requireNamespace, FUN.VALUE = logical(1), quietly = TRUE)]
if (length(to_install) > 0) {
  install.packages(to_install, repos = "https://cloud.r-project.org")
}

library(readr)
library(readxl)
library(writexl)
library(stringr)

# ------------------------------------------------------------
# User inputs start
# ------------------------------------------------------------
# Select or paste the microdata file path (Windows / macOS / Linux safe)

# NOTE:
# Do not hard-code file paths with backslashes (\) inside quotes.
# Paste the path or use the file picker below.

MICRODATA_FILE <- readline("Paste full path to microdata file (or press Enter to browse): ")

if (MICRODATA_FILE == "") {
  MICRODATA_FILE <- file.choose()
}

# Normalize path:
# - converts backslashes to forward slashes
# - works across Windows / macOS / Linux
MICRODATA_FILE <- normalizePath(
  gsub("\\\\", "/", MICRODATA_FILE),
  winslash = "/",
  mustWork = FALSE
)
# Keep same defaults as Python
ADD_LABEL_COLUMNS <- FALSE  # if FALSE, original coded value is REPLACED by labelled value
LABEL_SUFFIX <- "_label"
DERIVED_FIELDS_SHEET <- "derived_fields"
MIN_OK_MATCH_RATE <- 0.95
# ------------------------------------------------------------
# User inputs end
# ------------------------------------------------------------

# Dictionary download URLs (ArcGIS items)
DICT_URL_1 <- "https://hqfao.maps.arcgis.com/sharing/rest/content/items/e59d08ded7c1440587493bf65236cf44/data"
DICT_URL_2 <- "https://hqfao.maps.arcgis.com/sharing/rest/content/items/41fa55934d2f462f86cd381ee8dc1fda/data"

# ---------------------------
# Helpers
# ---------------------------

get_ext <- function(path) {
  ext <- tolower(tools::file_ext(path))
  if (nzchar(ext)) paste0(".", ext) else ""
}

# Vectorized equivalent of Python _normalize_code()
# - trims
# - empty -> NA
# - if numeric: integer-like -> "int", else -> "float as character"
normalize_code <- function(x) {
  s <- as.character(x)
  s <- str_trim(s)
  s[s == "" | is.na(s)] <- NA_character_

  suppressWarnings({
    num <- as.numeric(s)
    ok_num <- !is.na(num) & !is.na(s)

    if (any(ok_num)) {
      # Treat as integer if "close enough" (safer than exact equality)
      is_int <- abs(num - round(num)) < 1e-9
      idx_int <- ok_num & is_int
      idx_float <- ok_num & !is_int

      if (any(idx_int)) s[idx_int] <- as.character(as.integer(round(num[idx_int])))
      if (any(idx_float)) s[idx_float] <- as.character(num[idx_float])
    }
  })

  s
}

# Read only columns (header) to detect infrastructure
read_header_columns <- function(input_path) {
  ext <- get_ext(input_path)
  if (ext == ".csv") {
    df0 <- readr::read_csv(input_path, n_max = 0, show_col_types = FALSE)
    return(names(df0))
  }
  if (ext %in% c(".xlsx", ".xls")) {
    df0 <- readxl::read_excel(input_path, n_max = 0, .name_repair = "minimal")
    return(names(df0))
  }
  stop(sprintf("Unsupported microdata extension '%s'. Use .xlsx, .xls, or .csv.", ext))
}

# Read microdata like Python:
# - CSV: all columns as character
# - Excel: read all as text to avoid numeric coercion issues
read_microdata_any <- function(input_path) {
  ext <- get_ext(input_path)

  if (ext == ".csv") {
    return(readr::read_csv(
      input_path,
      col_types = readr::cols(.default = readr::col_character()),
      na = c("", "NA", "NaN", "nan"),
      show_col_types = FALSE
    ))
  }

  if (ext %in% c(".xlsx", ".xls")) {
    return(readxl::read_excel(input_path, col_types = "text", .name_repair = "minimal"))
  }

  stop(sprintf("Unsupported microdata extension '%s'. Use .xlsx, .xls, or .csv.", ext))
}

# Build mapping from a dictionary sheet (expects columns "code" and "label" case-insensitive)
build_mapping_from_sheet <- function(df_sheet) {
  if (is.null(df_sheet) || nrow(df_sheet) == 0) return(NULL)

  cols_lower <- tolower(str_trim(names(df_sheet)))
  code_idx <- which(cols_lower == "code")
  label_idx <- which(cols_lower == "label")

  if (length(code_idx) != 1 || length(label_idx) != 1) return(NULL)

  codes <- normalize_code(df_sheet[[code_idx]])
  labels <- str_trim(as.character(df_sheet[[label_idx]]))

  keep <- !is.na(codes)
  codes <- codes[keep]
  labels <- labels[keep]

  if (length(codes) == 0) return(NULL)

  # Named vector: names are codes, values are labels
  setNames(labels, codes)
}

# Read derived fields (field names are in the SECOND column)
read_derived_fields <- function(path_xlsx, sheet_name) {
  df <- tryCatch(readxl::read_excel(path_xlsx, sheet = sheet_name, .name_repair = "minimal"),
                 error = function(e) NULL)
  if (is.null(df) || nrow(df) == 0 || ncol(df) < 2) return(character(0))

  v <- as.character(df[[2]])
  v <- str_trim(v)
  v <- v[!is.na(v) & v != ""]
  unique(v)
}

# Output path with "_LABELLED" and "same format" like Python, with safe .xls handling
output_path_same_format <- function(input_path, suffix = "_LABELLED") {
  base <- tools::file_path_sans_ext(input_path)
  ext <- get_ext(input_path)

  if (ext == ".xls") {
    # Writing .xls is not reliably supported in modern R; write .xlsx instead.
    return(paste0(base, suffix, ".xlsx"))
  }

  paste0(base, suffix, ext)
}

# Save output keeping same format when feasible
save_with_same_format <- function(df, input_path, out_path) {
  in_ext <- get_ext(input_path)

  if (in_ext == ".csv") {
    readr::write_csv(df, out_path, na = "")
    return(invisible(TRUE))
  }

  if (in_ext %in% c(".xlsx", ".xls")) {
    # Always write .xlsx (out_path already adjusted for .xls)
    writexl::write_xlsx(df, out_path)
    return(invisible(TRUE))
  }

  stop(sprintf("Unsupported input extension '%s'. Use .xlsx, .xls, or .csv.", in_ext))
}

# Get top N unmapped codes as named integer vector
top_unmapped_counts <- function(x, n = 15) {
  x <- x[!is.na(x)]
  if (length(x) == 0) return(setNames(integer(0), character(0)))
  tb <- sort(table(x), decreasing = TRUE)
  tb <- head(tb, n)
  setNames(as.integer(tb), names(tb))
}

# Download dictionary if missing and validate it's an actual workbook
download_dictionary_if_missing <- function(dest_path, url) {
  if (!file.exists(dest_path)) {
    message(sprintf("Downloading dictionary to: %s", dest_path))
    utils::download.file(url, dest_path, mode = "wb", quiet = TRUE)
  } else {
    message(sprintf("Dictionary already present, skipping download: %s", dest_path))
  }

  # Validate workbook readability (clearer error than later failures)
  ok <- TRUE
  tryCatch({
    readxl::excel_sheets(dest_path)
  }, error = function(e) {
    ok <<- FALSE
  })

  if (!ok) {
    stop(
      "Downloaded dictionary file is not a readable Excel workbook. ",
      "This can happen if the URL returned an HTML/login page instead of the .xlsx file. ",
      "Try downloading the codebook manually and placing it next to the microdata file, then rerun."
    )
  }
}

# ---------------------------
# Main
# ---------------------------
main <- function() {
  message("Reading microdata header: ", MICRODATA_FILE)
  micro_cols <- read_header_columns(MICRODATA_FILE)

  infr <- if ("qc_step0_date" %in% micro_cols) 2.0 else 1.0

  # Dictionary path in same folder as microdata
  micro_folder <- dirname(MICRODATA_FILE)
  dict_name <- if (infr == 1.0) "DIEM_codebook_V10.xlsx" else "DIEM_codebook_V20.xlsx"
  DICT_XLSX <- file.path(micro_folder, dict_name)

  dict_url <- if (infr == 1.0) DICT_URL_1 else DICT_URL_2
  download_dictionary_if_missing(DICT_XLSX, dict_url)

  message("Reading microdata: ", MICRODATA_FILE)
  micro <- read_microdata_any(MICRODATA_FILE)

  # Dictionary expected to be Excel workbook
  dict_ext <- get_ext(DICT_XLSX)
  if (!dict_ext %in% c(".xlsx", ".xls")) {
    stop(sprintf("DICT_XLSX must be an Excel file with sheets (got '%s').", dict_ext))
  }

  message("Reading coded-value dictionary: ", DICT_XLSX)
  sheet_names <- readxl::excel_sheets(DICT_XLSX)

  OUT_PATH <- output_path_same_format(MICRODATA_FILE, "_LABELLED")
  message("Output will be saved as: ", OUT_PATH)
  if (get_ext(MICRODATA_FILE) == ".xls") {
    message("Note: input is .xls, output will be written as .xlsx to avoid .xls export issues.")
  }

  # Derived fields (0=No, 1=Yes)
  sheet_names_clean <- str_trim(sheet_names)
  derived_fields <- character(0)
  if (DERIVED_FIELDS_SHEET %in% sheet_names_clean) {
    derived_fields <- read_derived_fields(DICT_XLSX, DERIVED_FIELDS_SHEET)
  }

  yes_no_mapping <- c("0" = "No", "1" = "Yes")

  # Collect label columns here and add them once (fast)
  label_columns <- list()

  # ------------------------------------------------------------------
  # Tracking / stats (mirror Python)
  # ------------------------------------------------------------------
  labelled_fields <- character(0)  # fields actually processed (dedicated mapping or derived yes/no)
  skipped_fields <- list()         # field -> reason (for candidates we expected to label but did not)
  missing_in_micro <- character(0) # fields present in dict/derived list but not in microdata columns
  per_field_details <- list()      # field -> list with counts and top unmapped codes

  fields_labelled <- 0L
  fields_skipped_empty <- 0L
  fields_low_match <- 0L
  total_values <- 0L
  mapped_values <- 0L

  # Candidate fields from dedicated sheets (excluding derived_fields)
  dedicated_sheet_fields <- sheet_names_clean[sheet_names_clean != DERIVED_FIELDS_SHEET]

  # 1) Dedicated dictionary sheets (one sheet per field)
  for (sheet_name in sheet_names) {
    field <- str_trim(sheet_name)
    if (field == DERIVED_FIELDS_SHEET) next

    if (!field %in% names(micro)) {
      missing_in_micro <- c(missing_in_micro, field)
      next
    }

    df_sheet <- tryCatch(readxl::read_excel(DICT_XLSX, sheet = sheet_name, .name_repair = "minimal"),
                         error = function(e) NULL)
    mapping <- build_mapping_from_sheet(df_sheet)
    if (is.null(mapping)) {
      skipped_fields[[field]] <- "Dictionary sheet missing 'code' and/or 'label' columns"
      next
    }

    normalized <- normalize_code(micro[[field]])
    non_null_idx <- which(!is.na(normalized))

    if (length(non_null_idx) == 0) {
      fields_skipped_empty <- fields_skipped_empty + 1L
      skipped_fields[[field]] <- "All values are null/empty in microdata"
      next
    }

    labels <- unname(mapping[normalized])  # NA for unmapped

    mapped_mask <- !is.na(labels[non_null_idx])
    mapped_here <- sum(mapped_mask)
    total_here <- length(non_null_idx)
    match_rate_here <- if (total_here > 0) mapped_here / total_here else 0
    unmapped_here <- total_here - mapped_here

    unmapped_codes <- normalized[non_null_idx][!mapped_mask]
    top_unmapped <- top_unmapped_counts(unmapped_codes, n = 15)

    per_field_details[[field]] <- list(
      source = "dedicated_sheet",
      non_null_values = total_here,
      mapped_values = mapped_here,
      unmapped_values = unmapped_here,
      match_rate = match_rate_here,
      top_unmapped_codes = top_unmapped
    )

    if (match_rate_here < MIN_OK_MATCH_RATE) {
      fields_low_match <- fields_low_match + 1L
    }

    total_values <- total_values + total_here
    mapped_values <- mapped_values + mapped_here

    if (ADD_LABEL_COLUMNS) {
      label_columns[[paste0(field, LABEL_SUFFIX)]] <- labels
    } else {
      micro[[field]] <- labels
    }

    labelled_fields <- unique(c(labelled_fields, field))
    fields_labelled <- fields_labelled + 1L
  }

  # 2) Derived yes/no fields listed in derived_fields sheet
  for (field in derived_fields) {
    if (!field %in% names(micro)) {
      missing_in_micro <- c(missing_in_micro, field)
      next
    }

    # Do not overwrite if already created by a dedicated mapping sheet
    if (ADD_LABEL_COLUMNS) {
      if (paste0(field, LABEL_SUFFIX) %in% names(label_columns)) next
    } else {
      if (field %in% labelled_fields) next
    }

    normalized <- normalize_code(micro[[field]])
    non_null_idx <- which(!is.na(normalized))

    if (length(non_null_idx) == 0) {
      fields_skipped_empty <- fields_skipped_empty + 1L
      skipped_fields[[field]] <- "All values are null/empty in microdata (derived yes/no)"
      next
    }

    labels <- unname(yes_no_mapping[normalized])

    mapped_mask <- !is.na(labels[non_null_idx])
    mapped_here <- sum(mapped_mask)
    total_here <- length(non_null_idx)
    match_rate_here <- if (total_here > 0) mapped_here / total_here else 0
    unmapped_here <- total_here - mapped_here

    unmapped_codes <- normalized[non_null_idx][!mapped_mask]
    top_unmapped <- top_unmapped_counts(unmapped_codes, n = 15)

    per_field_details[[field]] <- list(
      source = "derived_yes_no",
      non_null_values = total_here,
      mapped_values = mapped_here,
      unmapped_values = unmapped_here,
      match_rate = match_rate_here,
      top_unmapped_codes = top_unmapped
    )

    if (match_rate_here < MIN_OK_MATCH_RATE) {
      fields_low_match <- fields_low_match + 1L
    }

    total_values <- total_values + total_here
    mapped_values <- mapped_values + mapped_here

    if (ADD_LABEL_COLUMNS) {
      label_columns[[paste0(field, LABEL_SUFFIX)]] <- labels
    } else {
      micro[[field]] <- labels
    }

    labelled_fields <- unique(c(labelled_fields, field))
    fields_labelled <- fields_labelled + 1L
  }

  # Add all label columns in one shot (fast)
  micro_out <- micro
  if (ADD_LABEL_COLUMNS && length(label_columns) > 0) {
    for (nm in names(label_columns)) micro_out[[nm]] <- label_columns[[nm]]
  }

  # ------------------------------------------------------------------
  # Final reporting (mirror Python)
  # ------------------------------------------------------------------
  candidate_fields <- unique(c(
    intersect(dedicated_sheet_fields, names(micro)),
    intersect(derived_fields, names(micro))
  ))

  not_labelled_candidates <- sort(setdiff(candidate_fields, labelled_fields))

  labelled_with_unmapped <- character(0)
  for (f in sort(unique(labelled_fields))) {
    d <- per_field_details[[f]]
    if (!is.null(d) && isTRUE(d$unmapped_values > 0)) {
      labelled_with_unmapped <- c(labelled_with_unmapped, f)
    }
  }

  cat("\n--- Labelling summary (global) ---\n")
  cat(sprintf("Candidate fields to label (present in micro + in dict/derived list): %d\n", length(candidate_fields)))
  cat(sprintf("Fields labelled: %d\n", length(unique(labelled_fields))))
  cat(sprintf("Fields not labelled (among candidates): %d\n", length(not_labelled_candidates)))
  cat(sprintf("Fields skipped (all null): %d\n", fields_skipped_empty))
  cat(sprintf("Fields with low match rate (< %d%%): %d\n", as.integer(MIN_OK_MATCH_RATE * 100), fields_low_match))
  if (total_values > 0) {
    cat(sprintf("Overall value match rate: %.1f%%\n", 100 * mapped_values / total_values))
  }

  if (length(not_labelled_candidates) > 0) {
    cat("\n--- Fields not labelled (candidates) ---\n")
    for (f in not_labelled_candidates) {
      reason <- skipped_fields[[f]]
      if (is.null(reason)) reason <- "Unknown reason (check dictionary sheet format and data)"
      cat(sprintf("- %s: %s\n", f, reason))
    }
  }

  missing_in_micro_unique <- sort(unique(missing_in_micro))
  if (length(missing_in_micro_unique) > 0) {
    cat("\n--- Fields present in dict/derived but missing in microdata ---\n")
    for (f in missing_in_micro_unique) cat(sprintf("- %s\n", f))
  }

  if (length(labelled_with_unmapped) > 0) {
    cat("\n--- Labelled fields with unmapped codes (issues to review) ---\n")
    for (f in labelled_with_unmapped) {
      d <- per_field_details[[f]]
      cat(sprintf(
        "- %s | source=%s | non-null=%d | mapped=%d | unmapped=%d | match=%.1f%%\n",
        f,
        d$source,
        as.integer(d$non_null_values),
        as.integer(d$mapped_values),
        as.integer(d$unmapped_values),
        100 * d$match_rate
      ))

      top_unmapped <- d$top_unmapped_codes
      if (!is.null(top_unmapped) && length(top_unmapped) > 0) {
        preview <- paste0(names(top_unmapped), "(", as.integer(top_unmapped), ")", collapse = ", ")
        cat(sprintf("  top unmapped codes: %s\n", preview))
      }
    }
  }

  cat("\nSaving output to: ", OUT_PATH, "\n", sep = "")
  save_with_same_format(micro_out, MICRODATA_FILE, OUT_PATH)
  cat("\nDone.\n")
}

# Run
main()
