#' Quickly write a tidy data frame to xlsx, with options for customization
#'
#' @description
#' A wrapper to the [openxlsx][openxlsx::openxlsx] package optimized for writing
#' tidy data frames. Includes arguments to quickly add customization like:
#' - conditional formatting written as R expressions
#' - data validation rules based on a tidy dictionary structure
#' - column-specific worksheet protection
#' - custom column names with original variable-names hidden in the row below
#'
#' @param x A data frame, or list of data frames
#' @param file Filename to write to. If `NULL` the resulting workbook is
#'   returned as an openxlsx object of class "Workbook" rather than written to a
#'   file.
#' @param wb An openxlsx workbook object to write to. Defaults to a fresh
#'   workbook created with [`openxlsx::createWorkbook`]. Only need to update
#'   when repeatedly calling `qxl()` to add worksheets to an existing workbook.
#' @param sheet Optional character vector of worksheet names. If `NULL` (the
#'   default) and `x` is a named list of data frames, worksheet names are taken
#'   from `names(x)`. Otherwise, names default to "Sheet1", "Sheet2", ...
#' @param header Optional column header. Defaults to `NULL` in which case column
#'   names are taken directly from the data frame(s) in `x`, to create normal
#'   single-row headers. Can alternatively pass a named character vector to set
#'   custom names as the first row and a subheader with variable names as a
#'   hidden second row.
#'   ```
#'   header = c(
#'     mpg = "Miles per US gallon",
#'     cyl = "Number of cylinders",
#'     disp = "Engine displacement (cubic in.)
#'   )
#'   ```
#' @param style_head Style for the header row. Set with [`qstyle()`], or set to
#'   `NULL` for no header styling. Defaults to bold text.
#' @param hide_subhead Logical indicating whether to hide the subheader (if
#'   present). Defaults to TRUE.
#' @param style Optional style, or list of styles, set using [`qstyle()`].
#'   Accepts any number of [`qstyle()`] objects:
#'   ```
#'   style = qstyle(halign = "center")
#'   style = list(qstyle(halign = "center"), qstyle(rows = cyl > 4, bgFill = "#fddbc7"))
#'   ```
#' @param style1,style2,style3,style4,style5,style6,style7,style8,style9
#'   `r lifecycle::badge("deprecated")` Use `style` instead.
#' @param group Optional vector of one or more column names used to create
#'   alternating groupings of rows, with every other row grouping styled as per
#'   argument `group_style`. See section __Grouping rows__.
#' @param group_style Optional style to apply to alternating groupings of rows,
#'   as specified using argument `groups`. Set using [`qstyle()`]
#' @param group_border Optional border to apply to alternating groupings of
#'   rows, as specified using argument `groups`.
#' @param row_heights Numeric vector of row heights (in Excel units). The vector
#'   is recycled if shorter than the number of rows in `x`. Defaults to `NULL`
#'   to use default row heights.
#' @param col_widths Vector of column widths (in Excel units). Can be numeric or
#'   character, and may include keyword "auto" for automatic column sizing. The
#'   vector is recycled if shorter than the number of columns in `x`. Defaults
#'   to "auto".
#'
#'   Use named vector to give column widths for specific columns, where names
#'   represent column names of `x` or the keyword ".default" to set a default
#'   column width for all columns not otherwise specified. E.g.
#'
#'   ```
#'   # specify widths for cols mpg and cyl, all others default to "auto"
#'   col_widths <- c(mpg = 5, cyl = 10)
#'
#'   # specify widths for cols mpg and cyl, and explicit default for all others
#'   col_widths <- c(mpg = 5, cyl = 10, .default = 7)
#'   ```
#' @param cols_hide Vector of column names for columns to be hidden.
#' @param freeze_row Integer specifying a row to freeze at. Defaults to `1` to
#'   add a freeze below the header row. Set to `0` or `NULL` to omit freezing.
#' @param freeze_col Integer specifying a column to freeze at. Defaults to
#'   `NULL`. Set to `0` or `NULL` to omit freezing.
#' @param protect Optional function specifying how to protect worksheet
#'   components from user modification. See function [`qprotect`].
#' @param validate Optional specification of list-style data validation for one
#'   or more columns. Can specify either as a list of vectors giving options for
#'   one or more column in `x`, e.g.:
#'   ```
#'   list(
#'     var_x = c("Yes", "No"),
#'     var_y = c("Small", "Medium", "Large")
#'   )
#'   ```
#'   or as a data.frame where the first column gives column names and the
#'   second column gives corresponding options, e.g.:
#'   ```
#'   data.frame(
#'     col = c("var_x", "var_x", "var_y", "var_y", "var_y"),
#'     val = c("Yes", "No", "Small", "Medium", "Large")
#'   )
#'   ```
#'   Validation options are written/appended to a hidden worksheet named
#'   "valid_options".
#' @param validate_cond Optional specification of conditional list-style
#'   validation, where the set of values to be allowed in a given column depends
#'   on the corresponding value within one or more other columns (e.g. the
#'   allowed values in column 'city' depend on the corresponding value in
#'   columns 'country' and 'province'). Must be a data.frame with at least two
#'   columns, where the first column(s) give the conditional entries (e.g.
#'   'country', 'province') and the last column gives the corresponding allowed
#'   entries (e.g. 'city') to be implemented as data validation. The column
#'   names in `validate_cond` should match the relevant columns within `x`.
#'
#'   Note that in the current implementation validation is based on values in
#'   the conditional column(s) of `x` at the time the workbook is written, and
#'   will not update in real time if those values are later changed.
#' @param validate_cond_all Optional vector of value(s) to always allow,
#'   independent of the value in the conditional column (e.g. "Unknown").
#' @param validate_cascade Optional data frame specifying cascading dropdown
#'   validation. Each column defines one level of the cascade, and all column
#'   names must match columns in `x`. Each level's dropdown options are
#'   dynamically filtered in Excel based on the value selected in the previous
#'   level. For example:
#'   ```
#'   data.frame(
#'     adm1 = c("Ontario", "Ontario", "Quebec", "Quebec"),
#'     adm2 = c("Toronto", "Ottawa",  "Montreal", "Quebec City")
#'   )
#'   ```
#'   Selecting "Ontario" in `adm1` will restrict the `adm2` dropdown to
#'   "Toronto" and "Ottawa".
#' @param filter Logical indicating whether to add column filters.
#' @param filter_cols Tidy-selection specifying which columns to filter. Only
#'   used if `filter` is `TRUE`. Defaults to `everything()` to select all
#'   columns.
#' @param date_format Excel format for date columns. Defaults to "yyyy-mm-dd".
#' @param zoom Integer specifying initial zoom percentage. Defaults to 130.
#' @param overwrite Logical indicating whether to overwrite existing file.
#'   Defaults to `TRUE`
#'
#' @section Grouping rows:
#' Given a dataset with multiple rows per group (e.g. repeated observations on a
#' given individual), it can sometimes be useful to uniquely stylize alternating
#' groups to allow for quick visual distinction of the rows belonging to any
#' given group.
#'
#' Given one or more grouping columns specified using argument `groups`, the
#' `qxl` function arranges the rows of the resulting worksheet by group and then
#' applies the style `group_style` to the rows in every *other* group, to create
#' an alternating pattern. The alternating pattern is achieved by first creating
#' a group index variable called `g` which is assigned a value of either `1` or
#' `0`: `1` for the 1st group, `0` for the 2nd, `1` for the 3rd, `0` for the
#' 4th, etc. The style specified by `group_style` is then applied conditionally
#' to rows where `g == 0`. The grouping variable is written in column A, which
#' is hidden.
#'
#' @return
#' If argument `file` is not specified, returns an openxlsx workbook object.
#' Otherwise writes workbook to file with no return.
#'
#' @examples
#' library(datasets)
#' qxl(mtcars, file = tempfile(fileext = ".xlsx"))
#'
#' @export qxl
qxl <- function(
  x,
  file = NULL,
  wb = openxlsx::createWorkbook(),
  sheet = NULL,
  header = NULL,
  style_head = qstyle(rows = 1, textDecoration = "bold"),
  hide_subhead = TRUE,
  style = NULL,
  style1 = NULL,
  style2 = NULL,
  style3 = NULL,
  style4 = NULL,
  style5 = NULL,
  style6 = NULL,
  style7 = NULL,
  style8 = NULL,
  style9 = NULL,
  group,
  group_style = qstyle(bgFill = "#ffcccb"),
  group_border = TRUE,
  row_heights = NULL,
  col_widths = "auto",
  cols_hide = NULL,
  freeze_row = 1L,
  freeze_col = NULL,
  protect,
  validate = NULL,
  validate_cond = NULL,
  validate_cond_all = NULL,
  validate_cascade = NULL,
  filter = FALSE,
  filter_cols = everything(),
  zoom = 120L,
  date_format = "yyyy-mm-dd",
  overwrite = TRUE
) {
  # handle deprecated style1-style9 args
  deprecated_styles <- list(style1, style2, style3, style4, style5, style6, style7, style8, style9)
  deprecated_styles <- Filter(Negate(is.null), deprecated_styles)

  if (length(deprecated_styles) > 0) {
    warning(
      "Arguments `style1`-`style9` are deprecated. ",
      "Use argument `style` with a list of `qstyle()` objects instead.",
      call. = FALSE
    )
    if (is.null(style)) {
      style <- deprecated_styles
    } else {
      style <- c(if (inherits(style, "qstyle")) list(style) else style, deprecated_styles)
    }
  }

  # normalise style to a list of qstyle objects
  if (inherits(style, "qstyle")) {
    style <- list(style)
  }

  # convert x to list if single data frame
  if (is.data.frame(x)) {
    x <- list(x)
  }

  # prepare sheet names
  if (is.null(sheet)) {
    if (is.null(names(x))) {
      names(x) <- paste0("Sheet", seq_along(x))
    }
  } else {
    if (length(sheet) != length(x)) {
      stop(
        "If specified, argument `sheets` must be the same length as `x`",
        call. = FALSE
      )
    }
    names(x) <- sheet
  }

  # truncate sheet names if over 31 characters
  if (any(nchar(names(x)) > 31)) {
    warning("Truncating sheet name(s) to 31 characters")
    names(x) <- substring(names(x), 1, 29)
  }

  # de-duplicate sheet names
  if (length(unique(names(x))) < length(names(x))) {
    warning("Deduplicating sheet names")
    names(x) <- make.unique(substring(names(x), 1, 28), sep = "_")
  }

  sheet <- names(x)

  # prepare header
  if (is.null(header)) {
    header <- lapply(x, names)
  } else {
    header <- replicate(length(x), header, simplify = FALSE)
  }

  # write first sheet
  wb <- qxl_(
    x = x[[1]],
    wb = wb,
    sheet = sheet[1],
    header = header[[1]],
    style_head = style_head,
    hide_subhead = hide_subhead,
    style = style,
    group = group,
    group_style = group_style,
    group_border = group_border,
    row_heights = row_heights,
    col_widths = col_widths,
    cols_hide = cols_hide,
    freeze_row = freeze_row,
    freeze_col = freeze_col,
    protect = protect,
    validate = validate,
    validate_cond = validate_cond,
    validate_cond_all = validate_cond_all,
    validate_cascade = validate_cascade,
    filter = filter,
    filter_cols = filter_cols,
    zoom = zoom,
    date_format = date_format
  )

  # write subsequent sheets (if any)
  for (i in seq_along(x)[-1]) {
    wb <- qxl_(
      x = x[[i]],
      wb = wb,
      sheet = sheet[i],
      header = header[[i]],
      style_head = style_head,
      hide_subhead = hide_subhead,
      style = style,
      group = group,
      group_style = group_style,
      group_border = group_border,
      row_heights = row_heights,
      col_widths = col_widths,
      cols_hide = cols_hide,
      freeze_row = freeze_row,
      freeze_col = freeze_col,
      protect = protect,
      validate = validate,
      validate_cond = validate_cond,
      validate_cond_all = validate_cond_all,
      validate_cascade = validate_cascade,
      filter = filter,
      filter_cols = filter_cols,
      zoom = zoom,
      date_format = date_format
    )
  }

  ### return -------------------------------------------------------------------
  if (!is.null(file)) {
    qxl::qwrite(wb, file = file, overwrite = overwrite)
  } else {
    return(wb)
  }
}


#' @noRd
#' @import openxlsx
#' @importFrom stats setNames
#' @importFrom tidyr unnest
#' @importFrom dplyr everything select mutate arrange relocate `%>%` all_of
#'   bind_rows across cur_group_id group_by ungroup n summarize left_join
#'   distinct
#' @importFrom rlang enquo quo_get_expr .data .env `!!`
#'
qxl_ <- function(
  x,
  file = NULL,
  wb = openxlsx::createWorkbook(),
  sheet = "Sheet1",
  header = names(x),
  style_head = qstyle(rows = 1, textDecoration = "bold"),
  hide_subhead = TRUE,
  style = NULL,
  group,
  group_style = qstyle(bgFill = "#ffcccb"),
  group_border = TRUE,
  row_heights = NULL,
  col_widths = "auto",
  cols_hide = NULL,
  freeze_row = 1L,
  freeze_col = NULL,
  protect,
  validate = NULL,
  validate_cond = NULL,
  validate_cond_all = NULL,
  validate_cascade = NULL,
  filter = FALSE,
  filter_cols = everything(),
  zoom = 120L,
  date_format = "yyyy-mm-dd",
  overwrite = TRUE
) {
  ### grouping -----------------------------------------------------------------
  has_group <- !missing(group)

  if (has_group) {
    x <- x %>%
      arrange(across(all_of(group))) %>%
      group_by(across(all_of(group))) %>%
      mutate(g = cur_group_id() %% 2L, .before = 1L) %>%
      ungroup()

    header <- if (!is.null(names(header))) c("g" = "Grouping", header) else c("g", header)
  }

  ### initialize ---------------------------------------------------------------
  options("openxlsx.dateFormat" = date_format)
  wb <- openxlsx::copyWorkbook(wb) # to avoid overwriting global env
  openxlsx::addWorksheet(wb, sheet, zoom = zoom)

  if (!is.null(names(header))) {
    data_start_row <- 3L

    openxlsx::writeData(
      wb,
      sheet = sheet,
      rbind(stats::setNames(names(header), header)),
      colNames = TRUE
    )

    if (hide_subhead) {
      suppressWarnings(
        # bug in openxlsx generates unnecessary warning
        openxlsx::groupRows(
          wb,
          sheet,
          rows = 2,
          hidden = TRUE
        )
      )
    }

    openxlsx::addStyle(
      wb,
      sheet,
      rows = 2,
      cols = seq_len(ncol(x)),
      style = openxlsx::createStyle(textDecoration = "bold", halign = "left")
    )
  } else {
    data_start_row <- 2L

    openxlsx::writeData(
      wb,
      sheet = sheet,
      cols_to_chr(x[0, , drop = FALSE]),
      colNames = TRUE
    )
  }

  openxlsx::writeData(
    wb,
    sheet = sheet,
    x,
    startRow = data_start_row,
    colNames = FALSE
  )

  nrow_x <- nrow(x) + data_start_row - 1L # number of rows in Excel, incl. header
  ncol_x <- ncol(x)

  ### row height and col widths ------------------------------------------------
  if (!is.null(row_heights)) {
    openxlsx::setRowHeights(wb, sheet, rows = 1:nrow_x, heights = row_heights)
  }

  if (!is.null(col_widths)) {
    if (!is.null(names(col_widths))) {
      names_extra <- setdiff(names(col_widths), c(names(x), ".default"))

      if (length(names_extra) > 0) {
        stop(
          "Names within `col_names` do not match keyword \".default\" or colnames of `x`: ",
          paste_collapse_c(names_extra),
          call. = FALSE
        )
      }

      col_widths_raw <- col_widths

      width_default <- ifelse(
        ".default" %in% names(col_widths),
        col_widths[names(col_widths) == ".default"],
        "auto"
      )

      col_widths <- rep(as.character(width_default), ncol_x)

      for (j in seq_along(col_widths_raw)) {
        col_widths[names(x) %in% names(col_widths_raw)[j]] <- as.character(col_widths_raw[j])
      }
    }

    openxlsx::setColWidths(
      wb,
      sheet,
      cols = seq_len(ncol_x),
      widths = col_widths,
      hidden = names(x) %in% cols_hide
    )
  }

  ### freeze panes -------------------------------------------------------------
  if (is.null(freeze_row)) {
    freeze_row <- 0L
  }
  if (is.null(freeze_col)) {
    freeze_col <- 0L
  }

  openxlsx::freezePane(
    wb,
    sheet,
    firstActiveRow = freeze_row + 1L,
    firstActiveCol = freeze_col + 1L
  )

  ### protection ---------------------------------------------------------------
  if (!missing(protect)) {
    # protect
    protect_args <- c(
      wb = wb,
      sheet = sheet,
      protect[!names(protect) %in% c("cols", "row_buffer", "col_buffer")]
    )

    do.call(openxlsx::protectWorksheet, protect_args)

    # unprotect
    openxlsx::addStyle(
      wb,
      sheet,
      style = unprotect(),
      rows = seq(data_start_row, nrow_x + protect$row_buffer, by = 1L),
      cols = col_selection(x, protect$cols, invert = TRUE),
      gridExpand = TRUE,
      stack = TRUE
    )

    # unprotect buffer cols
    if (protect$col_buffer > 0) {
      openxlsx::addStyle(
        wb,
        sheet,
        style = unprotect(),
        rows = seq(1L, nrow_x + protect$row_buffer, by = 1L),
        cols = seq_len(protect$col_buffer) + ncol_x,
        gridExpand = TRUE,
        stack = TRUE
      )
    }
  }

  ### filter -------------------------------------------------------------------
  if (filter) {
    cols_filter <- col_selection(x, rlang::enquo(filter_cols))
    openxlsx::addFilter(wb, sheet, rows = 1, cols = cols_filter)
  }

  ### row styles ---------------------------------------------------------------
  if (!is.null(style_head)) {
    apply_row_style(
      data = x,
      wb = wb,
      sheet = sheet,
      style = style_head,
      data_start_row = data_start_row,
      nrow_x = nrow_x
    )
  }

  for (s in style) {
    apply_row_style(
      data = x,
      wb = wb,
      sheet = sheet,
      style = s,
      data_start_row = data_start_row,
      nrow_x = nrow_x
    )
  }

  if (has_group) {
    openxlsx::conditionalFormatting(
      wb = wb,
      sheet = sheet,
      cols = seq_len(ncol_x),
      rows = data_start_row:nrow_x,
      rule = paste0("$A", data_start_row, "==0"),
      style = group_style$style
    )

    suppressWarnings(
      openxlsx::groupColumns(
        wb = wb,
        sheet = sheet,
        cols = 1,
        hidden = TRUE
      )
    )

    if (group_border) {
      df_borders <- x %>%
        mutate(temp__rowid__ = 1:n()) %>%
        group_by(across(all_of(group))) %>%
        summarize(temp__rowid__ = min(.data$temp__rowid__) + .env$data_start_row - 1L, .groups = "drop")

      openxlsx::addStyle(
        wb = wb,
        sheet = sheet,
        style = openxlsx::createStyle(border = "top"),
        rows = df_borders$temp__rowid__,
        cols = seq_len(ncol_x),
        gridExpand = TRUE,
        stack = TRUE
      )
    }
  }

  ### validation ---------------------------------------------------------------
  if (!is.null(validate)) {
    if (is.data.frame(validate)) {
      validate_df <- validate
    } else {
      validate_df <- list_to_df(validate)
    }

    if ("valid_options" %in% wb$sheet_names) {
      opt <- openxlsx::readWorkbook(wb, "valid_options", colNames = FALSE)
      valid_start <- nrow(opt) + 2L
    } else {
      valid_start <- 1L
      openxlsx::addWorksheet(wb, "valid_options", visible = FALSE)
    }

    openxlsx::writeData(
      wb,
      "valid_options",
      x = validate_df,
      startRow = valid_start,
      colNames = FALSE
    )

    for (j in unique(validate_df[[1]])) {
      i_rng <- range(which(validate_df[[1]] %in% j)) + valid_start - 1L
      excel_range <- paste0("'valid_options'!", "$B$", i_rng[1], ":", "$B$", i_rng[2])

      suppressWarnings(
        openxlsx::dataValidation(
          wb,
          sheet,
          cols = which(names(x) %in% j),
          rows = data_start_row:nrow_x,
          type = "list",
          value = excel_range,
          allowBlank = TRUE,
          showInputMsg = TRUE,
          showErrorMsg = TRUE
        )
      )
    }
  }

  ### conditional validation ---------------------------------------------------
  if (!is.null(validate_cond)) {
    if (is.data.frame(validate_cond)) {
      validate_cond_df <- validate_cond
    } else {
      validate_cond_df <- list_to_df(validate_cond)
    }

    which_col_validation <- ncol(validate_cond_df)
    excel_col_validation <- LETTERS[which_col_validation]

    col_validation <- names(validate_cond_df)[which_col_validation]
    cols_cond <- setdiff(names(validate_cond_df), col_validation)

    if (!is.null(validate_cond_all)) {
      validate_cond_all_df <- unique(x[, cols_cond, drop = FALSE]) %>%
        mutate(replacement = list(validate_cond_all)) %>%
        tidyr::unnest("replacement") %>%
        stats::setNames(c(cols_cond, col_validation))

      validate_cond_df <- bind_rows(validate_cond_df, validate_cond_all_df) %>%
        arrange(across(all_of(cols_cond)))
    }

    if (!all(cols_cond %in% names(x))) {
      stop("Columns in argument `validate_cond` must also be column names in `x`")
    }

    # TODO: this needs to be moved up so that the new colname is actually written
    if (!col_validation %in% names(x)) {
      x[[col_validation]] <- NA_character_
    }

    if ("valid_options_cond" %in% wb$sheet_names) {
      opt <- openxlsx::readWorkbook(wb, "valid_options_cond", colNames = FALSE)
      valid_cond_start <- nrow(opt) + 2L
    } else {
      valid_cond_start <- 1L
      openxlsx::addWorksheet(wb, "valid_options_cond", visible = FALSE)
    }

    openxlsx::writeData(
      wb,
      "valid_options_cond",
      x = validate_cond_df,
      startRow = valid_cond_start,
      colNames = FALSE
    )

    validate_indices <- validate_cond_df %>%
      mutate(rowid = seq_len(n())) %>%
      group_by(across(all_of(cols_cond))) %>%
      summarize(
        range_min = min(.data$rowid) + .env$valid_cond_start - 1L,
        range_max = max(.data$rowid) + .env$valid_cond_start - 1L,
        .groups = "drop"
      )

    x_validate <- x %>%
      left_join(validate_indices, by = cols_cond) %>%
      mutate(
        excel_range = paste0(
          "'valid_options_cond'!$",
          .env$excel_col_validation,
          "$",
          .data$range_min,
          ":$",
          .env$excel_col_validation,
          "$",
          .data$range_max
        )
      )

    for (i in seq_len(nrow(x))) {
      if (!is.na(x_validate$range_min[i]) & !is.na(x_validate$range_max[i])) {
        suppressWarnings(
          openxlsx::dataValidation(
            wb,
            sheet,
            cols = which(names(x) %in% col_validation),
            rows = data_start_row + i - 1,
            type = "list",
            value = x_validate$excel_range[i],
            allowBlank = TRUE,
            showInputMsg = TRUE,
            showErrorMsg = TRUE
          )
        )
      }
    }
  }

  ### cascade validation -------------------------------------------------------
  if (!is.null(validate_cascade)) {
    if (!is.data.frame(validate_cascade)) {
      stop("`validate_cascade` must be a data frame", call. = FALSE)
    }
    if (ncol(validate_cascade) < 2L) {
      stop("`validate_cascade` must have at least 2 columns", call. = FALSE)
    }
    if (!all(names(validate_cascade) %in% names(x))) {
      stop(
        "All column names in `validate_cascade` must also be column names in `x`",
        call. = FALSE
      )
    }

    casc_sheet <- "valid_options_cascade"

    if (casc_sheet %in% wb$sheet_names) {
      casc_existing <- openxlsx::readWorkbook(wb, casc_sheet, colNames = FALSE)
      next_col <- ncol(casc_existing) + 1L
    } else {
      next_col <- 1L
      openxlsx::addWorksheet(wb, casc_sheet, visible = FALSE)
    }

    casc_cols <- names(validate_cascade)

    # level 1: plain list of unique first-column values
    lvl1_vals <- unique(validate_cascade[[1L]])
    lvl1_col <- next_col
    openxlsx::writeData(
      wb,
      casc_sheet,
      x = data.frame(v = lvl1_vals),
      startCol = lvl1_col,
      startRow = 1L,
      colNames = FALSE
    )
    next_col <- next_col + 1L

    suppressWarnings(
      openxlsx::dataValidation(
        wb,
        sheet,
        cols = which(names(x) == casc_cols[1L]),
        rows = data_start_row:nrow_x,
        type = "list",
        value = paste0(
          "'",
          casc_sheet,
          "'!$",
          COLS_EXCEL[lvl1_col],
          "$1:$",
          COLS_EXCEL[lvl1_col],
          "$",
          length(lvl1_vals)
        ),
        allowBlank = TRUE,
        showInputMsg = TRUE,
        showErrorMsg = TRUE
      )
    )

    # levels 2..n: one named range per ancestor-key, INDIRECT validation
    for (i in seq(2L, ncol(validate_cascade))) {
      child_col <- casc_cols[i]

      # compound key from all ancestor columns (levels 1..i-1)
      ancestor_keys <- apply(
        validate_cascade[seq_len(i - 1L)],
        1L,
        paste,
        collapse = "__"
      )
      groups <- lapply(split(validate_cascade[[child_col]], ancestor_keys), unique)

      lookup_keys <- character(0L)
      lookup_ids <- character(0L)

      for (key in names(groups)) {
        child_vals <- groups[[key]]
        range_name <- paste0("casc_", next_col)
        openxlsx::writeData(
          wb,
          casc_sheet,
          x = data.frame(v = child_vals),
          startCol = next_col,
          startRow = 1L,
          colNames = FALSE
        )
        openxlsx::createNamedRegion(
          wb,
          casc_sheet,
          cols = next_col,
          rows = seq_along(child_vals),
          name = range_name,
          overwrite = TRUE
        )
        lookup_keys <- c(lookup_keys, key)
        lookup_ids <- c(lookup_ids, range_name)
        next_col <- next_col + 1L
      }

      # write lookup table: keys column + ids column
      key_col_idx <- next_col
      id_col_idx <- next_col + 1L
      n_groups <- length(lookup_keys)
      openxlsx::writeData(
        wb,
        casc_sheet,
        x = data.frame(v = lookup_keys),
        startCol = key_col_idx,
        startRow = 1L,
        colNames = FALSE
      )
      openxlsx::writeData(
        wb,
        casc_sheet,
        x = data.frame(v = lookup_ids),
        startCol = id_col_idx,
        startRow = 1L,
        colNames = FALSE
      )
      next_col <- next_col + 2L

      # INDIRECT formula: MATCH the ancestor key against the lookup table, then
      # INDEX the corresponding range name and INDIRECT into it. This avoids any
      # text transformation of cell values, so Unicode characters are handled
      # correctly.
      ancestor_letters <- COLS_EXCEL[which(names(x) %in% casc_cols[seq_len(i - 1L)])]
      key_col_letter <- COLS_EXCEL[key_col_idx]
      id_col_letter <- COLS_EXCEL[id_col_idx]

      if (length(ancestor_letters) == 1L) {
        key_expr <- sprintf("$%s%d", ancestor_letters, data_start_row)
      } else {
        parts <- sprintf("$%s%d", ancestor_letters, data_start_row)
        key_expr <- paste0("CONCATENATE(", paste(parts, collapse = ',"__",'), ")")
      }

      indirect_formula <- sprintf(
        "INDIRECT(INDEX('valid_options_cascade'!$%s$1:$%s$%d,MATCH(%s,'valid_options_cascade'!$%s$1:$%s$%d,0)))",
        id_col_letter,
        id_col_letter,
        n_groups,
        key_expr,
        key_col_letter,
        key_col_letter,
        n_groups
      )
      suppressWarnings(
        openxlsx::dataValidation(
          wb,
          sheet,
          cols = which(names(x) == child_col),
          rows = data_start_row:nrow_x,
          type = "list",
          value = indirect_formula,
          allowBlank = TRUE,
          showInputMsg = TRUE,
          showErrorMsg = TRUE
        )
      )
    }
  }

  ### return -------------------------------------------------------------------
  if (!is.null(file)) {
    qwrite(wb, file = file, overwrite = overwrite)
  } else {
    return(wb)
  }
}


#' @noRd
list_to_df <- function(x) {
  data.frame(
    x1 = rep(names(x), times = lengths(x)),
    x2 = as.character(unlist(x)),
    stringsAsFactors = FALSE
  )
}


#' @noRd
resolve_style_rows <- function(rows_quo, data_start_row, nrow_x) {
  if (is_quo_numeric(rows_quo)) {
    return(rlang::eval_tidy(rows_quo))
  }
  rows_chr <- rlang::eval_tidy(rows_quo)
  if (rows_chr == "data") {
    if (nrow_x >= data_start_row) seq(data_start_row, nrow_x) else NULL
  } else {
    seq(1L, nrow_x) # "all"
  }
}


#' @noRd
#' @importFrom rlang eval_tidy
apply_row_style <- function(data, wb, sheet, style, data_start_row, nrow_x) {
  cols_style <- col_selection(data, style$cols)

  # static row range: numeric indices or keyword ("data" / "all")
  if (is_quo_numeric(style$rows) || is_quo_character(style$rows)) {
    rows_int <- resolve_style_rows(style$rows, data_start_row, nrow_x)
    if (!is.null(rows_int)) {
      openxlsx::addStyle(
        wb,
        sheet,
        style = style$style,
        rows = rows_int,
        cols = cols_style,
        gridExpand = TRUE,
        stack = TRUE
      )
    }
    return(invisible(NULL))
  }

  # conditional formatting (expression rows)
  cond <- rlang::quo_get_expr(style$rows)
  has_dotx <- ".x" %in% all.vars(cond)
  cols_named <- col_selection(data, style$cols, index = FALSE)

  for (j in cols_named) {
    cond_j <- if (has_dotx) {
      do.call("substitute", list(cond, list(.x = as.name(j))))
    } else {
      cond
    }
    rule <- expr_to_excel(rlang::enquo(cond_j), data, row_start = data_start_row)
    openxlsx::conditionalFormatting(
      wb,
      sheet,
      cols = which(names(data) == j),
      rows = data_start_row:nrow_x,
      rule = rule,
      style = style$style
    )
  }
}
