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
#'   single-row headers. Can alternatively pass a named character
#'   vector to set custom names as the first row and variable names as a hidden
#'   second row.
#'   ```
#'   header = c(
#'     mpg = "Miles per US gallon",
#'     cyl = "Number of cylinders",
#'     disp = "Engine displacement (cubic in.)
#'   )
#'   ```
#' @param style_head Style for the header row. Set with [`qstyle()`], or set to
#'   `NULL` for no header styling. Defaults to bold text.
#' @param style1 Optional style to set using [`qstyle()`]
#' @param style2 Optional style to set using [`qstyle()`]
#' @param style3 Optional style to set using [`qstyle()`]
#' @param row_heights Numeric vector of row heights (in Excel units). The vector
#'   is recycled if shorter than the number of rows in `x`. Defaults to `NULL`
#'   to use default row heights.
#' @param col_widths Numeric vector of column widths (in Excel units), or string
#'   "auto" to use automatic sizing. Defaults to "auto". The vector is recycled
#'   if shorter than the number of columns in `x`.
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
#' @param filter Logical indicating whether to add column filters.
#' @param filter_cols Tidy-selection specifying which columns to filter. Only
#'   used if `filter` is `TRUE`. Defaults to `everything()` to select all
#'   columns.
#' @param date_format Excel format for date columns. Defaults to "yyyy-mm-dd".
#' @param zoom Integer specifying initial zoom percentage. Defaults to 130.
#' @param overwrite Logical indicating whether to overwrite existing file.
#'   Defaults to `TRUE`
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
qxl <- function(x,
                file = NULL,
                wb = openxlsx::createWorkbook(),
                sheet = NULL,
                header = NULL,
                style_head = qstyle(rows = 1, textDecoration = "bold"),
                style1 = NULL,
                style2 = NULL,
                style3 = NULL,
                row_heights = NULL,
                col_widths = "auto",
                freeze_row = 1L,
                freeze_col = NULL,
                protect,
                validate = NULL,
                filter = FALSE,
                filter_cols = everything(),
                zoom = 120L,
                date_format = "yyyy-mm-dd",
                overwrite = TRUE) {


  # convert x to list if single data frame
  if (is.data.frame(x)) { x <- list(x) }

  # prepare sheet names
  if (is.null(sheet)) {
    if (is.null(names(x))) {
      sheet <- paste0("Sheet", seq_along(x))
    } else {
      sheet <- names(x)
    }
  } else {
    if (length(sheet) != length(x)) {
      stop(
        "If specified, argument `sheets` must be the same length as `x`",
        call. = FALSE
      )
    }
  }

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
    style1 = style1,
    style2 = style2,
    style3 = style3,
    row_heights = row_heights,
    col_widths = col_widths,
    freeze_row = freeze_row,
    freeze_col = freeze_col,
    protect = protect,
    validate = validate,
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
      style1 = style1,
      style2 = style2,
      style3 = style3,
      row_heights = row_heights,
      col_widths = col_widths,
      freeze_row = freeze_row,
      freeze_col = freeze_col,
      protect = protect,
      validate = validate,
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
#' @importFrom dplyr everything
#' @importFrom rlang enquo quo_get_expr
qxl_ <- function(x,
                 file = NULL,
                 wb = openxlsx::createWorkbook(),
                 sheet = "Sheet1",
                 header = names(x),
                 style_head = qstyle(rows = 1, textDecoration = "bold"),
                 style1 = NULL,
                 style2 = NULL,
                 style3 = NULL,
                 row_heights = NULL,
                 col_widths = "auto",
                 freeze_row = 1L,
                 freeze_col = NULL,
                 protect,
                 validate = NULL,
                 filter = FALSE,
                 filter_cols = everything(),
                 zoom = 120L,
                 date_format = "yyyy-mm-dd",
                 overwrite = TRUE) {

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
    suppressWarnings( # bug in openxlsx generates unnecessary warning
      openxlsx::groupRows(
        wb,
        sheet,
        rows = 2,
        hidden = TRUE
      )
    )
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
      x[0,],
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

  nrow_x <- nrow(x) + data_start_row - 1L
  ncol_x <- ncol(x)


  ### row height and col widths ------------------------------------------------
  if (!is.null(row_heights)) {
    openxlsx::setRowHeights(wb, sheet, rows = 1:nrow_x, heights = row_heights)
  }
  if (!is.null(col_widths)) {
    openxlsx::setColWidths(wb, sheet, cols = 1:ncol_x, widths = col_widths)
  }


  ### freeze panes -------------------------------------------------------------
  if (is.null(freeze_row)) freeze_row <- 0L
  if (is.null(freeze_col)) freeze_col <- 0L

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
      protect[!names(protect) %in% c("cols", "row_buffer")]
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

  if (!is.null(style1)) {
    apply_row_style(
      data = x,
      wb = wb,
      sheet = sheet,
      style = style1,
      data_start_row = data_start_row,
      nrow_x = nrow_x
    )
  }

  if (!is.null(style2)) {
    apply_row_style(
      data = x,
      wb = wb,
      sheet = sheet,
      style = style2,
      data_start_row = data_start_row,
      nrow_x = nrow_x
    )
  }

  if (!is.null(style3)) {
    apply_row_style(
      data = x,
      wb = wb,
      sheet = sheet,
      style = style3,
      data_start_row = data_start_row,
      nrow_x = nrow_x
    )
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
      wb, "valid_options",
      x = validate_df,
      startRow = valid_start,
      colNames = FALSE
    )

    for (j in unique(validate_df[[1]])) {

      i_rng <- range(which(validate_df[[1]] %in% j)) + valid_start - 1L
      excel_range <- paste0("'valid_options'!", "$B$", i_rng[1], ":", "$B$", i_rng[2])

      openxlsx::dataValidation(
        wb,
        sheet,
        col = which(names(x) %in% j),
        rows = data_start_row:nrow_x,
        type = "list",
        value = excel_range,
        allowBlank = TRUE,
        showInputMsg = TRUE,
        showErrorMsg = TRUE
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
#' @importFrom rlang eval_tidy
apply_row_style <- function(data,
                            wb,
                            sheet,
                            style,
                            data_start_row,
                            nrow_x) {

  cols_style <- col_selection(data, style$cols)

  if (is_quo_numeric(style$rows)) {
    # row-specific formatting (non-conditional)
    openxlsx::addStyle(
      wb,
      sheet,
      style = style$style,
      rows = rlang::eval_tidy(style$rows),
      cols = cols_style,
      gridExpand = TRUE,
      stack = TRUE
    )
  } else if (is_quo_character(style$rows)) {
    # character keyword 'data' or 'all'
    rows_chr <- rlang::eval_tidy(style$rows)
    if (rows_chr == "data") {
      rows_int <- seq(data_start_row, nrow_x, by = 1L)
    } else if (rows_chr == "all") (
      rows_int <- seq(1L, nrow_x, by = 1L)
    )

    openxlsx::addStyle(
      wb,
      sheet,
      style = style$style,
      rows = rows_int,
      cols = cols_style,
      gridExpand = TRUE,
      stack = TRUE
    )

  } else {

    # extract expression
    cond <- rlang::quo_get_expr(style$rows)

    # conditional formatting
    has_dotx <- ".x" %in% all.vars(cond)

    if (has_dotx) {
      # conditional formatting with .x-selector
      cols_dotx <- col_selection(data, style$cols, index = FALSE)
      for (j in cols_dotx) {
        cond_j <- do.call(
          "substitute",
          list(cond, list(.x = str2lang(j)))
        )
        cond_excel_j <- expr_to_excel(
          rlang::enquo(cond_j),
          data,
          row_start = data_start_row
        )
        openxlsx::conditionalFormatting(
          wb,
          sheet,
          cols = which(names(data) %in% j),
          rows = data_start_row:nrow_x,
          rule = cond_excel_j,
          style = style$style
        )
      }
    } else {
      # conditional formatting, no .x-selector
      rule <- expr_to_excel(
        style$rows, # already enquo
        data,
        row_start = data_start_row
      )
      for (j in cols_style) {
        openxlsx::conditionalFormatting(
          wb,
          sheet,
          cols = j,
          rows = data_start_row:nrow_x,
          rule = rule,
          style = style$style
        )
      }
    }
  }
}
