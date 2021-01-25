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
#' @param x A data frame
#' @param file Filename to write to. If `NULL` the resulting workbook is
#'   returned as an openxlsx object of class "Workbook" rather than written to a
#'   file.
#' @param sheet Name of sheet to write to. Defaults to "Sheet1".
#' @param header Column header. Defaults to `names(x)` to create normal
#'   single-row header of column names. Can alternatively pass a named character
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
#' @import openxlsx
#' @importFrom dplyr everything
#' @importFrom rlang enquo
#' @export qxl
qxl <- function(x,
                file = NULL,
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
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, sheet, zoom = zoom)

  has_header <- !is.null(names(header))

  if (has_header) {

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
    openxlsx::setRowHeights(wb, 1, rows = 1:nrow_x, heights = row_heights)
  }
  if (!is.null(col_widths)) {
    openxlsx::setColWidths(wb, 1, cols = 1:ncol_x, widths = col_widths)
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
      protect[!names(protect) %in% c("cols")]
    )

    do.call(openxlsx::protectWorksheet, protect_args)

    # unprotect
    openxlsx::addStyle(
      wb,
      sheet,
      style = unprotect(),
      rows = seq_len(nrow_x),
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
      x = x,
      wb = wb,
      sheet = sheet,
      style = style_head,
      has_header = has_header,
      data_start_row = data_start_row,
      nrow_x = nrow_x
    )
  }

  if (!is.null(style1)) {
    apply_row_style(
      x = x,
      wb = wb,
      sheet = sheet,
      style = style1,
      has_header = has_header,
      data_start_row = data_start_row,
      nrow_x = nrow_x
    )
  }

  if (!is.null(style2)) {
    apply_row_style(
      x = x,
      wb = wb,
      sheet = sheet,
      style = style2,
      has_header = has_header,
      data_start_row = data_start_row,
      nrow_x = nrow_x
    )
  }

  if (!is.null(style3)) {
    apply_row_style(
      x = x,
      wb = wb,
      sheet = sheet,
      style = style3,
      has_header = has_header,
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

    openxlsx::addWorksheet(wb, "options", visible = FALSE)
    openxlsx::writeData(wb, "options", x = validate_df, colNames = FALSE)

    for (j in unique(validate_df[[1]])) {

      i_rng <- range(which(validate_df[[1]] %in% j))
      excel_range <- paste0("'Options'!", "$B$", i_rng[1], ":", "$B$", i_rng[2])

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
    openxlsx::saveWorkbook(wb, file = file, overwrite = overwrite)
  } else {
    return(wb)
  }
}



#' Convert tidy-selection of columns to integer indexes
#'
#' @noRd
#' @importFrom rlang `!!`
#' @importFrom dplyr select
col_selection <- function(data, cols, invert = FALSE) {
  sel <- names(dplyr::select(data, !!cols))
  if (invert) {
    out <- which(!names(data) %in% sel)
  } else {
    out <- which(names(data) %in% sel)
  }
  out
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
expr_to_excel <- function(x, cols, has_header) {

  out_lang <- translate_traverse(x, cols, has_header)
  out <- deparse1(out_lang)
  out <- gsub("(?<![[:alpha:]])(?=[[:alpha:]]+[23]___TEMP_)", "$", out, perl = TRUE)
  out <- gsub("___TEMP_", "", out)
  out
}


#' Recursive function used to traverse expressions from top-down to break into
#' binary components (3 terms or fewer) that can be passed to
#' translate_options(), translate_missing(), and translate_equals()
#'
#' @noRd
translate_traverse <- function(x, cols, has_header) {

  if (is_expr_lowest(x)) {
    x <- expr_to_excel_helper(x, cols, has_header)
  } else {
    for (i in seq_len(length(x))) {
      if (is_expr_lowest(x[[i]])) {
        x[[i]] <- expr_to_excel_helper(x[[i]], cols, has_header)
      } else {
        x[[i]] <- translate_traverse(x[[i]], cols = cols, has_header = has_header)
      }
    }
  }

  and_i <- which(as.character(x) %in% "&")
  if (length(and_i) > 0) {
    x[[and_i]] <- str2lang("AND")
  }

  or_i <- which(as.character(x) %in% "|")
  if (length(or_i) > 0) {
    x[[or_i]] <- str2lang("OR")
  }

  x
}



#' @noRd
expr_to_excel_helper <- function(x, cols, has_header) {

  suffix <- ifelse(has_header, "3___TEMP_", "2___TEMP_")
  vars <- intersect(all.vars(x), cols)

  if (length(vars) > 0) {
    for (var_focal in vars) {
      var_i <- which(as.character(x) %in% var_focal)
      col_i <- which(cols %in% var_focal)
      var_replace <- str2lang(paste0(COLS_EXCEL[col_i], suffix))
      if (length(x) == 1) {
        x <- var_replace
      } else {
        x[[var_i]] <- var_replace
      }
    }
  }

  as_date_i <- which(as.character(x) %in% "as.Date")
  if (length(as_date_i) > 0) {
    for (i in as_date_i) {
      x[[i]] <- str2lang("DATEVALUE")
    }
  }

  x
}


#' Test whether an expression is binary (has at most 3 atomic terms)
#'
#' @param x A call returned by str2lang()
#'
#' @noRd
is_expr_lowest <- function(x) {
  all(lengths(as.list(x)) <= 1L) & length(x) <= 3L
}


#' @noRd
apply_row_style <- function(x,
                            wb,
                            sheet,
                            style,
                            has_header,
                            data_start_row,
                            nrow_x) {

  cols_style <- col_selection(x, style$cols)

  if (is_numeric(style$rows)) {
    # row-specific formatting (non-conditional)
    openxlsx::addStyle(
      wb,
      sheet,
      style = style$style,
      rows = eval(style$rows),
      cols = cols_style,
      gridExpand = TRUE,
      stack = TRUE
    )
  } else {
    # conditional formatting
    rule <- expr_to_excel(style$rows, names(x), has_header)
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
