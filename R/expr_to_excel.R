#' Translate an R expression into an Excel conditional formatting formula
#'
#' @param x An R-style expression
#' @param data A data frame that expression `x` relates to
#' @param row_start Integer reflecting the first row of data in the Excel sheet
#'   to be output. Defaults to 2.
#'
#' @return
#' A character string reflecting an Excel conditional formatting formula
#'
#' @examples
#' library(datasets)
#'
#' expr_to_excel(cyl > 4, mtcars)
#' expr_to_excel(cyl > 4 & hp < 200, mtcars)
#'
#' @importFrom rlang `!!`
#' @export expr_to_excel
expr_to_excel <- function(x, data, row_start = 2L) {

  # if expression passed as quosure, convert to lang
  is_quo <- try(rlang::is_quosure(x), silent = TRUE)
  if ("try-error" %in% class(is_quo)) is_quo <- FALSE
  if (is_quo) {
    x <- rlang::quo_get_expr(x)
  } else {
    x <- str2lang(deparse(substitute(x), width.cutoff = 500))
  }

  # translate to Excel
  x_lang <- translate_traverse(x, names(data), row_start)

  # convert to string
  x_chr <- deparse(x_lang, width.cutoff = 500)

  # add $ before relevant excel column (identified with suffix '___TEMP_')
  out <- gsub(
    sprintf("(?<![[:alpha:]])(?=[[:alpha:]]+%i___TEMP_)", row_start),
    "$",
    x_chr,
    perl = TRUE
  )

  # remove temp suffix '___TEMP_' and return
  gsub("___TEMP_", "", out)
}



#' Recursive function used to traverse expressions from top-down to break into
#' binary components (3 terms or fewer) that can be passed to
#' translate_options(), translate_missing(), and translate_equals()
#'
#' @noRd
translate_traverse <- function(x, cols, row_start) {

  if (is_expr_lowest(x)) {
    x <- expr_to_excel_helper(x, cols, row_start)
  } else {
    for (i in seq_len(length(x))) {
      if (is_expr_lowest(x[[i]])) {
        x[[i]] <- expr_to_excel_helper(x[[i]], cols, row_start)
      } else {
        x[[i]] <- translate_traverse(x[[i]], cols, row_start)
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
expr_to_excel_helper <- function(x, cols, row_start) {

  suffix <- paste0(row_start, "___TEMP_")
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

