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
  if (is_quo(x)) {
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

  x
}


#' @noRd
expr_to_excel_helper <- function(x, cols, row_start) {

  # convert R variable name to Excel-style column references, with temporary
  # suffix to aid in later processing
  suffix <- paste0(row_start, "___TEMP_")
  vars <- intersect(all.vars(x), cols)

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

  # convert R functions/operators to Excel equivalent
  swap_fn(x)
}


#' Swap common R operators/functions with Excel equivalent
#'
#' @param x A call returned by str2lang()
#'
#' @noRd
swap_fn <- function(x) {

  x_chr <- as.character(x)

  if (length(x_chr) == 1 && x_chr %in% ROP2EXCEL) {
    # if single operator
    m <- match(x_chr, ROP2EXCEL)
    x <- str2lang(names(ROP2EXCEL)[m])
  } else if (any(x_chr %in% RFN2EXCEL)) {
    # else if contains translatable function
    m <- match(x_chr, RFN2EXCEL)
    for (i in which(!is.na(m))) {
      x[[i]] <- str2lang(names(RFN2EXCEL)[m[i]])
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

