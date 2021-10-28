

# Excel columns
COLS_EXCEL <- c(
  LETTERS,
  apply(
    expand.grid(x1 = LETTERS, x2 = LETTERS),
    1,
    function(x) paste0(x[2], x[1])
  )
)

# R operators and Excel functions equivalents
ROP2EXCEL <- c(
  AND = "&",
  OR = "|",
  NOT = "!"
)


# R functions and Excel equivalents
# https://rforexcelusers.com/excel-r-function-formula-list/
RFN2EXCEL <- c(
  AND = "&",
  OR = "|",
  NOT = "!",
  ISBLANK = "is.na",
  DATEVALUE = "as.Date",
  ABS = "abs",
  LOG = "log",
  EXP = "exp",
  SQRT = "sqrt"
)


#' @importFrom rlang is_quosure
is_quo <- function(x) {
  out <- try(rlang::is_quosure(x), silent = TRUE)
  if ("try-error" %in% class(out)) out <- FALSE
  out
}


#' @noRd
#' @importFrom rlang eval_tidy
is_quo_numeric <- function(x) {
  out <- try(all(is.numeric(rlang::eval_tidy(x))), silent = TRUE)
  if ("try-error" %in% class(out)) out <- FALSE
  out
}


#' @noRd
#' @importFrom rlang eval_tidy
is_quo_character <- function(x) {
  out <- try(all(is.character(rlang::eval_tidy(x))), silent = TRUE)
  if ("try-error" %in% class(out)) out <- FALSE
  out
}


#' @noRd
#' @importFrom rlang is_quosure quo_get_expr
quo_to_lang <- function(x) {
  is_quo <- try(rlang::is_quosure(x), silent = TRUE)
  if ("try-error" %in% class(is_quo)) is_quo <- FALSE
  if (is_quo) {
    x <- rlang::quo_get_expr(x)
  }
  x
}


#' Convert tidy-selection of columns to integer indexes
#'
#' @noRd
#' @importFrom rlang `!!`
#' @importFrom dplyr select
col_selection <- function(data, cols, index = TRUE, invert = FALSE) {

  x <- names(dplyr::select(data, !!cols))

  if (invert) {
    x <- setdiff(names(data), x)
  }
  if (index) {
    x <- which(names(data) %in% x)
  }

  x
}


#' @noRd
paste_collapse <- function(x, quote = TRUE, collapse = ", ") {
  if (quote) {
    out <- paste(dQuote(x, q = FALSE), collapse = collapse)
  } else {
    out <- paste(x, collapse = collapse)
  }
  return(out)
}


#' @noRd
paste_collapse_c <- function(x, quote = TRUE, collapse = ", ") {
  paste0("c(", paste_collapse(x, quote = quote, collapse = collapse), ")")
}

