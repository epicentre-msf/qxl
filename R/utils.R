
#' @noRd
COLS_EXCEL <- c(
  LETTERS,
  apply(
    expand.grid(x1 = LETTERS, x2 = LETTERS),
    1,
    function(x) paste0(x[2], x[1])
  )
)


#' @noRd
is_numeric <- function(x) {
  out <- try(all(is.numeric(eval(x))), silent = TRUE)
  if ("try-error" %in% class(out)) out <- FALSE
  out
}

