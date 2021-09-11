#' Read xls and xlsx worksheet
#'
#' @description
#' Wrapper to [`readxl::read_excel`] with minor changes to default settings:
#' - columns of dates with no time component have class "Date" rather than "POSIX"
#' - empty columns are read in as class "character" rather than "logical"
#' - the max number of rows used to guess column types is 10k rather than 1k
#'
#' @inheritParams readxl::read_excel
#'
#' @param simplify_dates Logical indicating whether to convert date columns
#'   lacking a time component to class "Date". By default [readxl::read_excel]
#'   reads columns containing dates or datetimes as class POSIX, even if there
#'   is no time component (i.e. in which case the times will all be "00:00:00").
#'   If `simplify_posix` is TRUE (the default), columns containing dates with no
#'   nonzero time values are converted to class "Date" using
#'   [lubridate::as_date].
#'
#' @param empty_cols_to_chr Logical indicating whether columns of class
#'   "logical" containing all missing values should be converted to class
#'   "character". If argument `col_types` is NULL (the default), columns
#'   containing all missing values are read in by readxl::read_excel as class
#'   "logical". If empty_cols_to_chr is TRUE (the default), such columns are
#'   converted to class "character".
#'
#' @importFrom lubridate as_date
#' @import readxl
#' @export qread
qread <- function(path,
                  sheet = NULL,
                  range = NULL,
                  col_names = TRUE,
                  col_types = NULL,
                  simplify_dates = TRUE,
                  empty_cols_to_chr = TRUE,
                  na = "",
                  trim_ws = TRUE,
                  skip = 0,
                  n_max = Inf,
                  guess_max = min(10000, n_max),
                  progress = FALSE,
                  .name_repair = "unique") {

  x <- readxl::read_excel(
    path = path,
    sheet = sheet,
    range = range,
    col_names = col_names,
    col_types = col_types,
    na = na,
    trim_ws = trim_ws,
    skip = skip,
    n_max = n_max,
    guess_max = guess_max,
    progress = progress,
    .name_repair = .name_repair
  )

  # convert POSIX cols with no times to class Date
  if (simplify_dates) {
    posix_without_times <- which(vapply(x, posix_without_times, FALSE, USE.NAMES = FALSE))
    for (j in posix_without_times) x[[j]] <- lubridate::as_date(x[[j]])
  }

  # convert logical cols with all missing values to character
  if (empty_cols_to_chr) {
    logic_all_na <- which(vapply(x, logical_all_na, FALSE, USE.NAMES = FALSE))
    for (j in logic_all_na) x[[j]] <- as.character(x[[j]])
  }

  # return
  x
}



#' Test whether vector is of class POSIX and doesn't contain any non-zero times
#' @noRd
posix_without_times <- function(x) {
  "POSIXct" %in% class(x) && all(format(x, "%H:%M:%S") %in% c("00:00:00", NA))
}


#' Test whether vector is of class logical and contains all NA
#' @noRd
logical_all_na <- function(x) {
  "logical" %in% class(x) && all(is.na(x))
}

