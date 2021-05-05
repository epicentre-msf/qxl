#' Write workbook to an xlsx file
#'
#' @description
#' Wrapper to [`openxlsx::saveWorkbook`] to write an Excel workbook to file
#'
#' @inheritParams openxlsx::saveWorkbook
#'
#' @importFrom openxlsx saveWorkbook
#' @export qwrite
qwrite <- function(wb,
                   file,
                   overwrite = FALSE) {

  openxlsx::saveWorkbook(
    wb = wb,
    file = file,
    overwrite = overwrite
  )
}
