#' Protect one ore more columns of a worksheet
#'
#' Wrapper to [`openxlsx::protectWorksheet`] with additional argument `cols` to
#' enable protection to be limited to specific columns.
#'
#' @param cols Tidy-selection specifying the columns that protection should
#'   apply to. Defaults to [`dplyr::everything`] to select all columns.
#'
#' @inheritParams openxlsx::protectWorksheet
#'
#' @importFrom openxlsx protectWorksheet
#' @importFrom dplyr everything
#' @importFrom rlang enquo
#' @export qprotect
qprotect <- function(password = NULL,
                     cols = everything(),
                     protect = TRUE,
                     lockSelectingLockedCells = FALSE,
                     lockSelectingUnlockedCells = FALSE,
                     lockFormattingCells = FALSE,
                     lockFormattingColumns = FALSE,
                     lockFormattingRows = FALSE,
                     lockInsertingColumns = TRUE,
                     lockInsertingRows = TRUE,
                     lockInsertingHyperlinks = FALSE,
                     lockDeletingColumns = TRUE,
                     lockDeletingRows = TRUE,
                     lockSorting = FALSE,
                     lockAutoFilter = FALSE,
                     lockPivotTables = NULL,
                     lockObjects = NULL,
                     lockScenarios = NULL) {

  list(
    password = password,
    cols = rlang::enquo(cols),
    protect = protect,
    lockSelectingLockedCells = lockSelectingLockedCells,
    lockSelectingUnlockedCells = lockSelectingUnlockedCells,
    lockFormattingCells = lockFormattingCells,
    lockFormattingColumns = lockFormattingColumns,
    lockFormattingRows = lockFormattingRows,
    lockInsertingColumns = lockInsertingColumns,
    lockInsertingRows = lockInsertingRows,
    lockInsertingHyperlinks = lockInsertingHyperlinks,
    lockDeletingColumns = lockDeletingColumns,
    lockDeletingRows = lockDeletingRows,
    lockSorting = lockSorting,
    lockAutoFilter = lockAutoFilter,
    lockPivotTables = lockPivotTables,
    lockObjects = lockObjects,
    lockScenarios = lockScenarios
  )
}


#' @noRd
#' @importFrom openxlsx createStyle
unprotect <- function() {
  openxlsx::createStyle(locked = FALSE)
}

