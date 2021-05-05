#' Protect one ore more columns of a worksheet
#'
#' @description
#' Wrapper to [`openxlsx::protectWorksheet`] with additional argument `cols` to
#' enable protection to be limited to specific columns.
#'
#' In practice, protection is first applied to the entire worksheet, and then
#' subsequently columns not selected for protection (if any) are unlocked one by
#' one. This unlock step (when relevant) also requires a row specification,
#' which by default we limit to the range of the current data. Thus, in
#' 'unprotected' columns within protected worksheets, rows beyond the range of
#' the data will remain protected. As a hack to work around this, the user can
#' specify a 'buffer' of additional empty rows to unprotect within each
#' non-protected column (e.g. to allow further data entry).
#'
#' @param cols Tidy-selection specifying the columns that protection should
#'   apply to. Defaults to [`dplyr::everything`] to select all columns.
#'
#' @param row_buffer The number of additional rows (beyond the range of the
#'   current data) to unprotect within columns not specified in argument `cols`.
#'   See explanation in Description. Defaults to `0`.
#'
#' @inheritParams openxlsx::protectWorksheet
#'
#' @importFrom openxlsx protectWorksheet
#' @importFrom dplyr everything
#' @importFrom rlang enquo
#' @export qprotect
qprotect <- function(password = NULL,
                     cols = everything(),
                     row_buffer = 0L,
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
    row_buffer = row_buffer,
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

