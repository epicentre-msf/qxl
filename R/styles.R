#' Conditional cell styles
#'
#' @description
#' Wrappers to [`openxlsx::createStyle`] to create cell styles, with additional
#' arguments `rows` and `cols` to specify the rows and/or columns that the style
#' should apply to.
#'
#' @param rows Which rows the style should apply to. Can be set using
#'   either:\cr\cr
#'   __Integer rows indexes__: (e.g. `rows = c(2, 5, 6)`)\cr
#'   Note that in this case indexes represent Excel rows rather than R rows
#'   (i.e. the header is row 1).\cr\cr
#
#'   __An expression__: (e.g. `rows = cyl > 4`)\cr
#'   Given an expression the style is applied using conditional formatting, with
#'   the expression translated into its Excel formula equivalent.\cr
#'
#'   Expressions can optionally include a `.x` selector (e.g. `.x == 1`) to
#'   refer to multiple columns. See section __Using a .x selector__ below.\cr
#'
#'   Note that conditional formatting can update in real time if relevant data
#'   is changed within the workbook.
#'
#' @param cols Tidy-selection specifying the columns that the style should apply
#'   to. Defaults to [`dplyr::everything`] to select all columns.
#'
#' @section Using a .x selector:
#' An expression passed to the `rows` argument can optionally incorporate a `.x`
#' selector to refer to multiple columns within the worksheet.
#'
#' When a `.x` selector is used, each column specified in arguments `cols` is
#' independently swapped into the `.x` position of the expression, which is then
#' translated to the Excel formula equivalent and applied as conditional
#' formatting to the worksheet.
#'
#' For example, given the following `qstyle` specification with respect to the
#' [`mtcars`] dataset
#'
#' ```
#' qstyle(
#'   rows = .x == 1,
#'   cols = c(vs, am, carb),
#'   bgFill = "#FFC7CE"
#' )
#' ```
#' the style `bgFill = "#FFC7CE"` would be independently applied to any cell in
#' columns `vs`, `am`, or `carb` with a value of `1`.
#'
#' @inheritParams openxlsx::createStyle
#'
#' @importFrom openxlsx createStyle
#' @importFrom dplyr everything
#' @importFrom rlang enquo
#' @export qstyle
qstyle <- function(rows,
                   cols = everything(),
                   fontName = NULL,
                   fontSize = NULL,
                   fontColour = NULL,
                   border = NULL,
                   borderColour = getOption("openxlsx.borderColour", "black"),
                   borderStyle = getOption("openxlsx.borderStyle", "thin"),
                   bgFill = NULL,
                   fgFill = NULL,
                   halign = NULL,
                   valign = NULL,
                   textDecoration = NULL,
                   wrapText = FALSE,
                   textRotation = NULL,
                   indent = NULL,
                   locked = NULL,
                   hidden = NULL) {

  style <- openxlsx::createStyle(
    fontName = fontName,
    fontSize = fontSize,
    fontColour = fontColour,
    border = border,
    borderColour = borderColour,
    borderStyle = borderStyle,
    bgFill = bgFill,
    fgFill = fgFill,
    halign = halign,
    valign = valign,
    textDecoration = textDecoration,
    wrapText = wrapText,
    textRotation = textRotation,
    indent = indent,
    locked = locked,
    hidden = hidden
  )

  list(
    rows = rlang::enquo(rows),
    cols = rlang::enquo(cols),
    style = style
  )
}

