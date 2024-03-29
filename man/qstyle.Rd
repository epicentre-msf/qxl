% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/styles.R
\name{qstyle}
\alias{qstyle}
\title{Conditional cell styles}
\usage{
qstyle(
  rows = "data",
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
  hidden = NULL
)
}
\arguments{
\item{rows}{Which rows the style should apply to. Can be set using
either:\cr\cr
\strong{Keyword}: (e.g. \code{rows = "data"} or \code{rows = "all"})\cr
Keyword "data" (the default) applies a style to all data rows (excludes the
header), whereas keyword "all" applies a style to all rows (header and
data)\cr\cr

\strong{Integer rows indexes}: (e.g. \code{rows = c(2, 5, 6)})\cr
Note that in this case indexes represent Excel rows rather than R rows
(i.e. the header is row 1).\cr\cr
\strong{An expression}: (e.g. \code{rows = cyl > 4})\cr
Given an expression the style is applied using conditional formatting, with
the expression translated into its Excel formula equivalent.\cr

Expressions can optionally include a \code{.x} selector (e.g. \code{.x == 1}) to
refer to multiple columns. See section \strong{Using a .x selector} below.\cr

Note that conditional formatting can update in real time if relevant data
is changed within the workbook.}

\item{cols}{Tidy-selection specifying the columns that the style should apply
to. Defaults to \code{\link[dplyr:reexports]{dplyr::everything}} to select all columns.}

\item{fontName}{A name of a font. Note the font name is not validated. If fontName is NULL,
the workbook base font is used. (Defaults to Calibri)}

\item{fontSize}{Font size. A numeric greater than 0.
If fontSize is NULL, the workbook base font size is used. (Defaults to 11)}

\item{fontColour}{Colour of text in cell.  A valid hex colour beginning with "#"
or one of colours(). If fontColour is NULL, the workbook base font colours is used.
(Defaults to black)}

\item{border}{Cell border. A vector of "top", "bottom", "left", "right" or a single string).
\itemize{
\item{\strong{"top"}}{ Top border}
\item{\strong{bottom}}{ Bottom border}
\item{\strong{left}}{ Left border}
\item{\strong{right}}{ Right border}
\item{\strong{TopBottom} or \strong{c("top", "bottom")}}{ Top and bottom border}
\item{\strong{LeftRight} or \strong{c("left", "right")}}{ Left and right border}
\item{\strong{TopLeftRight} or \strong{c("top", "left", "right")}}{ Top, Left and right border}
\item{\strong{TopBottomLeftRight} or \strong{c("top", "bottom", "left", "right")}}{ All borders}
}}

\item{borderColour}{Colour of cell border vector the same length as the number of sides specified in "border"
A valid colour (belonging to colours()) or a valid hex colour beginning with "#"}

\item{borderStyle}{Border line style vector the same length as the number of sides specified in "border"
\itemize{
\item{\strong{none}}{ No Border}
\item{\strong{thin}}{ thin border}
\item{\strong{medium}}{ medium border}
\item{\strong{dashed}}{ dashed border}
\item{\strong{dotted}}{ dotted border}
\item{\strong{thick}}{ thick border}
\item{\strong{double}}{ double line border}
\item{\strong{hair}}{ Hairline border}
\item{\strong{mediumDashed}}{ medium weight dashed border}
\item{\strong{dashDot}}{ dash-dot border}
\item{\strong{mediumDashDot}}{ medium weight dash-dot border}
\item{\strong{dashDotDot}}{ dash-dot-dot border}
\item{\strong{mediumDashDotDot}}{ medium weight dash-dot-dot border}
\item{\strong{slantDashDot}}{ slanted dash-dot border}
}}

\item{bgFill}{Cell background fill colour.
A valid colour (belonging to colours()) or a valid hex colour beginning with "#".
--  \strong{Use for conditional formatting styles only.}}

\item{fgFill}{Cell foreground fill colour.
A valid colour (belonging to colours()) or a valid hex colour beginning with "#"}

\item{halign}{Horizontal alignment of cell contents
\itemize{
\item{\strong{left}}{ Left horizontal align cell contents}
\item{\strong{right}}{ Right horizontal align cell contents}
\item{\strong{center}}{ Center horizontal align cell contents}
\item{\strong{justify}}{ Justify horizontal align cell contents}
}}

\item{valign}{A name
Vertical alignment of cell contents
\itemize{
\item{\strong{top}}{ Top vertical align cell contents}
\item{\strong{center}}{ Center vertical align cell contents}
\item{\strong{bottom}}{ Bottom vertical align cell contents}
}}

\item{textDecoration}{Text styling.
\itemize{
\item{\strong{bold}}{ Bold cell contents}
\item{\strong{strikeout}}{ Strikeout cell contents}
\item{\strong{italic}}{ Italicise cell contents}
\item{\strong{underline}}{ Underline cell contents}
\item{\strong{underline2}}{ Double underline cell contents}
\item{\strong{accounting}}{ Single accounting underline cell contents}
\item{\strong{accounting2}}{ Double accounting underline cell contents}
}}

\item{wrapText}{Logical. If \code{TRUE} cell contents will wrap to fit in column.}

\item{textRotation}{Rotation of text in degrees. 255 for vertical text.}

\item{indent}{Horizontal indentation of cell contents.}

\item{locked}{Whether cell contents are locked (if worksheet protection is turned on)}

\item{hidden}{Whether the formula of the cell contents will be hidden (if worksheet protection is turned on)}
}
\description{
Wrapper to \code{\link[openxlsx:createStyle]{openxlsx::createStyle}} to create cell styles, with additional
arguments \code{rows} and \code{cols} to specify the rows and/or columns that the style
should apply to.
}
\section{Using a .x selector}{

An expression passed to the \code{rows} argument can optionally incorporate a \code{.x}
selector to refer to multiple columns within the worksheet.

When a \code{.x} selector is used, each column specified in arguments \code{cols} is
independently swapped into the \code{.x} position of the expression, which is then
translated to the Excel formula equivalent and applied as conditional
formatting to the worksheet.

For example, given the following \code{qstyle} specification with respect to the
\code{\link{mtcars}} dataset

\if{html}{\out{<div class="sourceCode">}}\preformatted{qstyle(
  rows = .x == 1,
  cols = c(vs, am, carb),
  bgFill = "#FFC7CE"
)
}\if{html}{\out{</div>}}

the style \code{bgFill = "#FFC7CE"} would be independently applied to any cell in
columns \code{vs}, \code{am}, or \code{carb} with a value of \code{1}.
}

\examples{
# apply style halign = "center" to all data rows (by default rows = "data")
qstyle(halign = "center")

# apply style halign = "center" to all rows including header
qstyle(rows = "all", halign = "center")

# apply style halign = "center" to Excel rows 2:10
qstyle(rows = 2:10, halign = "center")

# apply conditional formatting to rows where cyl == 8 & mpg > 16
qstyle(cyl == 8 & mpg > 16, fgFill = "#fddbc7", textDecoration = "bold")

}
