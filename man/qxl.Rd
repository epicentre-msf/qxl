% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/qxl.R
\name{qxl}
\alias{qxl}
\title{Quickly write a tidy data frame to xlsx, with options for customization}
\usage{
qxl(
  x,
  file = NULL,
  wb = openxlsx::createWorkbook(),
  sheet = NULL,
  header = NULL,
  style_head = qstyle(rows = 1, textDecoration = "bold"),
  hide_subhead = TRUE,
  style1 = NULL,
  style2 = NULL,
  style3 = NULL,
  style4 = NULL,
  style5 = NULL,
  group,
  group_style = qstyle(bgFill = "#ffcccb"),
  row_heights = NULL,
  col_widths = "auto",
  freeze_row = 1L,
  freeze_col = NULL,
  protect,
  validate = NULL,
  validate_cond = NULL,
  validate_cond_all = NULL,
  filter = FALSE,
  filter_cols = everything(),
  zoom = 120L,
  date_format = "yyyy-mm-dd",
  overwrite = TRUE
)
}
\arguments{
\item{x}{A data frame, or list of data frames}

\item{file}{Filename to write to. If \code{NULL} the resulting workbook is
returned as an openxlsx object of class "Workbook" rather than written to a
file.}

\item{wb}{An openxlsx workbook object to write to. Defaults to a fresh
workbook created with \code{\link[openxlsx:createWorkbook]{openxlsx::createWorkbook}}. Only need to update
when repeatedly calling \code{qxl()} to add worksheets to an existing workbook.}

\item{sheet}{Optional character vector of worksheet names. If \code{NULL} (the
default) and \code{x} is a named list of data frames, worksheet names are taken
from \code{names(x)}. Otherwise, names default to "Sheet1", "Sheet2", ...}

\item{header}{Optional column header. Defaults to \code{NULL} in which case column
names are taken directly from the data frame(s) in \code{x}, to create normal
single-row headers. Can alternatively pass a named character vector to set
custom names as the first row and a subheader with variable names as a
hidden second row.

\if{html}{\out{<div class="sourceCode">}}\preformatted{header = c(
  mpg = "Miles per US gallon",
  cyl = "Number of cylinders",
  disp = "Engine displacement (cubic in.)
)
}\if{html}{\out{</div>}}}

\item{style_head}{Style for the header row. Set with \code{\link[=qstyle]{qstyle()}}, or set to
\code{NULL} for no header styling. Defaults to bold text.}

\item{hide_subhead}{Logical indicating whether to hide the subheader (if
present). Defaults to TRUE.}

\item{style1, style2, style3, style4, style5}{Optional style to set using \code{\link[=qstyle]{qstyle()}}}

\item{group}{Optional vector of one or more column names used to create
alternating groupings of rows, with every other row grouping styled as per
argument \code{group_style}. See section \strong{Grouping rows}.}

\item{group_style}{Optional style to apply to alternating groupings of rows,
as specified using argument \code{groups}. Set using \code{\link[=qstyle]{qstyle()}}}

\item{row_heights}{Numeric vector of row heights (in Excel units). The vector
is recycled if shorter than the number of rows in \code{x}. Defaults to \code{NULL}
to use default row heights.}

\item{col_widths}{Vector of column widths (in Excel units). Can be numeric or
character, and may include keyword "auto" for automatic column sizing. The
vector is recycled if shorter than the number of columns in \code{x}. Defaults
to "auto".

Use named vector to give column widths for specific columns, where names
represent column names of \code{x} or the keyword ".default" to set a default
column width for all columns not otherwise specified. E.g.

\if{html}{\out{<div class="sourceCode">}}\preformatted{# specify widths for cols mpg and cyl, all others default to "auto"
col_widths <- c(mpg = 5, cyl = 10)

# specify widths for cols mpg and cyl, and explicit default for all others
col_widths <- c(mpg = 5, cyl = 10, .default = 7)
}\if{html}{\out{</div>}}}

\item{freeze_row}{Integer specifying a row to freeze at. Defaults to \code{1} to
add a freeze below the header row. Set to \code{0} or \code{NULL} to omit freezing.}

\item{freeze_col}{Integer specifying a column to freeze at. Defaults to
\code{NULL}. Set to \code{0} or \code{NULL} to omit freezing.}

\item{protect}{Optional function specifying how to protect worksheet
components from user modification. See function \code{\link{qprotect}}.}

\item{validate}{Optional specification of list-style data validation for one
or more columns. Can specify either as a list of vectors giving options for
one or more column in \code{x}, e.g.:

\if{html}{\out{<div class="sourceCode">}}\preformatted{list(
  var_x = c("Yes", "No"),
  var_y = c("Small", "Medium", "Large")
)
}\if{html}{\out{</div>}}

or as a data.frame where the first column gives column names and the
second column gives corresponding options, e.g.:

\if{html}{\out{<div class="sourceCode">}}\preformatted{data.frame(
  col = c("var_x", "var_x", "var_y", "var_y", "var_y"),
  val = c("Yes", "No", "Small", "Medium", "Large")
)
}\if{html}{\out{</div>}}

Validation options are written/appended to a hidden worksheet named
"valid_options".}

\item{validate_cond}{Optional specification of conditional list-style
validation, where the set of values to be allowed in a given column depends
on the corresponding value within one or more other columns (e.g. the
allowed values in column 'city' depend on the corresponding value in
columns 'country' and 'province'). Must be a data.frame with at least two
columns, where the first column(s) give the conditional entries (e.g.
'country', 'province') and the last column gives the corresponding allowed
entries (e.g. 'city') to be implemented as data validation. The column
names in \code{validate_cond} should match the relevant columns within \code{x}.

Note that in the current implementation validation is based on values in
the conditional column(s) of \code{x} at the time the workbook is written, and
will not update in real time if those values are later changed.}

\item{validate_cond_all}{Optional vector of value(s) to always allow,
independent of the value in the conditional column (e.g. "Unknown").}

\item{filter}{Logical indicating whether to add column filters.}

\item{filter_cols}{Tidy-selection specifying which columns to filter. Only
used if \code{filter} is \code{TRUE}. Defaults to \code{everything()} to select all
columns.}

\item{zoom}{Integer specifying initial zoom percentage. Defaults to 130.}

\item{date_format}{Excel format for date columns. Defaults to "yyyy-mm-dd".}

\item{overwrite}{Logical indicating whether to overwrite existing file.
Defaults to \code{TRUE}}
}
\value{
If argument \code{file} is not specified, returns an openxlsx workbook object.
Otherwise writes workbook to file with no return.
}
\description{
A wrapper to the \link[openxlsx:openxlsx]{openxlsx} package optimized for writing
tidy data frames. Includes arguments to quickly add customization like:
\itemize{
\item conditional formatting written as R expressions
\item data validation rules based on a tidy dictionary structure
\item column-specific worksheet protection
\item custom column names with original variable-names hidden in the row below
}
}
\section{Grouping rows}{

Given a dataset with multiple rows per group (e.g. repeated observations on a
given individual), it can sometimes be useful to uniquely stylize alternating
groups to allow for quick visual distinction of the rows belonging to any
given group.

Given one or more grouping columns specified using argument \code{groups}, the
\code{qxl} function arranges the rows of the resulting worksheet by group and then
applies the style \code{group_style} to the rows in every \emph{other} group, to create
an alternating pattern. The alternating pattern is achieved by first creating
a group index variable called \code{g} which is assigned a value of either \code{1} or
\code{0}: \code{1} for the 1st group, \code{0} for the 2nd, \code{1} for the 3rd, \code{0} for the
4th, etc. The style specified by \code{group_style} is then applied conditionally
to rows where \code{g == 0}. The grouping variable is written in column A, which
is hidden.
}

\examples{
library(datasets)
qxl(mtcars, file = tempfile(fileext = ".xlsx"))

}
