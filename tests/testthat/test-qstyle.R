context("qstyle")

test_that("qstyle works as expected", {
  # test basic
  x <- qstyle(1:5, halign = "center")
  expect_equal(x$style$halign, "center")
  expect_equal(rlang::eval_tidy(x$rows), 1:5)

  # test non-valid keyword to arg 'rows'
  expect_error(qstyle("blah", halign = "center"))

  dat <- data.frame(
    x1 = 1:8,
    x2 = letters[1:8],
    x3 = c(0, NA, 0, 4, NA, 2, 0, NA),
    stringsAsFactors = FALSE
  )

  ## static row paths (addStyle) -------------------------------------------

  # rows = "data" (default): style applies to data rows only, runs silently
  expect_silent(qxl(dat, style = qstyle(rows = "data", bgFill = "#fddbc7")))

  # rows = "all": style applies to header + data rows, runs silently
  expect_silent(qxl(dat, style = qstyle(rows = "all", bgFill = "#fddbc7")))

  # rows = integer indices: explicit row selection, runs silently
  expect_silent(qxl(dat, style = qstyle(rows = 2:5, bgFill = "#fddbc7")))

  # rows = "data" with 0-row data frame: should not error (issue #1)
  expect_silent(qxl(dat[0, ], style = qstyle(bgFill = "#fddbc7")))

  ## conditional formatting paths (expression rows) ------------------------

  # expression, no .x: rule applies to all columns, one range per column
  wb <- qxl(dat, style = qstyle(rows = x1 >= 4 & x2 == "f", bgFill = "#fddbc7"))
  expect_equal(
    names(wb$worksheets[[1]]$conditionalFormatting),
    c("A2:A9", "B2:B9", "C2:C9")
  )

  # expression, cols restriction: conditional formatting only on selected cols
  wb <- qxl(dat, style = qstyle(rows = x1 >= 4, cols = x1, bgFill = "#fddbc7"))
  expect_equal(
    names(wb$worksheets[[1]]$conditionalFormatting),
    "A2:A9"
  )

  # expression with .x selector: each selected col gets its own rule
  wb <- qxl(dat, style = qstyle(rows = .x > 2, cols = c(x1, x3), bgFill = "#fddbc7"))
  expect_equal(
    names(wb$worksheets[[1]]$conditionalFormatting),
    c("A2:A9", "C2:C9")
  )
})
