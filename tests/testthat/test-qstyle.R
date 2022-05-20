context("qstyle")

test_that("qstyle works as expected", {

  # test basic
  x <- qstyle(1:5, halign = "center")
  expect_equal(x$style$halign, "center")
  expect_equal(rlang::eval_tidy(x$rows), 1:5)

  # test non-valid keyword to arg 'rows'
  expect_error(qstyle("blah", halign = "center"))

  # test within qxl
  dat <- data.frame(
    x1 = 1:8,
    x2 = letters[1:8],
    x3 = c(0, NA, 0, 4, NA, 2, 0, NA),
    stringsAsFactors = FALSE
  )

  # create cond formatting which applies to all 3 columns (A, B, C) and all
  #  data rows (2:9)
  wb <- qxl(
    dat,
    style1 = qstyle(rows = x1 >= 4 & x2 == "f", bgFill = "#fddbc7")
  )

  expect_equal(
    names(wb$worksheets[[1]]$conditionalFormatting),
    c("A2:A9", "B2:B9", "C2:C9")
  )

  # test style applied to data.frame with 0 rows (issue #1)
  expect_silent(qxl(dat[0,], style1 = qstyle(bgFill = "#fddbc7")))

})
