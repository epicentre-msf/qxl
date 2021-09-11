context("qread")

test_that("qread works as expected", {

  x <- data.frame(
    x1 = 1:4,
    x2 = rep(NA, 4),
    x3 = letters[1:4],
    x4 = as.Date(c("2021-05-01", "2021-05-02", "2021-04-26", NA)),
    x5 = as.POSIXct(c("2021-05-01 13:05:23", NA, "2021-04-26 21:41:08", NA)),
    stringsAsFactors = FALSE
  )

  file_write <- tempfile(fileext = ".xlsx")
  qxl(x, file = file_write)

  # test default column types
  y <- qread(file_write)
  expect_s3_class(y, "tbl_df")
  expect_type(y$x1, "double")
  expect_type(y$x2, "character")
  expect_type(y$x3, "character")
  expect_s3_class(y$x4, "Date")
  expect_s3_class(y$x5, "POSIXct")

  # test argument simplify_dates
  expect_s3_class(qread(file_write, simplify_dates = FALSE)$x4, "POSIXct")

  # test argument simplify_dates
  expect_type(qread(file_write, empty_cols_to_chr = FALSE)$x2, "logical")

})

