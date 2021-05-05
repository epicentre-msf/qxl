context("basic")

test_that("basic functionality works as expected", {

  library(datasets)
  file_write <- tempfile(fileext = ".xlsx")

  mtcars_tbl <- tibble::as_tibble(tibble::rownames_to_column(mtcars, "model"))

  # test simple write
  qxl(mtcars_tbl, file = file_write)
  expect_true(file.exists(file_write))
  expect_equal(mtcars_tbl, readxl::read_xlsx(file_write))

  # test return workbook
  x <- qxl(mtcars_tbl)
  expect_is(x, "Workbook")

  # test multiple sheets
  x <- qxl(mtcars_tbl, sheet = "Sheet1")
  x <- qxl(mtcars_tbl[1:4,], wb = x, sheet = "BBB")
  x <- qxl(mtcars_tbl[1:5,], wb = x, sheet = "CCC")
  expect_equal(x$sheet_names, c("Sheet1", "BBB", "CCC"))

  # test argument header
  qxl(
    mtcars_tbl,
    file = file_write,
    header = setNames(toupper(names(mtcars_tbl)), names(mtcars_tbl))
  )

  expect_equal(
    names(readxl::read_xlsx(file_write)),
    toupper(names(mtcars_tbl))
  )

  expect_equal(
    names(readxl::read_xlsx(file_write, skip = 1)),
    names(mtcars_tbl)
  )

  # test writing list of data frames
  mtcars_split <- split(mtcars_tbl, mtcars_tbl$cyl)
  names(mtcars_split) <- paste0("cyl", names(mtcars_split))

  x <- qxl(mtcars_split)
  expect_equal(x$sheet_names, names(mtcars_split))

  x <- qxl(mtcars_split, sheet = c("x1", "x2", "x3"))
  expect_equal(x$sheet_names, c("x1", "x2", "x3"))

  qxl(mtcars_split, file = file_write, overwrite = TRUE)
  expect_equal(mtcars_split[["cyl8"]], readxl::read_xlsx(file_write, "cyl8"))
})

