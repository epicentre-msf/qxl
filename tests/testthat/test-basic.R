context("basic")

test_that("basic functionality works as expected", {

  library(datasets)
  file_write <- tempfile(fileext = ".xlsx")

  mtcars_tbl <- tibble::as_tibble(tibble::rownames_to_column(mtcars, "model"))

  # test simple write
  qxl(mtcars_tbl, file = file_write)
  expect_true(file.exists(file_write))
  expect_equal(mtcars_tbl, readxl::read_xlsx(file_write))

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
})

