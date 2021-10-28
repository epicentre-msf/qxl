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

  # test argument col_widths
  df <- data.frame(x1 = 1:3, x2 = 1:3, x3 = 1:3)

  expect_silent(
    qxl(df, col_widths = c(x1 = 4, .default = 6))
  )

  expect_error(
    qxl(df, col_widths = c(x1 = 4, .default = 6, NONVALID = 6))
  )

  # test validation
  df <- data.frame(
    x1 = c("a", "a", "b", "c"),
    x2 = NA_character_
  )

  qxl(
    df,
    file = file_write,
    validate = list(x2 = c("Yes", "No", "Maybe"))
  )

  opts <- readxl::read_xlsx(
    file_write,
    sheet = "valid_options",
    col_names = c("cols", "allowed")
  )

  expect_equal(opts$allowed, c("Yes", "No", "Maybe"))

  # test conditional validation
  df <- data.frame(
    adm2 = c("AB", "SK", "SK", "AB"),
    adm3 = NA_character_
  )

  dict <- data.frame(
    adm2 = c("AB", "AB", "SK", "SK"),
    adm3 = c("Calgary", "Edmonton", "Regina", "Saskatoon")
  )

  qxl(
    df,
    file = file_write,
    validate_cond = dict
  )

  opts <- readxl::read_xlsx(
    file_write,
    sheet = "valid_options_cond",
    col_names = names(dict)
  )

  expect_identical(as.data.frame(opts), as.data.frame(dict))

  # test group style
  qxl(
    mtcars_tbl,
    file = file_write,
    group = "model"
  )

  expect_equal(
    names(readxl::read_xlsx(file_write)),
    c("g", names(mtcars_tbl))
  )

  expect_setequal(
    readxl::read_xlsx(file_write)$g,
    c(0, 1)
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

