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

  # test multiple styles
  wb <- qxl(
    mtcars_tbl,
    style1 = qstyle(
      mpg > 20,
      cols = mpg,
      bgFill = "#fddbc7"
    ),
    style2 = qstyle(
      disp > 300,
      cols = disp,
      bgFill = "#fddbc7"
    ),
    style3 = qstyle(
      drat > 3.2,
      cols = drat,
      bgFill = "#fddbc7"
    ),
    style4 = qstyle(
      wt > 3,
      cols = wt,
      bgFill = "#fddbc7"
    )
  )

  expect_length(wb$worksheets[[1]]$conditionalFormatting, 4L)

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

  # test validate_cond_all
  qxl(
    df,
    file = file_write,
    validate_cond = dict,
    validate_cond_all = c("<Missing>", "<Unknown>")
  )

  opts <- readxl::read_xlsx(
    file_write,
    sheet = "valid_options_cond",
    col_names = names(dict)
  )

  dict_expect <- data.frame(
    adm2 = c(
      "AB", "AB", "AB", "AB",
      "SK", "SK", "SK", "SK"
    ),
    adm3 = c(
      "Calgary", "Edmonton", "<Missing>", "<Unknown>",
      "Regina", "Saskatoon", "<Missing>", "<Unknown>"
    )
  )

  expect_identical(as.data.frame(opts), as.data.frame(dict_expect))

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

  # test sheet name truncation
  sheets <- list(
    "this_sheet_name_is_too_long_for_excel" = data.frame(x = 1, y = 2),
    "seriously_what_an_absurd_length_for_a_sheet_name" = data.frame(x = 1, y = 2),
    "why_do_I_keep_giving_these_sheets_such_long_names" = data.frame(x = 1, y = 2)
  )

  x <- suppressWarnings(qxl::qxl(sheets))
  expect_equal(x$sheet_names, substr(names(sheets), 1, 29))

  # test sheet name truncation with deduplication
  sheets <- list(
    "this_sheet_name_is_far_too_long_for_excel_1" = data.frame(x = 1, y = 2),
    "this_sheet_name_is_far_too_long_for_excel_2" = data.frame(x = 1, y = 2),
    "this_sheet_name_is_far_too_long_for_excel_3" = data.frame(x = 1, y = 2)
  )

  x <- suppressWarnings(qxl::qxl(sheets))

  expect_true(all(nchar(x$sheet_names)) < 32)
  expect_length(unique(x$sheet_names), 3L)

  # test writing columns of chron class 'times' (issue #5)
  df <- data.frame(
    x1 = 1:2,
    x2 = structure(c(0.1, 0.3), format = "h:m:s", class = "times")
  )

  qxl(df, file = file_write)

  x <- qread(file_write)
  expect_equal(x$x2, as.character(df$x2))

})

