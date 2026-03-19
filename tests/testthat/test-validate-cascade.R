context("validate_cascade")

test_that("validate_cascade works as expected", {
  cascade <- data.frame(
    adm1 = c(
      "Ontario",
      "Ontario",
      "Ontario",
      "Ontario",
      "Ontario",
      "Quebec",
      "Quebec",
      "Quebec",
      "Quebec",
      "Quebec"
    ),
    adm2 = c(
      "Toronto",
      "Toronto",
      "Toronto",
      "Ottawa",
      "Ottawa",
      "Quebec City",
      "Quebec City",
      "Montreal",
      "Montreal",
      "Montreal"
    ),
    adm3 = c(
      "Toronto 1",
      "Toronto 2",
      "Toronto 3",
      "Ottawa 1",
      "Ottawa 2",
      "Quebec City 1",
      "Quebec City 2",
      "Montreal 1",
      "Montreal 2",
      "Montreal 3"
    ),
    stringsAsFactors = FALSE
  )

  df <- data.frame(
    adm1 = c("Ontario", "Quebec"),
    adm2 = NA_character_,
    adm3 = NA_character_,
    stringsAsFactors = FALSE
  )

  wb <- qxl(df, validate_cascade = cascade)

  # hidden sheet created
  expect_true("valid_options_cascade" %in% wb$sheet_names)

  named_regions <- openxlsx::getNamedRegions(wb)

  # level 2 named ranges: keyed by adm1 only
  expect_true(all(c("Ontario", "Quebec") %in% named_regions))

  # level 3 named ranges: keyed by compound adm1__adm2 key to avoid collisions
  expect_true(all(
    c("Ontario__Toronto", "Ontario__Ottawa", "Quebec__Quebec_City", "Quebec__Montreal") %in%
      named_regions
  ))

  # simple adm2 names must NOT appear as standalone ranges at level 3
  expect_false(any(c("Toronto", "Ottawa", "Quebec_City", "Montreal") %in% named_regions))

  # three data validations written to the main sheet (one per cascade column)
  sheet_idx <- which(wb$sheet_names == "Sheet1")
  dvs <- wb$worksheets[[sheet_idx]]$dataValidationsLst
  expect_length(dvs, 3L)

  # level 1 uses a direct range reference, not INDIRECT
  expect_true(grepl("valid_options_cascade", dvs[[1]]))
  expect_false(grepl("INDIRECT", dvs[[1]]))

  # levels 2 and 3 use INDIRECT with SUBSTITUTE for space handling
  expect_true(grepl("INDIRECT.*SUBSTITUTE", dvs[[2]]))
  expect_true(grepl("INDIRECT.*SUBSTITUTE", dvs[[3]]))

  # level 2 formula references only adm1 (col A)
  expect_true(grepl("\\$A", dvs[[2]]))
  expect_false(grepl("\\$B", dvs[[2]]))

  # level 3 formula references both adm1 (col A) and adm2 (col B)
  expect_true(grepl("\\$A", dvs[[3]]))
  expect_true(grepl("\\$B", dvs[[3]]))
})


test_that("validate_cascade handles duplicate child values across parents", {
  # "Orange County" appears under three different states
  cascade <- data.frame(
    state = c("New York", "New York", "California", "California", "Florida"),
    county = c("Orange County", "Albany County", "Orange County", "LA County", "Orange County"),
    city = c("Newburgh", "Albany", "Anaheim", "Los Angeles", "Orlando"),
    stringsAsFactors = FALSE
  )

  df <- data.frame(
    state = c("New York", "California", "Florida"),
    county = NA_character_,
    city = NA_character_,
    stringsAsFactors = FALSE
  )

  wb <- qxl(df, validate_cascade = cascade)

  named_regions <- openxlsx::getNamedRegions(wb)

  # each state gets its own named range at level 2
  expect_true(all(c("New_York", "California", "Florida") %in% named_regions))

  # each state+county combo gets its own named range at level 3
  expect_true(all(
    c(
      "New_York__Orange_County",
      "New_York__Albany_County",
      "California__Orange_County",
      "California__LA_County",
      "Florida__Orange_County"
    ) %in%
      named_regions
  ))

  # "Orange_County" must NOT appear as a standalone (collision) range
  expect_false("Orange_County" %in% named_regions)

  # each Orange County range contains only the cities for that state
  nr_df <- data.frame(
    name = named_regions,
    sheet = attr(named_regions, "sheet"),
    pos = attr(named_regions, "position"),
    stringsAsFactors = FALSE
  )

  casc_sheet_idx <- which(wb$sheet_names == "valid_options_cascade")

  read_range <- function(wb, sheet_idx, position) {
    rc <- strsplit(position, ":")[[1]]
    cols <- match(gsub("[0-9]", "", rc), LETTERS)
    rows <- as.integer(gsub("[A-Z]", "", rc))
    as.character(openxlsx::readWorkbook(wb, sheet_idx, colNames = FALSE)[
      rows[1]:rows[2],
      cols[1]
    ])
  }

  ny_oc <- nr_df$pos[nr_df$name == "New_York__Orange_County"]
  ca_oc <- nr_df$pos[nr_df$name == "California__Orange_County"]
  fl_oc <- nr_df$pos[nr_df$name == "Florida__Orange_County"]

  expect_equal(read_range(wb, casc_sheet_idx, ny_oc), "Newburgh")
  expect_equal(read_range(wb, casc_sheet_idx, ca_oc), "Anaheim")
  expect_equal(read_range(wb, casc_sheet_idx, fl_oc), "Orlando")
})


test_that("validate_cascade input validation catches bad args", {
  df <- data.frame(adm1 = "Ontario", adm2 = NA_character_, stringsAsFactors = FALSE)

  # not a data frame
  expect_error(
    qxl(df, validate_cascade = list(adm1 = "Ontario", adm2 = "Toronto")),
    "`validate_cascade` must be a data frame"
  )

  # fewer than 2 columns
  expect_error(
    qxl(df, validate_cascade = data.frame(adm1 = "Ontario")),
    "at least 2 columns"
  )

  # column names not in x
  expect_error(
    qxl(df, validate_cascade = data.frame(foo = "a", bar = "b")),
    "column names in `validate_cascade` must also be column names in `x`"
  )
})


test_that("sanitize_range_name handles edge cases", {
  expect_equal(sanitize_range_name("Quebec City"), "Quebec_City")
  expect_equal(sanitize_range_name("foo bar baz"), "foo_bar_baz")

  # special chars stripped
  expect_equal(sanitize_range_name("foo!@#bar"), "foobar")

  # leading digit gets underscore prefix
  expect_equal(sanitize_range_name("1abc"), "_1abc")
})
