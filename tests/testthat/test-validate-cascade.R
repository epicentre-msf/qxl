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

  # level 2 named ranges: opaque casc_N ids (split() sorts alphabetically:
  # Ontario=col2, Quebec=col3)
  expect_true(all(c("casc_2", "casc_3") %in% named_regions))

  # level 3 named ranges: lookup table occupies cols 4-5, so lvl3 groups start
  # at col 6 (split() order: Ontario__Ottawa, Ontario__Toronto,
  # Quebec__Montreal, Quebec__Quebec City)
  expect_true(all(c("casc_6", "casc_7", "casc_8", "casc_9") %in% named_regions))

  # no human-readable or sanitized names should appear
  expect_false(any(
    c("Ontario", "Quebec", "Ontario__Toronto", "Ontario__Ottawa", "Quebec__Quebec_City", "Quebec__Montreal") %in%
      named_regions
  ))

  # three data validations written to the main sheet (one per cascade column)
  sheet_idx <- which(wb$sheet_names == "Sheet1")
  dvs <- wb$worksheets[[sheet_idx]]$dataValidationsLst
  expect_length(dvs, 3L)

  # level 1 uses a direct range reference, not INDIRECT
  expect_true(grepl("valid_options_cascade", dvs[[1]]))
  expect_false(grepl("INDIRECT", dvs[[1]]))

  # levels 2 and 3 use INDIRECT with MATCH+INDEX lookup (not SUBSTITUTE)
  expect_true(grepl("INDIRECT.*INDEX.*MATCH", dvs[[2]]))
  expect_true(grepl("INDIRECT.*INDEX.*MATCH", dvs[[3]]))
  expect_false(grepl("SUBSTITUTE", dvs[[2]]))
  expect_false(grepl("SUBSTITUTE", dvs[[3]]))

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

  # each state gets its own named range at level 2 (split() alphabetical order:
  # California=col2, Florida=col3, New York=col4; lookup at cols 5-6)
  expect_true(all(c("casc_2", "casc_3", "casc_4") %in% named_regions))

  # each state+county combo gets its own named range at level 3 (split()
  # alphabetical order: California__LA County=col7, California__Orange
  # County=col8, Florida__Orange County=col9, New York__Albany County=col10,
  # New York__Orange County=col11)
  expect_true(all(
    c("casc_7", "casc_8", "casc_9", "casc_10", "casc_11") %in% named_regions
  ))

  # no human-readable or sanitized names
  expect_false(any(
    c("Orange_County", "New_York", "California", "Florida") %in% named_regions
  ))

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

  ca_oc <- nr_df$pos[nr_df$name == "casc_8"] # California__Orange County
  fl_oc <- nr_df$pos[nr_df$name == "casc_9"] # Florida__Orange County
  ny_oc <- nr_df$pos[nr_df$name == "casc_11"] # New York__Orange County

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


test_that("validate_cascade works correctly with non-ASCII (Arabic) values", {
  # values that differ only in Arabic characters must not collide, and the
  # dropdowns must resolve correctly despite the non-ASCII text
  cascade <- data.frame(
    country = c("Egypt", "Egypt", "Egypt"),
    city = c(
      "Cairo \u0627\u0644\u0642\u0627\u0647\u0631\u0629",
      "Alexandria \u0627\u0644\u0625\u0633\u0643\u0646\u062f\u0631\u064a\u0629",
      "Giza \u0627\u0644\u062c\u064a\u0632\u0629"
    ),
    district = c("district1", "district2", "district3"),
    stringsAsFactors = FALSE
  )

  df <- data.frame(
    country = "Egypt",
    city = NA_character_,
    district = NA_character_,
    stringsAsFactors = FALSE
  )

  wb <- qxl(df, validate_cascade = cascade)

  named_regions <- openxlsx::getNamedRegions(wb)

  # level 2: one group (all cities under Egypt) at col 2; lookup at cols 3-4
  expect_true("casc_2" %in% named_regions)

  # level 3: three distinct groups at cols 5-7 (split() alphabetical order:
  # Egypt__Alexandria..., Egypt__Cairo..., Egypt__Giza...)
  expect_true(all(c("casc_5", "casc_6", "casc_7") %in% named_regions))

  # exactly 4 casc_N ranges total (1 lvl2 + 3 lvl3), no spurious extras from
  # collision/stripping
  casc_names <- grep("^casc_", named_regions, value = TRUE)
  expect_length(casc_names, 4L)

  # formula uses MATCH+INDEX, not SUBSTITUTE
  sheet_idx <- which(wb$sheet_names == "Sheet1")
  dvs <- wb$worksheets[[sheet_idx]]$dataValidationsLst
  expect_false(grepl("SUBSTITUTE", dvs[[3]]))
  expect_true(grepl("INDIRECT.*INDEX.*MATCH", dvs[[3]]))
})


if (FALSE) {
  # test performance with larger dataset
  geo_nga <- obtdata::fetch_georef("HTI")$ref

  cascade <- geo_nga |>
    dplyr::filter(level == 4) |>
    distinct(adm1_name, adm2_name, adm3_name, adm4_name) |>
    arrange(adm1_name, adm2_name, adm3_name, adm4_name)

  df <- data.frame(
    id = 1:10,
    adm1_name = c("Sud", rep(NA, 9)),
    adm2_name = NA_character_,
    adm3_name = NA_character_,
    adm4_name = NA_character_,
    stringsAsFactors = FALSE
  )

  qxl::qxl(
    df,
    "~/desktop/test_cascade.xlsx",
    validate_cascade = cascade
  )
}
