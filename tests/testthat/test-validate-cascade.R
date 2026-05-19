context("validate_cascade")


# helper: read a single column block from the hidden cascade sheet
read_casc_col <- function(wb, col_letter, n) {
  sheet_idx <- which(wb$sheet_names == "valid_options_cascade")
  df <- openxlsx::readWorkbook(wb, sheet_idx, colNames = FALSE)
  col_idx <- match(col_letter, qxl:::int_to_excel_col(seq_len(ncol(df))))
  as.character(df[seq_len(n), col_idx])
}


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

  # hidden sheet created, no named regions used
  expect_true("valid_options_cascade" %in% wb$sheet_names)
  expect_length(openxlsx::getNamedRegions(wb), 0L)

  # data validations: one per cascade column
  sheet_idx <- which(wb$sheet_names == "Sheet1")
  dvs <- wb$worksheets[[sheet_idx]]$dataValidationsLst
  expect_length(dvs, 3L)

  # level 1: direct range reference, no OFFSET/INDIRECT
  expect_true(grepl("valid_options_cascade", dvs[[1]]))
  expect_false(grepl("OFFSET", dvs[[1]]))
  expect_false(grepl("INDIRECT", dvs[[1]]))

  # levels 2 and 3: OFFSET + MATCH + COUNTIF
  expect_true(grepl("OFFSET.*MATCH.*COUNTIF", dvs[[2]]))
  expect_true(grepl("OFFSET.*MATCH.*COUNTIF", dvs[[3]]))

  # level 2 key_expr references only adm1 (col A); level 3 concatenates A and B
  expect_false(grepl("CONCATENATE", dvs[[2]]))
  expect_true(grepl("MATCH\\(\\$A\\d", dvs[[2]]))
  expect_true(grepl("CONCATENATE\\(\\$A\\d+,\"__\",\\$B\\d", dvs[[3]]))

  # level 1 lookup: alphabetical "Ontario", "Quebec"
  expect_equal(read_casc_col(wb, "A", 2L), c("Ontario", "Quebec"))

  # level 2 lookup keys + values, sorted by key
  expect_equal(
    read_casc_col(wb, "B", 4L),
    c("Ontario", "Ontario", "Quebec", "Quebec")
  )
  expect_equal(
    read_casc_col(wb, "C", 4L),
    c("Ottawa", "Toronto", "Montreal", "Quebec City")
  )

  # level 3 lookup: 10 unique (compound-key, value) pairs, sorted by (key, value)
  expect_equal(
    read_casc_col(wb, "D", 10L),
    c(
      rep("Ontario__Ottawa", 2L),
      rep("Ontario__Toronto", 3L),
      rep("Quebec__Montreal", 3L),
      rep("Quebec__Quebec City", 2L)
    )
  )
  expect_equal(
    read_casc_col(wb, "E", 10L),
    c(
      "Ottawa 1",
      "Ottawa 2",
      "Toronto 1",
      "Toronto 2",
      "Toronto 3",
      "Montreal 1",
      "Montreal 2",
      "Montreal 3",
      "Quebec City 1",
      "Quebec City 2"
    )
  )
})


test_that("validate_cascade handles duplicate child values across parents", {
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

  # no named regions — pure formula-based lookup
  expect_length(openxlsx::getNamedRegions(wb), 0L)

  # the compound-key column at level 3 has one entry per (state, county) pair
  # in sorted order, so "Orange County" appears three times under different
  # state prefixes and resolves to a different city each time
  casc_idx <- which(wb$sheet_names == "valid_options_cascade")
  casc <- openxlsx::readWorkbook(wb, casc_idx, colNames = FALSE)

  # find the level-3 key column: it's the second 2-col block after level 1.
  # level 1 = col 1, level 2 keys/vals = cols 2-3, level 3 keys/vals = cols 4-5
  keys3 <- as.character(casc[!is.na(casc[[4]]), 4])
  vals3 <- as.character(casc[!is.na(casc[[5]]), 5])
  lookup <- setNames(vals3, keys3)

  expect_equal(lookup[["California__Orange County"]], "Anaheim")
  expect_equal(lookup[["Florida__Orange County"]], "Orlando")
  expect_equal(lookup[["New York__Orange County"]], "Newburgh")
})


test_that("validate_cascade supports long-format input with trailing NAs", {
  # Long format: some adm3 levels exist with no adm4 children; some adm2 with
  # no adm3 children. A row contributes a value at level i only when cols 1..i
  # are all non-NA.
  cascade <- data.frame(
    adm1 = c("ON", "ON", "ON", "ON", "QC", "QC", "QC"),
    adm2 = c(NA, "Toronto", "Toronto", "Ottawa", NA, "Montreal", "Montreal"),
    adm3 = c(NA, NA, "NorthYork", NA, NA, NA, "Plateau"),
    adm4 = c(NA, NA, NA, NA, NA, NA, NA),
    stringsAsFactors = FALSE
  )

  df <- data.frame(
    adm1 = c("ON", "QC"),
    adm2 = NA_character_,
    adm3 = NA_character_,
    adm4 = NA_character_,
    stringsAsFactors = FALSE
  )

  wb <- qxl(df, validate_cascade = cascade)

  # level 1: both adm1 values still appear (the bare "ON" / "QC" rows contribute)
  expect_equal(read_casc_col(wb, "A", 2L), c("ON", "QC"))

  # level 2: only rows with non-NA adm2 contribute. ON has Ottawa + Toronto,
  # QC has Montreal. Bare "ON" / "QC" rows are skipped.
  expect_equal(
    read_casc_col(wb, "B", 3L),
    c("ON", "ON", "QC")
  )
  expect_equal(
    read_casc_col(wb, "C", 3L),
    c("Ottawa", "Toronto", "Montreal")
  )

  # level 3: only rows with non-NA adm3 contribute
  expect_equal(
    read_casc_col(wb, "D", 2L),
    c("ON__Toronto", "QC__Montreal")
  )
  expect_equal(read_casc_col(wb, "E", 2L), c("NorthYork", "Plateau"))

  # level 4: no rows have non-NA adm4, so no validation for adm4
  sheet_idx <- which(wb$sheet_names == "Sheet1")
  dvs <- wb$worksheets[[sheet_idx]]$dataValidationsLst
  expect_length(dvs, 3L) # adm1, adm2, adm3 only
})


test_that("validate_cascade warns on gap rows (NA followed by non-NA)", {
  # gap: adm2 NA but adm3 set — likely a data error
  cascade <- data.frame(
    adm1 = c("ON", "ON"),
    adm2 = c("Toronto", NA),
    adm3 = c("NorthYork", "Phantom"),
    stringsAsFactors = FALSE
  )

  df <- data.frame(
    adm1 = "ON",
    adm2 = NA_character_,
    adm3 = NA_character_,
    stringsAsFactors = FALSE
  )

  expect_warning(
    qxl(df, validate_cascade = cascade),
    "gap"
  )
})


test_that("validate_cascade detects separator collisions", {
  # collision: ("Region__A", "X") and ("Region", "A__X") both produce
  # compound key "Region__A__X" — MATCH would be ambiguous in Excel
  cascade_collide <- data.frame(
    adm1 = c("Region__A", "Region"),
    adm2 = c("X", "A__X"),
    adm3 = c("child1", "child2"),
    stringsAsFactors = FALSE
  )

  df <- data.frame(
    adm1 = c("Region__A", "Region"),
    adm2 = NA_character_,
    adm3 = NA_character_,
    stringsAsFactors = FALSE
  )

  expect_error(
    qxl(df, validate_cascade = cascade_collide),
    "Separator collision.*Region__A__X"
  )

  # but a single "__" inside a value, with no collision, is fine
  cascade_safe <- data.frame(
    adm1 = c("Region__A", "Region__B"),
    adm2 = c("X", "Y"),
    adm3 = c("child1", "child2"),
    stringsAsFactors = FALSE
  )
  df_safe <- data.frame(
    adm1 = c("Region__A", "Region__B"),
    adm2 = NA_character_,
    adm3 = NA_character_,
    stringsAsFactors = FALSE
  )
  expect_error(qxl(df_safe, validate_cascade = cascade_safe), NA)
})


test_that("validate_cascade input validation catches bad args", {
  df <- data.frame(adm1 = "Ontario", adm2 = NA_character_, stringsAsFactors = FALSE)

  expect_error(
    qxl(df, validate_cascade = list(adm1 = "Ontario", adm2 = "Toronto")),
    "`validate_cascade` must be a data frame"
  )

  expect_error(
    qxl(df, validate_cascade = data.frame(adm1 = "Ontario")),
    "at least 2 columns"
  )

  expect_error(
    qxl(df, validate_cascade = data.frame(foo = "a", bar = "b")),
    "column names in `validate_cascade` must also be column names in `x`"
  )
})


test_that("validate_cascade works correctly with non-ASCII (Arabic) values", {
  cascade <- data.frame(
    country = c("Egypt", "Egypt", "Egypt"),
    city = c(
      "Cairo القاهرة",
      "Alexandria الإسكندرية",
      "Giza الجيزة"
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

  # level 3 lookup keys preserve the Arabic city names verbatim in the compound key
  casc_idx <- which(wb$sheet_names == "valid_options_cascade")
  casc <- openxlsx::readWorkbook(wb, casc_idx, colNames = FALSE)
  keys3 <- as.character(casc[!is.na(casc[[4]]), 4])
  vals3 <- as.character(casc[!is.na(casc[[5]]), 5])

  expect_length(keys3, 3L)
  expect_true(all(grepl("^Egypt__", keys3)))
  expect_equal(sort(vals3), c("district1", "district2", "district3"))

  # formula uses OFFSET+MATCH+COUNTIF
  sheet_idx <- which(wb$sheet_names == "Sheet1")
  dvs <- wb$worksheets[[sheet_idx]]$dataValidationsLst
  expect_true(grepl("OFFSET.*MATCH.*COUNTIF", dvs[[3]]))
})


test_that("validate_extra_rows extends cascade validation to extra rows", {
  cascade <- data.frame(
    adm1 = c("ON", "ON", "QC"),
    adm2 = c("Toronto", "Ottawa", "Montreal"),
    stringsAsFactors = FALSE
  )
  df <- data.frame(
    adm1 = c("ON", "QC"),
    adm2 = NA_character_,
    stringsAsFactors = FALSE
  )

  # helper: extract the last row in a data validation's sqref attribute
  dv_end_row <- function(dv) {
    m <- regmatches(dv, regexpr("[A-Z]+\\d+:[A-Z]+(\\d+)", dv))
    as.integer(sub(".*:[A-Z]+", "", m))
  }

  sheet_idx <- function(wb) which(wb$sheet_names == "Sheet1")

  # default: 1000 extra rows past data (data ends at row 3 -> validation ends at 1003)
  wb_def <- qxl(df, validate_cascade = cascade)
  dvs <- wb_def$worksheets[[sheet_idx(wb_def)]]$dataValidationsLst
  expect_equal(dv_end_row(dvs[[1]]), 3L + 1000L)
  expect_equal(dv_end_row(dvs[[2]]), 3L + 1000L)

  # 0 extra rows: matches old behavior (data range only)
  wb_zero <- qxl(df, validate_cascade = cascade, validate_extra_rows = 0L)
  dvs <- wb_zero$worksheets[[sheet_idx(wb_zero)]]$dataValidationsLst
  expect_equal(dv_end_row(dvs[[1]]), 3L) # data ends at row 3
  expect_equal(dv_end_row(dvs[[2]]), 3L)

  # Inf: extends to Excel's row limit (1,048,576)
  wb_inf <- qxl(df, validate_cascade = cascade, validate_extra_rows = Inf)
  dvs <- wb_inf$worksheets[[sheet_idx(wb_inf)]]$dataValidationsLst
  expect_equal(dv_end_row(dvs[[1]]), 1048576L)
  expect_equal(dv_end_row(dvs[[2]]), 1048576L)

  # explicit small value
  wb_50 <- qxl(df, validate_cascade = cascade, validate_extra_rows = 50L)
  dvs <- wb_50$worksheets[[sheet_idx(wb_50)]]$dataValidationsLst
  expect_equal(dv_end_row(dvs[[1]]), 3L + 50L)
  expect_equal(dv_end_row(dvs[[2]]), 3L + 50L)
})


test_that("validate_extra_rows extends simple validate dropdowns too", {
  df <- data.frame(
    x = c("a", "b"),
    stringsAsFactors = FALSE
  )

  wb <- qxl(df, validate = list(x = c("a", "b", "c")))
  dvs <- wb$worksheets[[which(wb$sheet_names == "Sheet1")]]$dataValidationsLst
  m <- regmatches(dvs[[1]], regexpr("[A-Z]+\\d+:[A-Z]+\\d+", dvs[[1]]))
  # data ends at row 3 (header row + 2 data rows); +1000 padding -> 1003
  expect_equal(sub(".*:[A-Z]+", "", m), as.character(3L + 1000L))
})


test_that("validate_cascade handles deep cascades without column-letter overflow", {
  # Reproducer for the original bug: the old design wrote one column per
  # ancestor-key group on the hidden sheet, so a 4-level cascade with hundreds
  # of level-4 groups overflowed COLS_EXCEL (capped at ZZ) and produced invalid
  # references. The new design uses only 2 columns per level, so this asserts
  # that property plus formula integrity on a moderately deep cascade.
  cascade <- expand.grid(
    adm1 = paste0("A", 1:3),
    adm2 = paste0("B", 1:4),
    adm3 = paste0("C", 1:5),
    adm4 = paste0("D", 1:6),
    KEEP.OUT.ATTRS = FALSE,
    stringsAsFactors = FALSE
  )
  # 3*4*5*6 = 360 rows; 3 adm1, 12 adm1__adm2, 60 adm1__adm2__adm3 groups

  dat <- data.frame(
    id = 1:3,
    adm1 = NA_character_,
    adm2 = NA_character_,
    adm3 = NA_character_,
    adm4 = NA_character_,
    stringsAsFactors = FALSE
  )

  wb <- qxl(dat, validate_cascade = cascade)

  sheet_idx <- which(wb$sheet_names == "Sheet1")
  dvs <- wb$worksheets[[sheet_idx]]$dataValidationsLst
  expect_length(dvs, 4L)

  # every data-validation formula resolves to valid Excel references (no "NA"s)
  for (dv in dvs) {
    expect_false(grepl(">NA<|\\$NA\\$", dv))
  }

  # hidden sheet uses just 1 + 2*(n-1) = 7 columns regardless of group counts
  casc_idx <- which(wb$sheet_names == "valid_options_cascade")
  casc <- openxlsx::readWorkbook(wb, casc_idx, colNames = FALSE)
  expect_equal(ncol(casc), 7L)
  expect_length(openxlsx::getNamedRegions(wb), 0L)

  # spot-check lookup-table sizes: level 3 keys = 12 distinct (adm1, adm2)
  # pairs each repeated 5 times for the 5 adm3 values = 60 rows
  expect_equal(sum(!is.na(casc[[4]])), 60L)
  # level 4 keys = 60 distinct (adm1, adm2, adm3) triples each repeated 6
  # times for the 6 adm4 values = 360 rows
  expect_equal(sum(!is.na(casc[[6]])), 360L)
})
