context("expr_to_excel")

test_that("expr_to_excel works as expected", {

  # data for testing
  dat <- data.frame(
    x1 = 1:6,
    x2 = letters[1:6],
    x3 = c(0, NA, 0, 4, NA, NA),
    x4 = as.Date(c(18721, NA, 18700, 18698, NA, 18700), origin = "1970-01-01"),
    stringsAsFactors = FALSE
  )

  # test simple R expressions
  expect_equal(
    expr_to_excel(x1 == 4, dat),
    "$A2 == 4"
  )
  expect_equal(
    expr_to_excel(x1 > 4 | x2 == "e", dat),
    "OR($A2 > 4, $B2 == \"e\")"
  )
  expect_equal(
    expr_to_excel(x1 > 4 & x2 == "e" & !is.na(x3), dat),
    "AND(AND($A2 > 4, $B2 == \"e\"), NOT(ISBLANK($C2)))"
  )

  # test arg row_start
  expect_equal(
    expr_to_excel(x1 == 4, dat, row_start = 3),
    "$A3 == 4"
  )
})

