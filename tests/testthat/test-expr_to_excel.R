context("expr_to_excel")

test_that("expr_to_excel works as expected", {

  # data for testing
  dat <- data.frame(
    x1 = 1:8,
    x2 = letters[1:8],
    x3 = c(0, NA, 0, 4, NA, 2, 0, NA),
    stringsAsFactors = FALSE
  )

  # test simple R expressions
  expect_equal(
    expr_to_excel(x1 == 8, dat),
    "$A2 == 8"
  )
  expect_equal(
    expr_to_excel(x1 > 4 | x2 == "e", dat),
    "OR($A2 > 4, $B2 == \"e\")"
  )
  expect_equal(
    expr_to_excel(x1 > 4 & x2 == "e" & x3 == 0, dat),
    "AND(AND($A2 > 4, $B2 == \"e\"), $C2 == 0)"
  )

  # test arg row_start
  expect_equal(
    expr_to_excel(x1 == 8, dat, row_start = 3),
    "$A3 == 8"
  )
})

