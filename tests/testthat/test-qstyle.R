context("qstyle")

test_that("qstyle works as expected", {

  # data for testing
  dat <- data.frame(
    x1 = 1:8,
    x2 = letters[1:8],
    x3 = c(0, NA, 0, 4, NA, 2, 0, NA),
    stringsAsFactors = FALSE
  )

  wb <- qxl(dat, style1 = qstyle(rows = x1 >= 4 & x2 == "f", bgFill = "#fddbc7"))

  expect_equal(
    wb$worksheets[[1]]$conditionalFormatting,
    c(
      `A2:A9` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"3\"><formula>AND($A2 &gt;= 4, $B2 = &quot;f&quot;)</formula></cfRule>",
      `B2:B9` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"2\"><formula>AND($A2 &gt;= 4, $B2 = &quot;f&quot;)</formula></cfRule>",
      `C2:C9` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"1\"><formula>AND($A2 &gt;= 4, $B2 = &quot;f&quot;)</formula></cfRule>"
    )
  )
})

