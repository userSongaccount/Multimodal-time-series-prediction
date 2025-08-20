rm (list = ls())
library(tidyverse)
library(readxl)

# read inflation rate data
df <- read_excel("C:\\Users\\huawei\\OneDrive\\桌面\\data.xlsx")

# change names, select columns, and create a summary table
library(flextable)
library(officer)
library(gtsummary)
library(broom)

st <- df %>%
  rename(
    `GDP (100 millions)` = GDP,
    `Shenzhen City Average Light Value (nW/cm²/sr)` = city_mean_light,
    `Guangdong Province Average Light Value (nW/cm²/sr)` = region_mean_light,
    `Principal Component 1` = pc1,
    `Principal Component 2` = pc2,
    `Principal Component 3` = pc3
  ) %>%
  select(
    `GDP (100 millions)`,
    `Shenzhen City Average Light Value (nW/cm²/sr)`,
    `Guangdong Province Average Light Value (nW/cm²/sr)`,
    `Principal Component 1`,
    `Principal Component 2`,
    `Principal Component 3`
    ) %>%
  tbl_summary(
    type = all_continuous() ~ "continuous2",
    statistic = all_continuous() ~ c("{mean}",
                                     "{median}({p25}/{p75})",
                                     "{sd}",
                                     "{min} / {max}"
                                     ),
    missing = "no"
  ) %>%
  as_flex_table()

# add header
st1 <- add_header_row(
  x = st, values = c("Descriptive Statistics of Indicators from 2008 to 2025"),
  colwidths = c(2))

# change size
st1 <- fontsize(st1, part = "header", size = 15)
st1 <- font(st1, part = "header", fontname = "Arial")

# add footer
st2 <- add_footer_lines(st1, "Note: All data is quarterly data from 2008 to first quarter of 2025, excluding 2012 and first quarter of 2013 data.
")

st2 <- fontsize(st2, part = "footer", size = 9)
st2 <- font(st2, part = "footer", fontname = "Arial")

# add lines
st3 <- theme_vanilla(st2)
st3

# output table
save_as_docx(
  "Table A1" = st3,
  path = "C:\\Users\\huawei\\OneDrive\\桌面\\XGB.docx")