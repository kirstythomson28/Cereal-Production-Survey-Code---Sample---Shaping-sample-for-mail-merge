library(readr)
library(dplyr)
library(tidyr)
library(stringr)
library(writexl)

# Read in main sample received from census team (test farms excluded)
full_data <- read_csv("Main_sample_2025 - Excludes test farms.csv")

# Prepare November data: extract BRN, email, and formatted location code
nov_data <- full_data %>%
  select(brn, primary_email, parish, holding) %>%
  mutate(
    location = str_c(
      str_pad(as.character(parish), 3, pad = "0"),
      str_pad(as.character(holding), 4, pad = "0"),
      sep = "/"
    )
  ) %>%
  filter(!is.na(location) & location != "")  # Remove missing/blank locations

# Reshape data for mail merge: one row per BRN/email, locations spread across columns
mail_merge <- nov_data %>%
  group_by(brn, primary_email) %>%
  arrange(location, .by_group = TRUE) %>%
  mutate(location_num = row_number()) %>%
  pivot_wider(
    names_from = location_num,
    values_from = location,
    names_prefix = "location"
  ) %>%
  ungroup()

# Export spreadsheet ready for mail merge
write_xlsx(mail_merge, "2025-26 - Production - Materials - Sample - November sample for mailmerge excluding friendly farm.xlsx")





