library(readr)
library(dplyr)
library(tidyr)
library(stringr)
library(writexl)

#Read in main sample recieved from census team with test farms excluded
full_data <- read_csv("Main_sample_2025 - Excludes test farms.csv")


#filter november spreadsheet to just brn, email and location codes 
nov_data <- full_data %>% 
  select(brn, primary_email, parish, holding) %>%
  mutate( parish = str_pad(as.character(parish), width = 3, pad = "0"), 
          holding = str_pad(as.character(holding), width = 4, pad = "0"), 
          location = paste(parish, holding, sep = "/") ) %>%
  select(-parish, -holding) %>% 
  filter(!is.na(location) & location != "") # removes missing/blank locations

#transform data into long format 
mail_merge <- nov_data %>% 
  group_by(brn, primary_email) %>% 
  arrange(location, .by_group = TRUE) %>% 
  mutate(location_num = row_number()) %>% 
  pivot_wider( names_from = location_num, values_from = location, names_prefix = "location" ) %>% 
  ungroup()

#export spreadsheet ready for the mailmerge 
write_xlsx(mail_merge, "2025-26 - Production - Materials - Sample - November sample for mailmerge excluding friendly farm.xlsx")