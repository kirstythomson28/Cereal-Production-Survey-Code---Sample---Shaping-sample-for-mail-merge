
library(readr)
library(dplyr)
library(tidyr)
library(readxl)
library(writexl)
library(purrr)
library(stringr)

#Read in mail merge spreasheet from november
full_data <- read_excel("Cereal Production and Disposal Survey - 2024-25 - Production - Materials - Sample - Main sample mailmerge excluding friendly farm.xlsx")
#read in july sample
brn_data <- read_excel("Cereal Production and Disposal Survey - 2024-25 - Disposals - July 2025 - Location data upload - draft.xlsx")

# give july sample headings
colnames(brn_data) <- c("parish", "holding")
# put july sample in location code format with leading zeros
brn_data <- brn_data %>%
  mutate(
    location1 = str_pad(parish, width = 3, pad = "0") %>%
      paste0("/", str_pad(holding, width = 4, pad = "0"))
  )

#filter november spreadsheet to just brn, email and location codes
nov_data <- full_data %>%
  select(,3:50)%>%
  select(,-3)

#remove anylocations with no values
nov_data <- nov_data[, colSums(!is.na(nov_data)) > 0]

#transform data into long format
long_df <- nov_data %>%
  pivot_longer(
    cols = starts_with("location"),
    names_to = "location_type",
    values_to = "location1"
  ) %>%
  filter(!is.na(location1)) %>%
  select(-location_type)        

#filter data to remove any locations not in july sample and then back into wide format
mail_merge <- brn_data %>%
  left_join(long_df, by = "location1")%>%
  select(,-c(1,2))%>%
  group_by(brn) %>%
  mutate(location_num = paste0("location", row_number())) %>%
  ungroup()%>%
  pivot_wider(
    names_from = location_num,
    values_from = location1
  )


#export spreadsheet ready for the mailmerge
write_xlsx(mail_merge, "Cereal Production and Disposal Survey - 2024-25 - Disposals - Materials - Sample - July sample mailmerge excluding friendly farm.xlsx")





