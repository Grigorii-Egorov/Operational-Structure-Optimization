#Add essentia packages
library(readxl)
library(openxlsx)
library(writexl)
library(corrplot)
library(dplyr)
library(tidyr)
library(stringr)
library(ggplot2)
library(lmtest)
library(stargazer)
library(EnvStats)
library(RColorBrewer)
library(reshape2)
library(FinCal)
library(FinancialMath)
library(pander)
library(car)
library(lubridate)
library(tidyverse)



data <- read_csv("/Users/grigoriiegorov/Desktop/SHBN.csv")
data <- data %>%
  rename(place = `Площадка/ Дивизион`) %>%
  rename(employee =`Административное подчинение (в расч. На одного)`)


# 2. Filter rows where Должность contains (any case)
mayo_df <- data %>%
  filter(str_detect(Должность, regex("ороситель", ignore_case = TRUE))) %>%
  filter(str_detect(Функция, regex("01 Операционная деятельность", ignore_case = TRUE)))



row_order <- c(
  "Численность сейчас",
  "В подчинении на одного",
  "Целевая оптимизация"
)





clean <- mayo_df %>%
  rename(`В подчинении на одного` = employee) %>%      # rename first
  mutate(
    `Численность сейчас` = as.numeric(`Численность сейчас`),
    `В подчинении на одного` = as.numeric(`В подчинении на одного`),
    `Целевая оптимизация` = as.numeric(`Целевая оптимизация`)
  ) %>%
  replace_na(list(`Целевая оптимизация` = 0)) %>%       # replace NA → 0
  select(
    place,
    Должность,
    `Численность сейчас`,
    `В подчинении на одного`,
    `Целевая оптимизация`
  ) %>%
  mutate(
    `В подчинении всего` = `В подчинении на одного` * `Численность сейчас`,
    `Целевая численность` = `Численность сейчас` - `Целевая оптимизация`
  )


row_order <- c(
  "Численность сейчас",
  "В подчинении на одного",
  "В подчинении всего",
  "Целевая оптимизация",
  "Целевая численность"
)

long <- clean %>%
  pivot_longer(
    cols = c(
      `Численность сейчас`,
      `В подчинении на одного`,
      `В подчинении всего`,
      `Целевая оптимизация`,
      `Целевая численность`
    ),
    names_to = "Показатель",
    values_to = "Значение"
  ) %>%
  mutate(
    Значение = as.numeric(Значение)
  )

result <- long %>%
  group_by(place, Должность, Показатель) %>%            # ← ДУБЛИКАТЫ
  summarise(
    Значение = sum(Значение, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    Показатель = factor(Показатель, levels = row_order)
  ) %>%
  arrange(Показатель) %>%
  pivot_wider(
    names_from = place,
    values_from = Значение,
    values_fill = 0
  )

result

# File save directory
file_path <- "/Users/grigoriiegorov/Library/Mobile Documents/com~apple~CloudDocs/Work/Rusagro/Optimization/Devided/СХБН/01 Операционная деятельность/Оросительные системы.xlsx"
# Saving
write_xlsx(result, file_path)




library(readxl)
library(openxlsx)

# Folder containing your Excel files
folder <- "/Users/grigoriiegorov/Library/Mobile Documents/com~apple~CloudDocs/Work/Rusagro/Optimization/Devided/СХБН/01 Операционная деятельность/"

# List all .xlsx files in folder
files <- list.files(folder, pattern = "\\.xlsx$", full.names = TRUE)

# Create new workbook
wb <- createWorkbook()

# Loop through files and add each to a new sheet
for (f in files) {
  sheet_name <- tools::file_path_sans_ext(basename(f))  # filename as sheet name
  
  addWorksheet(wb, sheet_name)
  
  # Read first sheet of the file
  df <- read_excel(f)
  
  writeData(wb, sheet_name, df)
}

# Save final workbook
saveWorkbook(wb, "/Users/grigoriiegorov/Library/Mobile Documents/com~apple~CloudDocs/Work/Rusagro/Optimization/Devided/СХБН/01 Операционная деятельность/merged.xlsx", overwrite = TRUE)




