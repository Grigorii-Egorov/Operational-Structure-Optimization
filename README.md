# Operational-Structure-Optimization
This project processes internal staffing data, filters specific operational roles, computes key metrics, restructures the table into long and wide formats, and exports the results into Excel.   The final script also merges all generated Excel files into one consolidated workbook.



The code is modular and follows a clear linear pipeline:

1. **Filtering and cleaning**
2. **Normalization and numeric conversion**
3. **Metric calculations**
4. **Long-format transformation**
5. **Aggregation and reshaping**
6. **Excel export**
7. **Automatic merging of all output files**

---

## 1. Filtering & Cleaning

- Load raw staffing dataset (`SHBN.csv`).
- Rename important columns (e.g., `Площадка/ Дивизион` → `place`).
- Filter only:
  - rows whose position contains *"ороситель"*  
  - rows belonging to *01 Операционная деятельность*  
- Convert key columns to numeric.
- Replace missing optimization values with `0`.

---

## 2. Metric Calculations

The script computes:

- **В подчинении всего** = `В подчинении на одного` × `Численность сейчас`
- **Целевая численность** = `Численность сейчас` − `Целевая оптимизация`

These KPIs are later used for long-format restructuring.

---

## 3. Long Format

The table is transformed into a long structure via `pivot_longer()`,  
with:

- `Показатель` — type of metric  
- `Значение` — numeric value  

Then grouped and summarised by:

- place  
- position (`Должность`)  
- indicator (`Показатель`)

---

## 4. Wide Format Output

The cleaned long table is expanded back into a wide format:

- `Показатель` becomes the row index  
- Each `place` becomes a separate column  
- Missing values filled with `0`

The result is written to an Excel file.

---

## 5. Automatic Workbook Merge

All generated `.xlsx` files inside the output folder are:

- loaded  
- inserted into a single multi-sheet workbook  
- saved as `merged.xlsx`

This creates one consolidated file for review or management reporting.

---

## Packages Required

readxl
openxlsx
writexl
dplyr
tidyr
stringr
ggplot2
lubridate
tidyverse


## Note

Results and samples are not provided due to company restrictions :)






