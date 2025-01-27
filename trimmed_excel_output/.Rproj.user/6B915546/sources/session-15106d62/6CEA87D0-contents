# Load necessary libraries
library(readxl)      # For reading Excel files
library(writexl)     # For writing Excel files
library(dplyr)       # For data manipulation
library(VennDetail)  # For creating Venn diagrams

# (1) Get the current location of the current script and set it as the working directory
#setwd(dirname(sys.frame(1)$ofile))

# (2) Read the Excel file
file_path <- "./OUTPUT9.xlsx"
sheets <- excel_sheets(file_path)

# (3) Read the specific workshee`ts for comparison
pfizer_data <- read_excel(file_path, sheet = "PFIZER-BIONTECH")
pfizer_bivalent_data <- read_excel(file_path, sheet = "PFIZER-BIONTECH BIVALENT")
moderna_data <- read_excel(file_path, sheet = "MODERNA")
moderna_bivalent_data <- read_excel(file_path, sheet = "MODERNA BIVALENT")

# (4) Extract the first column (AE) for comparison
pfizer_ae <- pfizer_data[[1]]
pfizer_bivalent_ae <- pfizer_bivalent_data[[1]]
moderna_ae <- moderna_data[[1]]
moderna_bivalent_ae <- moderna_bivalent_data[[1]]

# (5) Create Venn diagrams
pfizer_venn <- VennDetail::venndetail(list(PFIZER = pfizer_ae, PFIZER_BIVALENT = pfizer_bivalent_ae))
moderna_venn <- VennDetail::venndetail(list(MODERNA = moderna_ae, MODERNA_BIVALENT = moderna_bivalent_ae))

# (6) Save the Venn diagrams as PDF images
pdf("PFIZER_Venn_Diagram_Mono-vs-Bi.pdf")
plot(pfizer_venn)
dev.off()

pdf("MODERNA_Venn_Diagram_Mono-vs-Bi.pdf")
plot(moderna_venn)
dev.off()

# (7) Save the Venn diagrams detail as an Excel file
pfizer_venn_detail <- result(pfizer_venn)
pfizer_combined <- pfizer_venn_detail %>%
  left_join(pfizer_data, by = c("Detail" = "AE")) %>%  # Merge with PFIZER data using 'Detail'
  left_join(pfizer_bivalent_data, by = c("Detail" = "AE"))  # Merge with PFIZER-BIVALENT data
write_xlsx(pfizer_combined, "PFIZER_Combined_Output.xlsx")

# Repeat for MODERNA
moderna_venn_detail <- result(moderna_venn)
moderna_combined <- moderna_venn_detail %>%
  left_join(moderna_data, by = c("Detail" = "AE")) %>%  # Merge with MODERNA data using 'Detail'
  left_join(moderna_bivalent_data, by = c("Detail" = "AE"))  # Merge with MODERNA BIVALENT data
write_xlsx(moderna_combined, "MODERNA_Combined_Output.xlsx")

