################################################################################
# This script builds 2-way (Mono vs Bivalent) Venn diagrams for Pfizer & Moderna
# and a 4-way diagram (Pfizer/Moderna × Mono/Bivalent) from an Excel workbook
# that contains sheets:
#   "PFIZER-BIONTECH", "PFIZER-BIONTECH BIVALENT",
#   "MODERNA", "MODERNA BIVALENT"
#
# Outputs (to ./VennOutput):
#   - *_PFIZER_VennDetail_Mono-vs-Bivalent.pdf
#   - *_MODERNA_VennDetail_Mono-vs-Bivalent.pdf
#   - *_FOURWAY_VennDetail_Pfizer-Modernas_Mono-vs-Bivalent.pdf
#   - *_PFIZER_Combined_Output_Mono-vs-Bivalent.xlsx
#   - *_MODERNA_Combined_Output_Mono-vs-Bivalent.xlsx
#   - *_FOURWAY_Combined_Output.xlsx  (intersections + joined sheets)
################################################################################

################################################################################
# Packages
################################################################################
if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman")
library(pacman)

# Include BiocManager so pacman can use its repos
cran_packages <- c("BiocManager", "readxl", "writexl", "dplyr", "tibble")
bioc_packages <- c("VennDetail")

p_load(char = cran_packages, install = TRUE)
p_load(char = bioc_packages, install = TRUE, repos = BiocManager::repositories())

################################################################################
# Set Current Working Directory
################################################################################
library(rstudioapi)
base_dir <- dirname(getActiveDocumentContext()$path)
setwd(base_dir)

################################################################################
# Define Output Directory
################################################################################
output_dir <- file.path(base_dir, "VennOutput")
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

################################################################################
# Helper: read a sheet safely (returns empty tibble if missing)
################################################################################
read_sheet_safe <- function(path, sheet_name) {
  sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character())
  if (!(sheet_name %in% sheets)) {
    warning("Sheet not found: '", sheet_name, "' in file: ", basename(path))
    return(tibble::tibble())
  }
  df <- readxl::read_excel(path, sheet = sheet_name)
  if (ncol(df) > 0) names(df)[1] <- "AE"  # standardize first column for joins
  df
}

################################################################################
# Function to Generate Venn Diagrams + Outputs
################################################################################
process_venn_diagrams <- function(file_path, prefix) {
  if (!file.exists(file_path)) {
    warning("File not found: ", file_path)
    return(invisible(NULL))
  }
  
  # Read source sheets
  pfizer_mono  <- read_sheet_safe(file_path, "PFIZER-BIONTECH")
  pfizer_bi    <- read_sheet_safe(file_path, "PFIZER-BIONTECH BIVALENT")
  moderna_mono <- read_sheet_safe(file_path, "MODERNA")
  moderna_bi   <- read_sheet_safe(file_path, "MODERNA BIVALENT")
  
  # Extract AE vectors (first column)
  vec <- function(df) if (ncol(df) > 0) unique(na.omit(df[[1]])) else character(0)
  pf_mono_vec  <- vec(pfizer_mono)
  pf_bi_vec    <- vec(pfizer_bi)
  md_mono_vec  <- vec(moderna_mono)
  md_bi_vec    <- vec(moderna_bi)
  
  # Color palette
  col_pf    <- "#023981"   # Pfizer mono
  col_pf_bi <- "#92C1FE"   # Pfizer bivalent
  col_md    <- "#d62728"   # Moderna mono
  col_md_bi <- "#ff9896"   # Moderna bivalent
  
  # ---------------------------
  # 2-way Pfizer (Mono vs Bi)
  # ---------------------------
  pf_list <- list("\n\n\nPfizer\nMono" = pf_mono_vec,
                  "\n\n\nPfizer\nBi"   = pf_bi_vec)
  pf_venn <- VennDetail::venndetail(pf_list)
  
  pdf(file.path(output_dir, paste0(prefix, "_PFIZER_VennDetail_Mono-vs-Bivalent.pdf")),
      width = 10, height = 8)
  plot(pf_venn,
       mycol = c(col_pf, col_pf_bi),
       col   = c(col_pf, col_pf_bi),
       alpha = 0.5, cex = 3.8, cat.cex = 3, cat.fontface = "bold", margin = 0.1)
  dev.off()
  
  pf_intersections <- VennDetail::result(pf_venn)
  pf_joined <- pf_intersections %>%
    dplyr::left_join(pfizer_mono, by = c("Detail" = "AE"), suffix = c("", ".PFIZER_MONO")) %>%
    dplyr::left_join(pfizer_bi,   by = c("Detail" = "AE"), suffix = c("", ".PFIZER_BI"))
  writexl::write_xlsx(
    pf_joined,
    file.path(output_dir, paste0(prefix, "_PFIZER_Combined_Output_Mono-vs-Bivalent.xlsx"))
  )
  
  # ---------------------------
  # 2-way Moderna (Mono vs Bi)
  # ---------------------------
  md_list <- list("\n\n\nModerna\nMono" = md_mono_vec,
                  "\n\n\nModerna\nBi"   = md_bi_vec)
  md_venn <- VennDetail::venndetail(md_list)
  
  pdf(file.path(output_dir, paste0(prefix, "_MODERNA_VennDetail_Mono-vs-Bivalent.pdf")),
      width = 10, height = 8)
  plot(md_venn,
       mycol = c(col_md, col_md_bi),
       col   = c(col_md, col_md_bi),
       alpha = 0.5, cex = 3.8, cat.cex = 3, cat.fontface = "bold", margin = 0.1)
  dev.off()
  
  md_intersections <- VennDetail::result(md_venn)
  md_joined <- md_intersections %>%
    dplyr::left_join(moderna_mono, by = c("Detail" = "AE"), suffix = c("", ".MODERNA_MONO")) %>%
    dplyr::left_join(moderna_bi,   by = c("Detail" = "AE"), suffix = c("", ".MODERNA_BI"))
  writexl::write_xlsx(
    md_joined,
    file.path(output_dir, paste0(prefix, "_MODERNA_Combined_Output_Mono-vs-Bivalent.xlsx"))
  )
  
  # ---------------------------
  # 4-way (Pfizer/Moderna × Mono/Bi)
  # ---------------------------
  four_list <- list(
    "Pfizer\nMonovalent"  = pf_mono_vec,
    "Pfizer\nBivalent"    = pf_bi_vec,
    "Moderna\nMonovalent" = md_mono_vec,
    "Moderna\nBivalent"   = md_bi_vec
  )
  four_venn <- VennDetail::venndetail(four_list)
  
  pdf(file.path(output_dir, paste0(prefix, "_FOURWAY_VennDetail_Pfizer-Modernas_Mono-vs-Bivalent.pdf")),
      width = 11, height = 9)
  plot(four_venn,
       mycol = c(col_pf, col_pf_bi, col_md, col_md_bi),
       col   = c(col_pf, col_pf_bi, col_md, col_md_bi),
       alpha = 0.5, cex = 3.2, cat.cex = 2.6, cat.fontface = "bold", margin = 0.08)
  dev.off()
  
  four_intersections <- VennDetail::result(four_venn)
  
  # Heuristic to find the grouping column when name varies
  group_col <- {
    cand <- names(four_intersections)[sapply(names(four_intersections), function(nm) {
      is.character(four_intersections[[nm]]) &&
        any(grepl("Pfizer|Moderna|Mono|Bi", four_intersections[[nm]], ignore.case = TRUE), na.rm = TRUE)
    })]
    if (length(cand) == 0) cand <- intersect(c("Group","Groups","Set","Sets","Category","Cate"),
                                             names(four_intersections))
    if (length(cand) == 0) cand <- names(four_intersections)[1]
    cand[1]
  }
  
  four_joined <- four_intersections %>%
    dplyr::left_join(pfizer_mono,  by = c("Detail" = "AE"), suffix = c("", ".PFIZER_MONO")) %>%
    dplyr::left_join(pfizer_bi,    by = c("Detail" = "AE"), suffix = c("", ".PFIZER_BI")) %>%
    dplyr::left_join(moderna_mono, by = c("Detail" = "AE"), suffix = c("", ".MODERNA_MONO")) %>%
    dplyr::left_join(moderna_bi,   by = c("Detail" = "AE"), suffix = c("", ".MODERNA_BI"))
  
  writexl::write_xlsx(
    list(
      Intersections       = four_intersections,
      CombinedWithAllSheets = four_joined,
      MembershipOnly        = four_intersections[, c(group_col, "Detail"), drop = FALSE]
    ),
    path = file.path(output_dir, paste0(prefix, "_FOURWAY_Combined_Output.xlsx"))
  )
  
  cat("Processed: ", basename(file_path), "\n", sep = "")
}

################################################################################
# Process All Files
################################################################################
files <- list(
  NonTrimmed = file.path(base_dir, "excel_output", "OUTPUT9.xlsx"),
  Trimmed    = file.path(base_dir, "trimmed_excel_output", "OUTPUT9.xlsx")
)

for (nm in names(files)) process_venn_diagrams(files[[nm]], nm)

cat("All done. Outputs saved to '", output_dir, "'.\n", sep = "")
