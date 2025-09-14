################################################################################
# This script generates 3-set Venn diagrams (Pfizer, Moderna, Janssen) from
# specified Excel workbooks (each with sheets: PFIZER-BIONTECH, MODERNA, JANSSEN).
# It saves:
#   (1) ggVennDiagram PDF,
#   (2) VennDetail PDF,
#   (3) intersection table joined with source sheet rows (xlsx).
#
# Inputs: files under ./trimmed_excel_output and ./excel_output
# Outputs: ./VennOutput2
################################################################################

################################################################################
# Set the current working directory
################################################################################
library(rstudioapi) # make sure you have it installed
base_dir <- dirname(getActiveDocumentContext()$path)
setwd(base_dir)

################################################################################
## Load required packages
################################################################################
if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman")
library(pacman)

# CRAN packages
cran_packages <- c("readxl", "writexl", "dplyr", "ggplot2", "ggVennDiagram", "tibble")
# Bioconductor packages
bioc_packages <- c("VennDetail")

p_load(char = cran_packages, install = TRUE)
p_load(char = bioc_packages, install = TRUE, repos = BiocManager::repositories())

################################################################################
# Define output directory
################################################################################
output_dir <- file.path(base_dir, "VennOutput")
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

################################################################################
# Helper: ensure all 3-way subsets exist (padding with dummy values if empty)
################################################################################
ensure_all_subsets <- function(venn_list) {
  # Only meaningful for <=3 sets; here we have exactly 3
  set_names <- names(venn_list)
  if (length(set_names) != 3) return(venn_list)
  
  all_combos <- unlist(lapply(1:3, function(k) combn(set_names, k, simplify = FALSE)), recursive = FALSE)
  dummy_prefix <- "__dummy__"
  counter <- 1
  
  for (combo in all_combos) {
    intersection <- Reduce(intersect, venn_list[combo])
    if (length(intersection) == 0) {
      dummy_val <- paste0(dummy_prefix, counter)
      counter <- counter + 1
      # add dummy to all participating sets so the intersection is non-empty
      for (nm in combo) {
        venn_list[[nm]] <- c(venn_list[[nm]], dummy_val)
      }
    }
  }
  venn_list
}

################################################################################
# Function to generate Venn diagrams and save outputs
################################################################################
process_venn_diagrams <- function(file_path, prefix, vaccines = c("PFIZER", "MODERNA", "JANSSEN")) {
  # Read data from each sheet
  pfizer_data  <- read_excel(file_path, sheet = "PFIZER-BIONTECH")
  moderna_data <- read_excel(file_path, sheet = "MODERNA")
  janssen_data <- read_excel(file_path, sheet = "JANSSEN")
  
  # Standardize first column name to "AE" for joins later
  if (ncol(pfizer_data)  > 0) names(pfizer_data)[1]  <- "AE"
  if (ncol(moderna_data) > 0) names(moderna_data)[1] <- "AE"
  if (ncol(janssen_data) > 0) names(janssen_data)[1] <- "AE"
  
  # Extract AE vectors (first column)
  pfizer_ae  <- if (ncol(pfizer_data)  > 0) pfizer_data[[1]]  else character(0)
  moderna_ae <- if (ncol(moderna_data) > 0) moderna_data[[1]] else character(0)
  janssen_ae <- if (ncol(janssen_data) > 0) janssen_data[[1]] else character(0)
  
  # Prepare list for Venn
  venn_list <- list(
    Pfizer  = unique(na.omit(pfizer_ae)),
    Moderna = unique(na.omit(moderna_ae)),
    Janssen = unique(na.omit(janssen_ae))
  )
  
  # Colors (kept as a simple named vector; ggVennDiagram may ignore set-specific colors)
  custom_colors <- c(Pfizer = "#023981", Moderna = "#d62728", Janssen = "#2ca02c")
  
  # ------------------------------
  # ggVennDiagram PDF
  # ------------------------------
  venn_plot <- ggVennDiagram(
    venn_list,
    label = "count",
    label_alpha = 0
  ) +
    scale_fill_gradient(low = "white", high = "red") +
    theme(legend.position = "right")
  
  gg_out <- file.path(output_dir, paste0(prefix, "_VennDiagram_ggVennDiagram.pdf"))
  ggsave(filename = gg_out, plot = venn_plot, width = 10, height = 8)
  
  # ------------------------------
  # VennDetail PDF (robust intersections, even if empty -> pad)
  # ------------------------------
  venn_list_padded <- ensure_all_subsets(venn_list)
  three_way_venn <- VennDetail::venndetail(venn_list_padded)
  
  vd_out <- file.path(output_dir, paste0(prefix, "_VennDiagram_VennDetail.pdf"))
  pdf(vd_out, width = 10, height = 8)
  plot(
    three_way_venn,
    mycol = custom_colors,
    col = custom_colors,
    alpha = 0.5,
    cex = 3.8,
    cat.cex = 3,
    cat.fontface = "bold",
    margin = 0.1
  )
  dev.off()
  
  # ------------------------------
  # Intersection table + sheet joins
  # ------------------------------
  venn_results <- result(three_way_venn)  # from VennDetail
  # The column with element names is usually named "Detail"
  if (!"Detail" %in% names(venn_results)) {
    stop("Expected column 'Detail' not found in VennDetail result.")
  }
  
  combined_results <- venn_results %>%
    left_join(pfizer_data,  by = c("Detail" = "AE"), keep = FALSE, relationship = "many-to-many") %>%
    left_join(moderna_data, by = c("Detail" = "AE"), keep = FALSE, relationship = "many-to-many") %>%
    left_join(janssen_data, by = c("Detail" = "AE"), keep = FALSE, relationship = "many-to-many")
  
  xlsx_out <- file.path(output_dir, paste0(prefix, "_VennResults.xlsx"))
  writexl::write_xlsx(combined_results, xlsx_out)
  
  cat("Processed: ", basename(file_path), "\n", sep = "")
  cat("  - ggVennDiagram: ", gg_out, "\n", sep = "")
  cat("  - VennDetail:    ", vd_out, "\n", sep = "")
  cat("  - Results table: ", xlsx_out, "\n", sep = "")
}

################################################################################
# Main script
################################################################################

# Define input files
files <- list(
  Trimmed_OUTPUT6   = file.path(base_dir, "trimmed_excel_output", "OUTPUT6.xlsx"),
  Trimmed_OUTPUT9   = file.path(base_dir, "trimmed_excel_output", "OUTPUT9.xlsx"),
  NonTrimmed_OUTPUT6 = file.path(base_dir, "excel_output", "OUTPUT6.xlsx"),
  NonTrimmed_OUTPUT9 = file.path(base_dir, "excel_output", "OUTPUT9.xlsx")
)

# Process each file
for (file_name in names(files)) {
  if (file.exists(files[[file_name]])) {
    process_venn_diagrams(files[[file_name]], file_name, c("PFIZER", "MODERNA", "JANSSEN"))
  } else {
    warning("File not found: ", files[[file_name]])
  }
}

cat("All done. Outputs saved to '", output_dir, "'.\n", sep = "")
