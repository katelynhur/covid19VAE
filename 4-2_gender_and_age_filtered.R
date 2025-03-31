################################################################################
#####################  Packages  ###############################################
################################################################################

# Install pacman for package management if not already installed
if (!requireNamespace("pacman", quietly = TRUE)) {
  install.packages("pacman")
}
library(pacman)

# Required CRAN packages
cran_packages <- c(
  "readxl", "writexl", "tidyverse", "ggplot2", "ComplexUpset", 
  "UpSetR", "ComplexHeatmap"
)

# Automatically install and load packages
p_load(char = cran_packages, install = TRUE)

################################################################################
#####################  Set Current Working Directory ###########################
################################################################################

current_path <- rstudioapi::getActiveDocumentContext()$path 
setwd(dirname(current_path))
print(current_path)
base_dir <- paste0(dirname(current_path),"/vaxafe_male-vs-female_age/")

# Create output directory
data_dir <- paste0(base_dir,"downloaded/")
output_dir <- file.path(base_dir,"/", "filtered_age_vs_gender_output")
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

################################################################################
#####################  Read and Process Excel Files ############################
################################################################################

# Define the manufacturers and genders to include
valid_manufacturers <- c("PFIZER-BIONTECH_BIVALENT", "PFIZER-BIONTECH", "MODERNA", "MODERNA_BIVALENT")
valid_genders <- c("Female", "Male")

# Define the adverse events of interest
target_events <- c("Myocarditis", "Pericarditis", "Acute myocardial infarction", "Acute respiratory failure", "Hypoxia")

# Initialize a dataframe to store results
results_df <- data.frame()

for (file in excel_files) {
  # Extract vaccine, manufacturer, age group, and gender from filename
  file_name <- tools::file_path_sans_ext(basename(file))  
  file_info <- str_match(file_name, "^(.*?)_\\((.*?)\\)_(.*?)_(.*?)$")  
  
  if (is.na(file_info[1,1])) next  # Skip if filename format doesn't match
  
  manufacturer <- file_info[3]
  age_group <- file_info[4]
  gender <- file_info[5]
  
  # Extract only the manufacturer name inside parentheses
  manufacturer <- str_match(manufacturer, "\\(([^()]+)\\)")[,2]
  
  # **Apply Filters**
  if (!(manufacturer %in% valid_manufacturers)) next  # Skip invalid manufacturers
  if (!(gender %in% valid_genders)) next  # Skip "Any" gender
  
  # Read the Excel file
  df <- read_excel(file, sheet = 1)
  
  print(file)
  
  # Check for required columns
  if (!all(c("Adverse Event", "AE Cases", "Total Cases") %in% colnames(df))) {
    next  # Skip files without the expected structure
  }
  
  # Ensure AE Cases and Total Cases are numeric
  df <- df %>%
    mutate(`AE Cases` = as.numeric(`AE Cases`),
           `Total Cases` = as.numeric(`Total Cases`))
  
  # Extract relevant adverse events
  extracted_data <- df %>% 
    filter(`Adverse Event` %in% target_events) %>% 
    mutate(Percentage = (`AE Cases` / `Total Cases`) * 100) %>% 
    select(`Adverse Event`, Percentage)
  
  # Append to results_df
  if (nrow(extracted_data) > 0) {
    extracted_data <- extracted_data %>%
      mutate(Manufacturer = manufacturer,
             Gender = gender, Age_Group = age_group)
    
    results_df <- bind_rows(results_df, extracted_data)
  }
}

results_df <- results_df[,c("Manufacturer", "Adverse Event", "Gender", "Age_Group", "Percentage")]

# Save extracted data
write_xlsx(results_df, file.path(output_dir, "filtered_extracted_data.xlsx"))


################################################################################
#####################  Data Visualization ######################################
################################################################################

# Define color mapping for gender
gender_colors <- c("Female" = "red", "Male" = "blue")

# Function to create faceted plots
plot_adverse_event <- function(event) {
  event_data <- results_df %>% filter(`Adverse Event` == event)
  
  # Rename manufacturers
  event_data <- event_data %>%
    mutate(Manufacturer = case_when(
      Manufacturer == "PFIZER-BIONTECH" ~ "Pfizer Mono",
      Manufacturer == "PFIZER-BIONTECH_BIVALENT" ~ "Pfizer Bi",
      Manufacturer == "MODERNA" ~ "Moderna Mono",
      Manufacturer == "MODERNA_BIVALENT" ~ "Moderna Bi",
      TRUE ~ Manufacturer  # Keep other names unchanged
    ))
  
  # Bar Graph with facet wrap
  bar_plot <- ggplot(event_data, aes(x = Age_Group, y = Percentage, fill = Gender)) +
    geom_bar(stat = "identity", position = "dodge") +
    scale_fill_manual(values = gender_colors) +
    theme_minimal() +
    labs(title = event, x = "Age Group", y = "Percentage of Cases (%)") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
    facet_wrap(~ Manufacturer, ncol = 2)
  
  line_plot <- ggplot(event_data, aes(x = Age_Group, y = Percentage, group = Gender, color = Gender)) +
    geom_line(size = 1.5) +   # Thicker lines
    geom_point(size = 4) +    # Larger dots
    scale_color_manual(values = gender_colors) +
    theme_minimal() +
    labs(title = event, 
         x = "Age Group", y = "Percentage of Cases (%)") +
    theme(
      plot.title = element_text(size = 18, face = "bold", hjust = 0.5),   # Larger title
      axis.title.x = element_text(size = 16),               # Larger x-axis label
      axis.title.y = element_text(size = 16),               # Larger y-axis label
      axis.text.x = element_text(size = 16, angle = 45, hjust = 1),  # Larger x-axis text
      axis.text.y = element_text(size = 16),               # Larger y-axis text
      legend.title = element_text(size = 16),              # Larger legend title
      legend.text = element_text(size = 14),               # Larger legend labels
      strip.text = element_text(size = 16, face = "bold")  # Larger facet labels (Manufacturer names)
    ) +
    facet_wrap(~ Manufacturer, ncol = 2)
  
  # Save plots
  ggsave(filename = file.path(output_dir, paste0(event, "_Filtered_BarPlot.pdf")), plot = bar_plot, width = 12, height = 8)
  ggsave(filename = file.path(output_dir, paste0(event, "_Filtered_LinePlot.pdf")), plot = line_plot, width = 12, height = 8)
  ggsave(filename = file.path(output_dir, paste0(event, "_Filtered_BarPlot.png")), plot = bar_plot, width = 12, height = 8)
  ggsave(filename = file.path(output_dir, paste0(event, "__Filtered_LinePlot.png")), plot = line_plot, width = 12, height = 8)
}

# Generate and save plots for both events
plot_adverse_event("Myocarditis")
plot_adverse_event("Pericarditis")
plot_adverse_event("Acute myocardial infarction")
plot_adverse_event("Acute respiratory failure")
plot_adverse_event("Hypoxia")


