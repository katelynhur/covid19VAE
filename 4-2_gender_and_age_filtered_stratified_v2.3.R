################################################################################
#####################  Packages  ###############################################
################################################################################

if (!requireNamespace("pacman", quietly = TRUE)) {
  install.packages("pacman")
}
library(pacman)

cran_packages <- c(
  "readxl", "writexl", "tidyverse", "ggplot2", "ComplexUpset",
  "UpSetR", "ComplexHeatmap", "patchwork", "ggh4x"
)
p_load(char = cran_packages, install = TRUE)

################################################################################
#####################  Set Current Working Directory ###########################
################################################################################

current_path <- rstudioapi::getActiveDocumentContext()$path 
setwd(dirname(current_path))
print(current_path)
base_dir <- paste0(dirname(current_path),"/vaxafe_male-vs-female_age/")

data_dir <- paste0(base_dir,"downloaded/")
output_dir <- file.path(base_dir,"/", "AgeGender_v2.3")
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

################################################################################
#####################  Read and Process Excel Files ############################
################################################################################

excel_files <- list.files(data_dir, pattern = "\\.xlsx$", full.names = TRUE)

valid_manufacturers <- c("PFIZER-BIONTECH_BIVALENT", "PFIZER-BIONTECH", "MODERNA", "MODERNA_BIVALENT","JANSSEN", "NOVAVAX")
valid_genders <- c("Female", "Male")
target_events <- c("Injection site pruritus","Vaccination site warmth","Vaccination site pruritus","Vaccination site erythema",
                   "Paraesthesia oral","Injection site rash","Vaccination site swelling","Swollen tongue","Vaccination site rash",
                   "Lymph node pain","Vaccination site reaction","Hypoaesthesia oral","Axillary pain","Pharyngeal swelling",
                   "Aphonia","Migraine","Throat irritation","Hot flush","Lymphadenopathy","Dysgeusia","Contusion","Myocarditis",
                   "Troponin increased","Acute myocardial infarction","Death","Pericarditis","Mechanical ventilation",
                   "Myocardial infarction","Acute kidney injury","Blood creatinine increased","Sepsis","Acute respiratory failure",
                   "Obstructive sleep apnoea syndrome","Atrial fibrillation","Hypoxia","Angiogram pulmonary abnormal",
                   "Echocardiogram abnormal","Respiratory failure","Deep vein thrombosis","Electrocardiogram abnormal",
                   "Computerised tomogram thorax abnormal","Ultrasound Doppler abnormal",
				   "Thrombosis", "Guillain-Barre syndrome")

# # Define all age groups explicitly
# all_age_groups <- c("0-9", "10-19", "20-29", "30-39", "40-49",
#                     "50-59", "60-69", "70-79", "80+")

results_df <- data.frame()

for (file in excel_files) {
  file_name <- tools::file_path_sans_ext(basename(file))  
  file_info <- str_match(file_name, "^(.*?)_\\((.*?)\\)_(.*?)_(.*?)$")  
  
  if (is.na(file_info[1,1])) next
  
  manufacturer <- file_info[3]
  age_group <- file_info[4]
  gender <- file_info[5]
  manufacturer <- str_match(manufacturer, "\\(([^()]+)\\)")[,2]
  
  if (!(manufacturer %in% valid_manufacturers)) next
  if (!(gender %in% valid_genders)) next
  
  df <- read_excel(file, sheet = 1)
  print(file)
  
  if (!all(c("Adverse Event", "AE Cases", "Total Cases") %in% colnames(df))) next
  
  df <- df %>%
    mutate(`AE Cases` = as.numeric(`AE Cases`),
           `Total Cases` = as.numeric(`Total Cases`))
  
  extracted_data <- df %>% 
    filter(`Adverse Event` %in% target_events) %>% 
    mutate(Percentage = (`AE Cases` / `Total Cases`) * 100) %>% 
    select(`Adverse Event`, Percentage)
  
  if (nrow(extracted_data) > 0) {
    extracted_data <- extracted_data %>%
      mutate(Manufacturer = manufacturer,
             Gender = gender, Age_Group = age_group)
    
    results_df <- bind_rows(results_df, extracted_data)
  }
}

results_df <- results_df[,c("Manufacturer", "Adverse Event", "Gender", "Age_Group", "Percentage")]
write_xlsx(results_df, file.path(output_dir, "extracted_data.xlsx"))

################################################################################
#####################  Plotting with patchwork #################################
################################################################################

gender_colors <- c("Female" = "red", "Male" = "blue")

plot_adverse_event <- function(event) {
  event_data <- results_df %>% filter(`Adverse Event` == event)
  
  # Clean and relabel age groups
  event_data <- event_data %>%
    mutate(
      Age_Group = ifelse(Age_Group == "80-120", "80+", Age_Group)
    ) %>%
    filter(!is.na(Age_Group), Age_Group != "90+")
  
  # Set correct age group order
  all_age_groups <- c("0-9", "10-19", "20-29", "30-39", "40-49",
                      "50-59", "60-69", "70-79", "80+")
  event_data$Age_Group <- factor(event_data$Age_Group, levels = all_age_groups)
  
  # Relabel manufacturers
  event_data <- event_data %>%
    mutate(Manufacturer = case_when(
      Manufacturer == "PFIZER-BIONTECH" ~ "Pfizer Monovalent",
      Manufacturer == "PFIZER-BIONTECH_BIVALENT" ~ "Pfizer Bivalent",
      Manufacturer == "MODERNA" ~ "Moderna Monovalent",
      Manufacturer == "MODERNA_BIVALENT" ~ "Moderna Bivalent",
      Manufacturer == "JANSSEN" ~ "Janssen",
      Manufacturer == "NOVAVAX" ~ "Novavax",
      TRUE ~ Manufacturer
    ))
  
  manufacturer_levels <- c("Pfizer Monovalent", "Moderna Monovalent","Janssen",
                           "Pfizer Bivalent", "Moderna Bivalent"#, "Novavax"
                           )
  event_data$Manufacturer <- factor(event_data$Manufacturer, levels = manufacturer_levels)
  
  # Determine max y value for consistent scaling
  max_y <- ceiling(max(event_data$Percentage, na.rm = TRUE))
  
  plots <- lapply(manufacturer_levels, function(mfg) {
    df <- event_data %>% filter(Manufacturer == mfg)
    df$Age_Group <- factor(df$Age_Group, levels = all_age_groups)
    
    ggplot(df, aes(x = Age_Group, y = Percentage, group = Gender, color = Gender)) +
      geom_line(linewidth  = 1.5) +
      geom_point(size = 4) +
      scale_x_discrete(drop = FALSE) +
      scale_color_manual(values = gender_colors) +
      coord_cartesian(ylim = c(0, max_y)) +
      theme_minimal() +
      labs(title = mfg, x = "Age Group", y = "Percentage of Cases (%)") +
      theme(
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        axis.title.x = element_text(size = 14),
        axis.title.y = element_text(size = 14),
        axis.text.x = element_text(size = 12, angle = 45, hjust = 1),
        axis.text.y = element_text(size = 12),
        legend.title = element_text(size = 12),
        legend.text = element_text(size = 10)
      )
  })
  
  line_plot <- wrap_plots(plots, ncol = 3) +
    plot_layout(ncol = 3) +
    plot_annotation(
      title = event,
      theme = theme(plot.title = element_text(size = 18, face = "bold", hjust = 0.5))
    ) &
    theme(plot.margin = margin(t = 10, b = 20, l = 10, r = 10))  # Add vertical spacing

  # # alternatively
  # line_plot <- (plots[[1]] | plots[[2]] | plots[[3]]) /
  #   (plots[[4]] | plots[[5]] | plot_spacer())
  
  
  ggsave(filename = file.path(output_dir, paste0(event, "_Filtered_LinePlot.png")),
         plot = line_plot, width = 14, height = 10, bg = "white")
  
  ggsave(filename = file.path(output_dir, paste0(event, "_Filtered_LinePlot.pdf")),
         plot = line_plot, width = 14, height = 10, device = cairo_pdf)
}

for (event in target_events) {
  tryCatch({
    plot_adverse_event(event)
  }, error = function(e) {
    message(paste("Skipping", event, "due to error:", e$message))
  })
}


################################################################################
#####################  Combined 5x3 panel (free Y, tidy strips) ################
################################################################################

selected_aes <- c("Myocarditis", "Guillain-Barre syndrome", "Thrombosis")

combined_data <- results_df %>%
  dplyr::filter(`Adverse Event` %in% selected_aes) %>%
  dplyr::mutate(
    Age_Group = ifelse(Age_Group == "80-120", "80+", Age_Group),
    Manufacturer = dplyr::case_when(
      Manufacturer == "PFIZER-BIONTECH" ~ "Pfizer Monovalent",
      Manufacturer == "PFIZER-BIONTECH_BIVALENT" ~ "Pfizer Bivalent",
      Manufacturer == "MODERNA" ~ "Moderna Monovalent",
      Manufacturer == "MODERNA_BIVALENT" ~ "Moderna Bivalent",
      Manufacturer == "JANSSEN" ~ "Janssen",
      Manufacturer == "NOVAVAX" ~ "Novavax",
      TRUE ~ NA_character_       # Mark any other manufacturer as NA
    )
  ) %>%
  # Drop rows with NA Manufacturer or unwanted age groups
  dplyr::filter(!is.na(Age_Group), Age_Group != "90+", !is.na(Manufacturer)) %>%
  # Keep only the five main manufacturers
  dplyr::filter(Manufacturer %in% c(
    "Pfizer Monovalent", "Moderna Monovalent",
    "Janssen", "Pfizer Bivalent", "Moderna Bivalent"
  )) %>%
  dplyr::mutate(
    Age_Group = factor(
      Age_Group,
      levels = c("0-9","10-19","20-29","30-39","40-49",
                 "50-59","60-69","70-79","80+")
    ),
    Manufacturer = factor(
      Manufacturer,
      levels = c("Pfizer Monovalent","Moderna Monovalent",
                 "Janssen","Pfizer Bivalent","Moderna Bivalent")
    ),
    `Adverse Event` = factor(
      `Adverse Event`,
      levels = c("Myocarditis","Guillain-Barre syndrome","Thrombosis")
    )
  ) %>%
  droplevels()

combined_plot_freeY_clean <- ggplot(
  combined_data,
  aes(x = Age_Group, y = Percentage, group = Gender, color = Gender)
) +
  geom_line(linewidth = 1.5) +
  geom_point(size = 4) +
  scale_x_discrete(drop = FALSE) +
  scale_color_manual(values = gender_colors) +
  theme_minimal() +
  labs(
    title = "Age-by-Gender Profiles (Independent Y per Panel)",
    x = "Age Group",
    y = "Percentage of Cases (%)"
  ) +
  ggh4x::facet_grid2(
    rows = vars(Manufacturer),
    cols = vars(`Adverse Event`),
    scales = "free_y",
    independent = "y",
    switch = "both",           # move row labels left, col labels bottom
    labeller = label_value
  ) +
  theme(
    plot.title   = element_text(size = 18, face = "bold", hjust = 0.5),
    axis.title.x = element_text(size = 14),
    axis.title.y = element_text(size = 14),
    axis.text.x  = element_text(size = 12, angle = 45, hjust = 1),
    axis.text.y  = element_text(size = 12),
    legend.title = element_text(size = 12),
    legend.text  = element_text(size = 10),
    strip.placement = "outside",
    strip.text.x = element_text(size = 14, face = "bold"),
    strip.text.y = element_text(size = 14, face = "bold")
  )

ggsave(
  filename = file.path(output_dir, "Combined_Myocarditis_GBS_Thrombosis_FreeY_Clean.png"),
  plot = combined_plot_freeY_clean, width = 18, height = 20, bg = "white", dpi = 300
)
ggsave(
  filename = file.path(output_dir, "Combined_Myocarditis_GBS_Thrombosis_FreeY_Clean.pdf"),
  plot = combined_plot_freeY_clean, width = 18, height = 20, device = cairo_pdf
)
