# Profiling COVID-19 Vaccine Adverse Events Using VAERS Case Reports: What Has Changed Since 2021?

## Authors and Affiliations
- **Anna He**<sup>1,\*</sup>, **Katelyn Hur**<sup>2,\*</sup>, **Xingxian Li**<sup>3</sup>, **Jie Zheng**<sup>3</sup>, **Junguk Hur**<sup>4,$</sup>, **Yongqun He**<sup>3,$</sup>  
<sup>1</sup> Huron High School, Ann Arbor, Michigan  
<sup>2</sup> Red River High School, Grand Forks, North Dakota  
<sup>3</sup> University of Michigan, Ann Arbor, Michigan  
<sup>4</sup> University of North Dakota, Grand Forks, North Dakota  

**<sup>*</sup> Equal contribution**, **<sup>$</sup> Corresponding authors**  

---

## Abstract
**Background**: COVID-19 vaccines have been pivotal in reducing severe illness and death during the pandemic, but adverse events (AEs) remain a critical aspect of safety surveillance. In 2022, we reported the first systematic profiling of COVID-19 vaccine AEs using the Vaccine Adverse Event Reporting System (VAERS). Since then, vaccines have evolved with the introduction of bivalent formulations. This study provides an updated and expanded analysis to capture these evolving safety trends. <br>
**Methods**: Building upon our previous analysis of AE profiles, we systematically analyzed AE profiles for Pfizer, Moderna, and Janssen vaccines, along with the newer bivalent Pfizer and Moderna vaccines and the protein subunit Novavax vaccine, using VAERS data through June 28, 2024. <br>
**Results**: We observed a marked decrease in unique AEs reported for the Pfizer and Moderna vaccines, with an even greater reduction for Janssen. The bivalent versions of Pfizer and Moderna exhibited distinct AE profiles compared to their corresponding monovalent counterparts, though both remained generally safe. Significant differences were observed in thrombosis and myocarditis across the vaccines. Children aged 0-9 were more susceptible to AEs, and sex differences were evident. Females reported more common, general AEs, while males were more often linked to serious AEs, such as thrombosis, myocarditis, and Guillain-Barré syndrome. Ontology-based classification revealed that females were more likely to experience sensory-related AEs, whereas males were more prone to cardiovascular-related AEs. <br>
**Conclusion**: The current generation of COVID-19 vaccines appears to have safer profiles than earlier formulations, with new insights into age- and sex-specific AE patterns.  <br>

---

## Repository Contents
- **`1-1_convert_to_excel.py`** – Converts raw VAERS text output files into organized Excel spreadsheets, with one sheet per vaccine for easy downstream analysis.  
- **`4-1_vaxafe_massdownload_male_vs_female_age_v1.py`** – Automates bulk data downloads from the VaxAFE portal by vaccine, sex, and age range, and saves the results in structured Excel files for further statistical analysis.  
- **R Scripts** – Provide statistical analyses and visualizations. These can be loaded into **RStudio** and executed directly for reproducibility.

---

## Usage Examples

### **1. Converting raw VAERS output to Excel**
```bash
python 1-1_convert_to_excel.py
```
- **Input:** Raw `.txt` output files from VAERS stored in the `./vaxafe_output/` folder  
- **Output:** Cleaned `.xlsx` files saved in `./excel_output/` with each vaccine's data in separate sheets  

This script transforms unstructured VAERS data into structured Excel files for further analysis.

---

### **2. Downloading age- and sex-stratified VaxAFE data**
```bash
python 4-1_vaxafe_massdownload_male_vs_female_age_v1.py
```
- **Process:** Iterates through all combinations of vaccine, sex (male, female, any), and age groups (0–9, 10–19, etc.), submits queries to the VaxAFE portal, and downloads structured result files.  
- **Output:** 
  - Raw HTML query responses saved in `./vaxafe_male-vs-female_age/debug/`
  - Cleaned Excel files saved in `./vaxafe_male-vs-female_age/downloaded/`

This script automates what would otherwise be a manual, time-intensive download process.

---

### **3. Running R scripts**
- Open the desired R script in **RStudio**.  
- Click **Source** or run line-by-line to execute the workflow.  

These scripts perform data processing, statistical analyses, and generate visualizations to reproduce results and figures presented in the manuscript.

---
