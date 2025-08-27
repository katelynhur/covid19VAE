# Profiling COVID-19 Vaccine Adverse Events Using VAERS Case Reports: What Has Changed Since 2021?

## Authors and Affiliations
- **Anna He**<sup>1,*</sup>, **Katelyn Hur**<sup>2,*</sup>, **Xingxian Li**<sup>3</sup>, **Jie Zheng**<sup>3</sup>, **Junguk Hur**<sup>4,$</sup>, **Yongqun He**<sup>3,$</sup>  
<sup>1</sup> Huron High School, Ann Arbor, Michigan  
<sup>2</sup> Red River High School, Grand Forks, North Dakota  
<sup>3</sup> University of Michigan, Ann Arbor, Michigan  
<sup>4</sup> University of North Dakota, Grand Forks, North Dakota  

**<sup>*</sup> Equal contribution**, **<sup>$</sup> Corresponding authors**  

---

## Abstract
COVID-19 vaccines have been pivotal in reducing severe illness and death during the pandemic, but adverse events (AEs) remain a critical component of safety surveillance. In 2022, we published the first systematic profiling of COVID-19 vaccine AEs using the VAERS database. Since then, new vaccine formulations, including bivalent versions, have been introduced. This updated analysis incorporates VAERS data through June 2024 to provide a comprehensive assessment of evolving safety trends. We found that the overall number and severity of AEs decreased over time for Pfizer and Moderna vaccines, while Janssen showed a smaller reduction. Importantly, bivalent formulations demonstrated safer profiles, with no significant associations with rare but serious AEs such as myocarditis, thrombosis, or Guillain-Barré syndrome. Stratified analyses revealed that females were more likely to report common, non-serious AEs, whereas males were more frequently associated with severe, clinically significant AEs, particularly those affecting the cardiovascular system. These findings highlight the importance of ongoing, stratified safety monitoring to guide vaccine recommendations and risk communication.

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
