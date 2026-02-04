library(dplyr)
library(readxl)
library(tidyverse)
library(knitr)
library(broom)
library(corrplot)
library(psych)
library(htmlTable)
library(ggsignif)
library(ggpubr)
library(rstatix) 
library(Hmisc)
library(gt)

options(scipen = 999) 

data <- read_excel('/Users/aregawili/Library/CloudStorage/OneDrive-NationalInstitutesofHealth/Desktop/3-28-25 Catechol and Questionnaire Data .xlsx')

#formatting data

data[data == "." | data == "^"] <- NA

#dataset with 89 and 94 in one

combineddata <- data %>%
  filter(LP == 1) %>%
  mutate_at(c("CSF_MHPG", "CSF_MHPG_DHPG", "CSF_DA", "CSF_HVA"), as.numeric) %>%
  mutate(CSF_NE_DHPG_MHPG = CSF_NE + CSF_MHPG_DHPG) %>%
  relocate(CSF_NE_DHPG_MHPG, .after = CSF_MHPG_DHPG)%>%
  mutate(CSF_NE_DHPG = CSF_NE + CSF_DHPG) %>%
  relocate(CSF_NE_DHPG, .after = CSF_MHPG_DHPG)%>%
  mutate(CSF_DA_DOPAC_HVA = CSF_DA + CSF_DOPAC + CSF_HVA)%>%
  relocate(CSF_DA_DOPAC_HVA, .after = CSF_HVA)%>%
  mutate(CSF_TH = CSF_DA_DOPAC_HVA + CSF_NE_DHPG_MHPG + CSF_DOPA +CSF_Cys_DOPA)%>%
  relocate(CSF_TH, .after = CSF_DA_DOPAC_HVA)%>%
  mutate(ID = case_when(
    Group == 0 ~ "HV",
    Group == 2 ~ "PD",
    str_detect(SubjectID, "8900") & Group == 1 ~ "PASC",
    str_detect(SubjectID, "711") ~ "PASC",
    str_detect(Protocol, "94N") ~ "PASC",
    str_detect(SubjectID, "map") & Group == 1 ~ "PI-ME/CFS",
    TRUE ~ " "
  )) %>%
  relocate(ID, .after = Group)

combineddata$ID <-  factor(combineddata$ID, levels=c('HV', 'PI-ME/CFS', 'PASC', 'PD'))


# all participants

corr <- combineddata[1:90, ] %>%
  filter(!str_detect(SubjectID, "CNCS"))%>%
  filter(Protocol != "000094N")

corr <- corr[, c(1, 3, 11, 29:67, 83:97)]

# filter here for group:

#corr_all <- corr_all %>%
#filter(ID == "HV")

corr_all <-as.data.frame(corr[, 3:57])

# correlation

# all participants
corr_all_matrix <- corr.test(corr_all, use = "pairwise.complete.obs", method = "spearman", adjust = "none")
corr_all_CI <- as.data.frame(corr_all_matrix$ci)
corr_all_CI <- corr_all_CI[1:54, ] 
corr_all_CI <- corr_all_CI %>%
  filter(p <= 0.05)
corr_all_CI[,1:3] <- round(corr_all_CI[,1:3], digits = 2)
corr_all_CI[,4] <- signif(corr_all_CI[,4], digits = 2)


corr_all_CI$ID <- c("General Fatigue", "Physical Fatigue", "Reduced Activity", "MFI Total Score", "Fatigue",
                    "Sleep Related Impairments", "Physical Health", "Mental Health", "Symptom Severity",
                    "PDS Total Score", "Physical Functioning", "Physical Role Limitations", "General Health", 
                    "Vitality", "Social Functioning", "Mental Component", "Physical Component", "PHQ Total Score",
                    "Right Hand", "Left Hand")

corr_all_CI %>%
  gt(rowname_col = "ID") %>%
  tab_header(
    title = md("**Norepinephrine Pathway Correlations**"), 
    subtitle = md("All Participants (*n*=63)")
  ) %>%
  cols_label(
    lower = "Lower CI",
    upper = "Upper CI",
    p = md("*p*"),
    r = md("*r*")
  ) %>%
  opt_table_font(google_font(name = "Times New Roman")) %>%
  tab_options(
    table.width = pct(80)) %>%
  tab_style(
    style = cell_text(align = "center"),
    locations = cells_column_labels(columns = everything())) %>%
  cols_align(align = "center") %>%
  tab_row_group(
    label = "Multidimensional Fatigue Inventory",
    rows = 1:4
  ) %>%
  tab_row_group(
    label = "Patient-Reported Outcomes Measurement Information System",
    rows = 5:8
  ) %>%
  tab_row_group(
    label = "Polysymptomatic Distress Scale",
    rows = 9:10
  ) %>%
  tab_row_group(
    label = "36 Item Short-Form Survey",
    rows = 11:17
  ) %>%
  tab_row_group(
    label = "Patient Health Questionnaire",
    rows = 18
  ) %>%
  tab_row_group(
    label = "50% Threshold Handgrip Duration",
    rows = 19:20
  ) 


# *~*~~*~~**~*~*~*~*~*~~*~~~~*all patients*~*~~*~~**~*~*~*~*~*~~*~~~~*

# filter here for group:

corr_patients <- corr %>%
  filter(ID != "HV")

corr_patients <-as.data.frame(corr_patients[, 3:57])

# correlation

corr_patients_matrix <- corr.test(corr_patients, use = "pairwise.complete.obs", method = "spearman", adjust = "none")
corr_patients_CI <- as.data.frame(corr_patients_matrix$ci)
corr_patients_CI <- corr_patients_CI[1:54, ] 
corr_patients_CI <- corr_patients_CI %>%
  filter(p <= 0.05)
corr_patients_CI[,1:3] <- round(corr_patients_CI[,1:3], digits = 2)
corr_patients_CI[,4] <- signif(corr_patients_CI[,4], digits = 2)


corr_patients_CI$ID <- c("General Fatigue", "Physical Fatigue", "Fatigue", "General Health", 
                         "Vitality", "Social Functioning", "Emotional Role Limitations",
                         "Right Hand", "Left Hand")

corr_patients_CI %>%
  gt(rowname_col = "ID") %>%
  tab_header(
    title = md("**Norepinephrine Pathway Correlations**"), 
    subtitle = md("All Patients (*n*=38)")
  ) %>%
  cols_label(
    lower = "Lower CI",
    upper = "Upper CI",
    p = md("*p*"),
    r = md("*r*")
  ) %>%
  opt_table_font(google_font(name = "Times New Roman")) %>%
  tab_options(
    table.width = pct(80)) %>%
  tab_style(
    style = cell_text(align = "center"),
    locations = cells_column_labels(columns = everything())) %>%
  cols_align(align = "center") %>%
  tab_row_group(
    label = "Multidimensional Fatigue Inventory",
    rows = 1:2
  ) %>%
  tab_row_group(
    label = "Patient-Reported Outcomes Measurement Information System",
    rows = 3
  ) %>%
  tab_row_group(
    label = "36 Item Short-Form Survey",
    rows = 4:7
  ) %>%
  tab_row_group(
    label = "50% Threshold Handgrip Duration",
    rows = 8:9
  ) 

