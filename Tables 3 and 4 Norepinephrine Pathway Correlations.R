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


path <- "/Users/aregawili/Downloads" # change path to your own directory

setwd(path)

data_sheets <- excel_sheets('Supplementary File 1 Indiv. Data.xlsx')

sheets = lapply(setNames(data_sheets, data_sheets), 
                function(x) read_excel('Supplementary File 1 Indiv. Data.xlsx', sheet=x))

data <- sheets$Data

#formatting data

data[data == "." | data == "^"] <- NA

#dataset with 89 and 94 in one

combineddata <- data %>%
  filter(LP == 1) %>% # only people with lumbar punctures
  mutate_at(c("CSF_MHPG", "CSF_MHPG_DHPG", "CSF_DA", "CSF_HVA"), as.numeric) %>% 
  mutate(CSF_DA_DOPAC_HVA = CSF_DA + CSF_DOPAC + CSF_HVA)%>%
  relocate(CSF_DA_DOPAC_HVA, .after = CSF_HVA)
 

combineddata$Cohort <-  factor(combineddata$Cohort, levels=c('HV', 'PI-ME/CFS', 'PASC', 'PD'))


# all participants

corr <- combineddata[1:90, ] %>%
  filter(!str_detect(SubjectID, "CNCS"))%>%
  filter(Protocol != "000094N")

corr <- corr[, c(1, 7, 16, 27:65, 81:95)]

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
  filter(Cohort != "HV")

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

