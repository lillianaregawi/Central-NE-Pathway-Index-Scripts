library(dplyr)
library(readxl)
library(tidyverse)
library(knitr)
library(broom)
library(corrplot)
library(psych)
library(ggsignif)
library(htmlTable)
library(ggpubr)
library(rstatix) 



file_path <- file.choose()

data_sheets <- excel_sheets(file_path)
 
sheets = lapply(setNames(data_sheets, data_sheets), 
                function(x) read_excel(file_path, sheet=x))

data <- sheets$Data



#formatting data, removing NA values

data[data == "." | data == "^"] <- NA

#dataset with 89 PASC and 94 PASC in one file

combineddata <- data %>%
  filter(LP == 1) %>% # only people with lumbar punctures
  mutate_at(c("CSF_MHPG", "CSF_MHPG_DHPG", "CSF_DA", "CSF_HVA"), as.numeric) %>% 
  mutate(CSF_DA_DOPAC_HVA = CSF_DA + CSF_DOPAC + CSF_HVA)%>%
  relocate(CSF_DA_DOPAC_HVA, .after = CSF_HVA)
 

combineddata$Cohort <-  factor(combineddata$Cohort, levels=c('HV', 'PI-ME/CFS', 'PASC', 'PD'))

# dataset with 89 PASC and 94 PASC separated (94 PASC did not have PEM)

sepdata <- data %>%
  filter(LP == 1) %>%
  mutate_at(c("CSF_MHPG", "CSF_MHPG_DHPG", "CSF_DA", "CSF_HVA"), as.numeric) %>%
  mutate(CSF_DA_DOPAC_HVA = CSF_DA + CSF_DOPAC + CSF_HVA)%>%
  relocate(CSF_DA_DOPAC_HVA, .after = CSF_HVA)%>%
  mutate(ID = case_when(
    Group == 0 ~ "HV",
    Group == 2 ~ "PD",
    str_detect(SubjectID, "8900") & Group == 1 ~ "PASC 89",
    str_detect(SubjectID, "711") ~ "PASC 89",
    str_detect(Protocol, "94N") ~ "PASC 94",
    str_detect(SubjectID, "map") & Group == 1 ~ "PI-ME/CFS",
    TRUE ~ " "
  )) %>%
  relocate(ID, .after = Group)

sepdata$ID <-  factor(sepdata$ID, levels=c('HV', 'PI-ME/CFS', 'PASC 89',
                                           'PASC 94', 'PD'))



# Plots for levels of Catecholamines + Metabolites

# Doing it with and without PEM for pasc 


# fxn for n of each group of plot

stat_box_data <- function(y) {
  return( 
    data.frame(
      y = 0.5+1.1*max(y), 
      label = paste('n =', length(y), '\n'
                    #,'mean =', round(mean(y), 1), '\n' if we also want mean
      )
    )
  )
}


# no pem


#plot of CSF Norepinephrine Pathway

#anova

aov <- aov(CSF_NE_DHPG_MHPG ~ Cohort, data = combineddata) %>%
  tukey_hsd(conf.level = .95)
n1 <-ggboxplot(combineddata , x = "Cohort", y = "CSF_NE_DHPG_MHPG", outlier.shape = NA) +
  geom_jitter(aes(fill = Cohort),width = .2, height = 0, size = 4, shape = 21, stroke = 1 ) +
  stat_boxplot(geom ='errorbar', width = 0.4) +
  guides(y = guide_axis(minor.ticks = TRUE))+
  scale_fill_manual(values  = c("grey", "blue", "magenta",  "red"))+
  ylab("NE Pathway Index (pmol/mL)") + 
  theme(
    axis.title.x = element_blank(),
    axis.text.y = element_text(colour="black"),
    axis.minor.ticks.y.left = element_line(color = "black", size = 0.5),
    axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1),
    axis.text=element_text(size=20),
    axis.title =element_text(size=20),
    panel.background = element_rect(fill='transparent'),
    plot.background = element_rect(fill='transparent', color=NA),
    panel.grid.major = element_blank(),
    plot.margin = margin(2, 1, 1, 1, "cm"),
    panel.grid.minor = element_blank(),
    legend.background = element_rect(fill='transparent'),
    legend.box.background = element_rect(fill='transparent'),
    legend.title=element_blank()
  ) + theme(legend.position="none") +
  scale_y_continuous(breaks = seq(0, 110, by = 10),
                     minor_breaks = seq(0, 110, by = 5)) +
  stat_pvalue_manual(aov, label = "{scales::pvalue(p.adj, accuracy = 0.0001)}", hide.ns = TRUE,
                     y.position= c( 105, 110, 115, 120), size = 7) 


#plot of CSF Dopamine Pathway

# anova
aov2 <- aov(CSF_DA_DOPAC_HVA ~ Cohort, data = combineddata) %>%
  tukey_hsd(conf.level= 0.95)
d1 <- ggboxplot(combineddata , x = "Cohort", y = "CSF_DA_DOPAC_HVA",  outlier.shape = NA) +
  geom_jitter(aes(fill = Cohort),width = .2, height = 0, size = 4, shape = 21, stroke = 1 ) +
  stat_boxplot(geom ='errorbar', width = 0.4) +
  guides(y = guide_axis(minor.ticks = TRUE))+
  scale_fill_manual(values  = c("grey", "blue", "magenta",  "red"))+
  scale_y_continuous(breaks = seq(0, 500, by = 50)) +
  ylab("DA Pathway Index (pmol/mL)") + 
  theme(
    axis.title.x = element_blank(),
    axis.text.y = element_text(colour="black"),
    axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1),
    axis.text=element_text(size=20),
    axis.title =element_text(size=20),
    panel.background = element_rect(fill='transparent'),
    plot.background = element_rect(fill='transparent', color=NA),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    plot.margin = margin(2, 1, 1, 1, "cm"),
    legend.background = element_rect(fill='transparent'),
    legend.box.background = element_rect(fill='transparent'),
    legend.title=element_blank()
  )  + theme(legend.position="none") +
  stat_pvalue_manual(aov2, label = "{scales::pvalue(p.adj, accuracy = 0.0001)}", hide.ns = TRUE, , size = 7,
                     y.position= c( 407, 430, 467))


### sep data #####################################################


# CSF Norenephrine Pathway (Looking at PEM)

pemsepdata <- sepdata %>%
  filter(ID != "PASC 94") %>%
  mutate(ID =case_when(
    str_detect(SubjectID, "8900") & PEM == 1 ~ "PASC PEM",
    str_detect(SubjectID, "711") & PEM == 1 ~ "PASC PEM",
    str_detect(SubjectID, "8900") & PEM == 0 ~ "PASC No PEM",
    T ~ ID
    
  )) %>%
  filter(ID != "PASC 89") %>%
  filter(ID != "PI-ME/CFS" ) %>%
  filter(ID != "PD") 



pemsepdata$ID <-  factor(pemsepdata$ID, levels=c('HV', "PASC PEM", 'PASC No PEM'))


#plot of CSF NE Pathway 

#anova
aov3 <- aov(CSF_NE_DHPG_MHPG ~ ID, data = pemsepdata) %>%
  tukey_hsd(conf.level = .95)
group.colors <- c( "HV" = "grey", "PASC PEM" = "black", "PASC No PEM" ="white")
npem <-ggboxplot(pemsepdata , x = "ID", y = "CSF_NE_DHPG_MHPG", outlier.shape = NA) +
  geom_jitter(aes(fill = ID),width = .2, height = 0, size = 4, shape = 21, stroke = 1 ) +
  stat_boxplot(geom ='errorbar', width = 0.4) +
  guides(y = guide_axis(minor.ticks = TRUE))+
  ylab("NE Pathway Index (pmol/mL)") + 
  scale_fill_manual(values=group.colors)+
  theme(
    axis.title.x = element_blank(),
    axis.text.y = element_text(colour="black"),
    axis.text=element_text(size=20),
    axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1),
    axis.title =element_text(size=20),
    panel.background = element_rect(fill='transparent'),
    plot.background = element_rect(fill='transparent', color=NA),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    legend.background = element_rect(fill='transparent'),
    legend.box.background = element_rect(fill='transparent'),
    plot.margin = margin(2, 1, 1, 1, "cm"),
    legend.title=element_blank()
  ) + theme(legend.position="none") +
  scale_y_continuous(breaks = seq(0, 110, by = 10)) +
  stat_pvalue_manual(aov3, label = "{scales::pvalue(p.adj, accuracy = 0.0001)}", hide.ns = TRUE,
                     y.position= 105, size = 7)


ggarrange(n1,d1,npem, labels = c("A", "B", "C"), nrow = 1, font.label = list(size = 35),  vjust = 1.2)


