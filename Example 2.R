############################Rmarkdown and loopcode#########################################

#Criteria


#load packages----

#clear environ rm(list=ls())

library(scales)
library (reshape2)
library (grDevices)
library(tidyr)
library(dplyr) #for data manipulation
library(stringr)
library(readxl) #reading excel docs
library(readr) #reading r
library(tidyverse)
#library (ggplot2) #for plotting
#library (plotly)
library (lattice)

#load data ----


d4 <- read_excel('data_3.xlsx', sheet = 'sheet_name') 



#1617 analysis----

#rename columns
names(d4)[names(d4)=="TFSM6CLA1A"] <- "disadv1617"
names(d4)[names(d4)=="TNOTFSM6CLA1A"] <- "non_disadv"
names(d4)[names(d4)=="PTFSM6CLA1ABASICS_95"] <- "dis_5e_m1617"
names(d4)[names(d4)=="PTNOTFSM6CLA1ABASICS_95"] <- "nondis_5e_m1617"
names(d4)[names(d4)=="P8MEA_FSM6CLA1A"] <- "dis_p8_1617"
names(d4)[names(d4)=="P8MEA_NFSM6CLA1A"] <- "nondis_p8_1617"
names(d4)[names(d4)=="PTEBACC_EFSM6CLA1A_PTQ_EE"] <- "dis_ebacc_ent_1617"
names(d4)[names(d4)=="PTEBACC_ENFSM6CLA1A_PTQ_EE"] <- "nondis_ebacc_ent_1617"
names(d4)[names(d4)=="NFTYPE3"] <- "institution_type"
names(d4)[names(d4)=="ADMPOL2"] <- "addmission_type"


#convert to numeric values

d4$disadv1617 <- as.numeric(d4$disadv1617)
d4$dis_5e_m1617 <- as.numeric(d4$dis_5e_m1617)
d4$nondis_5e_m1617  <- as.numeric(d4$nondis_5e_m1617) 
d4$dis_p8_1617 <- as.numeric(d4$dis_p8_1617)
d4$nondis_p8_1617 <- as.numeric(d4$nondis_p8_1617)
d4$dis_ebacc_ent <- as.numeric(d4$pro_8_dis_1617)
d4$dis_ebacc_ent_1617 <- as.numeric(d4$dis_ebacc_ent_1617)
d4$nondis_ebacc_ent_1617 <- as.numeric(d4$nondis_ebacc_ent_1617)
d4$OFSTED  <- as.numeric(d4$OFSTED)
d4$OFSTED2  <- as.numeric(d4$OFSTED2)


d4 %>% 
  select(URN, LAESTAB,  Region, Closed, SCHNAME, addmission_type, institution_type, ADDRESS1, ADDRESS2, ADDRESS3, TOWN, 
         PCODE, OFSTED, TELNUM, disadv1617,  
         dis_5e_m1617, nondis_5e_m1617, dis_p8_1617, nondis_p8_1617, dis_ebacc_ent_1617, nondis_ebacc_ent_1617) %>%
  filter (dis_5e_m1617 ==1, nondis_5e_m1617 ==1, dis_ebacc_ent_1617 ==1, nondis_ebacc_ent_1617 ==1, 
          dis_p8_1617==1, nondis_p8_1617 ==1) -> ks4top100

ks4top100 %>%
  filter(URN!="NA") -> ks4top100



#LOOP----



for (schs in unique(ks4top100$URN)){
  rmarkdown::render('Example 2_letter.Rmd',  
                    output_file =  paste(schs, " Letter", ".docx", sep=''), 
                    output_dir = 'Exports')}



write.csv(ks4top100, file="Top KS4 Schools.csv", row.names = FALSE)



