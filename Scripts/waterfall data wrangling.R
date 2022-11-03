library(tidyverse)
library(openxlsx)
library(readxl)
library(reshape2)
library(data.table)


###############################################################################################

#path for CoT file
CoT_path <- file.path(fldr, CoT_Previous)

#load in CoT Dashboard
df_cot <- read_xlsx(path=CoT_path, sheet = 'Waterfall Data')

#path for source file
path_in <- file.path(fldr, source_file)

# Create list of excel sheets and turn each into separate dataframes
sheet = excel_sheets(path_in)

mylist = lapply(setNames(sheet, sheet), function(x) read_excel(path_in, sheet=x))
names(mylist) <- sheet
list2env(mylist ,.GlobalEnv)



#reshape each df so column headers are in the right place
names(`TX_NEW `) <- `TX_NEW `[4,]
TX_NEW <- `TX_NEW `[-c(1:4),]

names(TX_CURR) <- TX_CURR[4,]
TX_CURR <- TX_CURR[-c(1:4),]


(hdr1_ML <- read_excel(path_in,
                    sheet = "TX_ML",
                    skip = 3,
                    n_max = 1,
                    .name_repair = "minimal") %>% 
    names())
(hdr2_ML <- read_excel(path_in,
                    sheet = "TX_ML",
                    skip = 4,
                    n_max = 1,
                    .name_repair = "minimal") %>% 
    names())
(clean_hdrs_ML <- tibble(hdr1_ML, hdr2_ML)%>% 
    mutate(hdr1_ML = na_if(hdr1_ML, "")) %>% 
    fill(hdr1_ML) %>% 
    unite(hdr, c(hdr1_ML, hdr2_ML)) %>% 
    pull())

#read in data frame skipping the header rows
(TX_ML <- read_excel(path_in,
                  sheet = "TX_ML",
                  range = cell_limits(c(5, 1), c(NA, NA)),
                  col_names = clean_hdrs_ML))

#delete extra title text from indicators
names(TX_ML)[-1] <- sub("(.*:[^:]+).*:", "", names(TX_ML)[-1])

#delete NAs added to the other cols
names(TX_ML) = gsub("NA_", "", x = names(TX_ML))

#remove first row
TX_ML <- TX_ML[-1,]

#process needs to be repeated for TX_RTT

(hdr1_RTT <- read_excel(path_in,
                    sheet = "TX_RTT",
                    skip = 1,
                    n_max = 1,
                    .name_repair = "minimal") %>% 
    names())
(hdr2_RTT <- read_excel(path_in,
                    sheet = "TX_RTT",
                    skip = 2,
                    n_max = 1,
                    .name_repair = "minimal") %>% 
    names())
(clean_hdrs_RTT <- tibble(hdr1_RTT, hdr2_RTT)%>% 
    mutate(hdr1_RTT = na_if(hdr1_RTT, "")) %>% 
    fill(hdr1_RTT) %>% 
    unite(hdr, c(hdr1_RTT, hdr2_RTT)) %>% 
    pull())

#read in data frame skipping the header rows
(TX_RTT <- read_excel(path_in,
                     sheet = "TX_RTT",
                     range = cell_limits(c(5, 1), c(NA, NA)),
                     col_names = clean_hdrs_RTT))

#remove irrelevant column strings for columns that contain MER
names(TX_RTT)[-1] <- ifelse(grepl("MER:", names(TX_RTT)[-1]), sub(".*_", "", names(TX_RTT)[-1]), names(TX_RTT)[-1])

#delete NAs added to the other cols
names(TX_RTT) = gsub("NA_", "", x = names(TX_RTT))


#---------------------------------TX_NEW---------------------------------
#drop columns with Unknown string
TX_NEW <- TX_NEW [, -grep("Unknown", colnames(TX_NEW))]

#choose columns 
TX_NEW_df <-TX_NEW %>% 
  select(c("DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
                     "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male")))) 
#concat age cols
TX_NEW_df <- TX_NEW_df %>%
  mutate_at(c(10:39), as.numeric)

TX_NEW_df <- TX_NEW_df %>%
  mutate("Female\r\n01-09yrs"=rowSums(TX_NEW_df[11:12])) %>%
  relocate("Female\r\n01-09yrs", .after="Female\r\n5-9yrs") %>%
  mutate("Female\r\n40-49yrs"=rowSums(TX_NEW_df[19:20])) %>%
  relocate("Female\r\n40-49yrs", .after="Female\r\n45-49 yrs") %>%
  mutate("Male\r\n01-09yrs"=rowSums(TX_NEW_df[26:27])) %>%
  relocate("Male\r\n01-09yrs", .after="Male\r\n5-9yrs") %>%
  mutate("Male\r\n40-49yrs"=rowSums(TX_NEW_df[34:35])) %>%
  relocate("Male\r\n40-49yrs", .after="Male\r\n45-49 yrs") %>%
  select(-c("Male", "Female", "Female\r\n1-4yrs", "Female\r\n5-9yrs", "Female\r\n40-44 yrs", "Female\r\n45-49 yrs",
                       "Male\r\n1-4yrs", "Male\r\n5-9yrs", "Male\r\n40-44 yrs", "Male\r\n45-49 yrs"))

#rename columns
TX_NEW_df <- TX_NEW_df %>% 
  rename("psnu" = "DHIS2 District", "psnuuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
          "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support", "period" = "Period", "Female\r\n<01yr" = "Female\r\n<1yr", "Male\r\n<01yr" = "Male\r\n<1yr", 
         "Female\r\n50+yrs" = "Female\r\n 50+ yrs" , "Male\r\n50+yrs"  = "Male\r\n 50+ yrs")

#transpose columns from wide to long
TX_NEW_df <- pivot_longer(TX_NEW_df, contains(c("Female", "Male")), names_to = "age", values_to = "TX_NEW_Now_R")
TX_NEW_df$age <- trimws(TX_NEW_df$age, which = c("right"))

#add Sex column
TX_NEW_df <- TX_NEW_df %>%
  mutate(sex="NA",
         .before="indicatortype")
TX_NEW_df <- TX_NEW_df %>%
  relocate(age, .before = sex)

#fill Sex column with Male or Female
TX_NEW_df$sex <- ifelse(grepl("Female", TX_NEW_df$age), "Female","Male")

#delete male and female strings from age column
TX_NEW_df <- TX_NEW_df %>%
  mutate_at("age", str_replace, c("Male"), "")
#delete male and female strings from age column
TX_NEW_df <- TX_NEW_df %>%
  mutate_at("age", str_replace, c("Female"), "")
#add white space for column values with no space between number and yr/yrs   
TX_NEW_df$age <- sub("([0-9])([y])", "\\1 \\2", TX_NEW_df$age)

#delete yrs and yr strings from age column
TX_NEW_df <- TX_NEW_df %>%
mutate_at("age", str_replace, "yrs", "")
TX_NEW_df <- TX_NEW_df %>%
  mutate_at("age", str_replace, "yr", "")

#add age_type column
TX_NEW_df <- TX_NEW_df %>%
  mutate(age_type=NA,
         .before="age")

#insert appropriate values for age_type
coarse_values <- "^<15|^15+"
TX_NEW_df_final <- TX_NEW_df %>% 
  mutate(age_type=ifelse(grepl(coarse_values, TX_NEW_df$age), "trendscoarse","trendsfine"))
#revisit
#TX_NEW_df_final <- TX_NEW_df %>% 
#  mutate(age_type = ifelse(age == '15-19', 'trendsfine' ,age_type))


#---------------------------------TX_CURR---------------------------------



#remove rows with all NAS, and strings: (Unknown, ARVs, Sex Workers) 
TX_CURR <- TX_CURR[,colSums(is.na(TX_CURR))<nrow(TX_CURR)]
TX_CURR <- TX_CURR [, -grep("Unknown", colnames(TX_CURR))]
TX_CURR <- TX_CURR [, -grep("ARVs", colnames(TX_CURR))]
TX_CURR <- TX_CURR [, -grep("Sex Workers", colnames(TX_CURR))]


#choose columns
TX_CURR_df <- TX_CURR %>%
  select(c("DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#delete and Male and Female cols
TX_CURR_df = select(TX_CURR_df, -c("Male", "Female", "<15 Years (<20 kg), Male",  "<15 Years (20+ kg), Male",
                                   "<15 Years (<20 kg), Female", "<15 Years (20+ kg), Female", "15+ Years, Female", "15+ Years, Male"))
                                   
#concat age cols
TX_CURR_df <- TX_CURR_df %>%
  mutate_at(c(10:43), as.numeric)

TX_CURR_df <- TX_CURR_df %>%
  mutate("Female\r\n01-09yrs"=rowSums(TX_CURR_df[11:12])) %>%
  relocate("Female\r\n01-09yrs", .after="Female\r\n5-9yrs") %>%
  mutate("Female\r\n40-49yrs"=rowSums(TX_CURR_df[19:20])) %>%
  relocate("Female\r\n40-49yrs", .after="Female\r\n45-49 yrs") %>%
  mutate("Female\r\n50+yrs"=rowSums(TX_CURR_df[21:24])) %>%
  relocate("Female\r\n50+yrs", .after="Female\r\n 65+ yrs") %>%
  mutate("Male\r\n01-09yrs"=rowSums(TX_CURR_df[28:29])) %>%
  relocate("Male\r\n01-09yrs", .after="Male\r\n5-9yrs") %>%
  mutate("Male\r\n40-49yrs"=rowSums(TX_CURR_df[36:37])) %>%
  relocate("Male\r\n40-49yrs", .after="Male\r\n45-49 yrs") %>%
  mutate("Male\r\n50+yrs"=rowSums(TX_CURR_df[38:41])) %>%
  relocate("Male\r\n50+yrs", .after="Male\r\n 65+ yrs") %>%
  select(-c("Female\r\n1-4yrs", "Female\r\n5-9yrs", "Female\r\n40-44 yrs", "Female\r\n45-49 yrs",
           "Female\r\n 50-54 yrs", "Female\r\n 55-59 yrs", "Female\r\n 60-64 yrs", "Female\r\n 65+ yrs",
           "Male\r\n1-4yrs", "Male\r\n5-9yrs", "Male\r\n40-44 yrs", "Male\r\n45-49 yrs", "Male\r\n 50-54 yrs",
           "Male\r\n 55-59 yrs", "Male\r\n 60-64 yrs", "Male\r\n 65+ yrs"))

#rename columns
TX_CURR_df <- TX_CURR_df %>% 
  rename("psnu" = "DHIS2 District", "psnuuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support", "period" = "Period", "Female\r\n<01yr" = "Female\r\n<1yr", "Male\r\n<01yr" = "Male\r\n<1yr")

#transpose columns from wide to long
TX_CURR_df <- pivot_longer(TX_CURR_df, contains(c("Female", "Male")), names_to = "age", values_to = "TX_CURR_Now_R")
TX_CURR_df$age <- trimws(TX_CURR_df$age, which = c("right"))

#add Sex column
TX_CURR_df <- TX_CURR_df %>% 
  mutate(sex="NA",
         .before="indicatortype")
TX_CURR_df <- TX_CURR_df %>%
  relocate(age, .before = sex)

#fill Sex column with Male or Female
TX_CURR_df$sex <- ifelse(grepl("Female", TX_CURR_df$age), "Female","Male")

#delete male and female strings from age column
TX_CURR_df <- TX_CURR_df %>%
  mutate_at("age", str_replace, "Female", "")
TX_CURR_df <- TX_CURR_df %>%
  mutate_at("age", str_replace, "Male", "")

#add white space for column values with no space between number and yr/yrs   
TX_CURR_df$age <- sub("([0-9])([y])", "\\1 \\2", TX_CURR_df$age)

#delete yrs and yr strings from age column
TX_CURR_df <- TX_CURR_df %>%
  mutate_at("age", str_replace, "yrs", "")
TX_CURR_df <- TX_CURR_df %>%
  mutate_at("age", str_replace, "yr", "")

#add age_type column
TX_CURR_df <- TX_CURR_df %>%
  mutate(age_type=NA,
         .before="age")

#insert appropriate values for age_type
coarse_values <- "^<15|^15+"
TX_CURR_df_final <- TX_CURR_df %>% 
  mutate(age_type=ifelse(grepl(coarse_values, TX_CURR_df$age), "trendscoarse","trendsfine"))
#TX_CURR_df_final<-  TX_CURR_df %>% 
#  mutate(age_type = ifelse(age == '15-19', 'trendsfine' ,age_type))

#---------------------------------TX_ML---------------------------------

#Deleting unnecessary columns
TX_ML <- TX_ML [, -grep("Unknown", colnames(TX_ML))]
TX_ML <- TX_ML [, -grep("Numerator:", colnames(TX_ML))]
TX_ML <- TX_ML [, -grep("resulting in", colnames(TX_ML))]
TX_ML <- TX_ML [, -grep("(FSW)", colnames(TX_ML))]
TX_ML <- TX_ML [, -grep("(MSM)", colnames(TX_ML))]

#choose columns
TX_ML_df <-TX_ML %>% 
  select(c("DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#concat age cols
TX_ML_df <- TX_ML_df %>%
  mutate_at(c(10:153), as.numeric)

#Died by
TX_ML_df <- TX_ML_df %>%
  mutate(" Died by Age/Sex_01 - 09 years, Female"=rowSums(TX_ML_df[11:12])) %>%
  relocate(" Died by Age/Sex_01 - 09 years, Female", .after=" Died by Age/Sex_(5 - 9) Years, Female") %>%
  mutate(" Died by Age/Sex_40 - 49 years, Female"=rowSums(TX_ML_df[19:20])) %>%
  relocate(" Died by Age/Sex_40 - 49 years, Female", .after=" Died by Age/Sex_45 - 49 Years, Female") %>%
  mutate(" Died by Age/Sex_01 - 09 years, Male"=rowSums(TX_ML_df[83:84])) %>%
  relocate(" Died by Age/Sex_01 - 09 years, Male", .after=" Died by Age/Sex_(5 - 9) Years, Male") %>%
  mutate(" Died by Age/Sex_40 - 49 years, Male"=rowSums(TX_ML_df[91:92])) %>%
  relocate(" Died by Age/Sex_40 - 49 years, Male", .after=" Died by Age/Sex_45 - 49 Years, Male") %>%
  select(-c(" Died by Age/Sex_1­ - 4 years, Female", " Died by Age/Sex_(5 - 9) Years, Female", 
            " Died by Age/Sex_40 - 44 Years, Female", " Died by Age/Sex_45 - 49 Years, Female",
            " Died by Age/Sex_1­ - 4 years, Male", " Died by Age/Sex_(5 - 9) Years, Male",
            " Died by Age/Sex_40 - 44 Years, Male", " Died by Age/Sex_45 - 49 Years, Male"))
#Transferred out
TX_ML_df <- TX_ML_df %>%  
  mutate("  transferred out by Age/Sex_01 - 09 years, Female"=rowSums(TX_ML_df[21:22])) %>%
  relocate("  transferred out by Age/Sex_01 - 09 years, Female", .after="  transferred out by Age/Sex_(5 - 9) Years, Female") %>%
  mutate("  transferred out by Age/Sex_40 - 49 Years, Female"=rowSums(TX_ML_df[29:30])) %>%
  relocate("  transferred out by Age/Sex_40 - 49 Years, Female", .after="  transferred out by Age/Sex_45 - 49 Years, Female") %>%
  mutate("  transferred out by Age/Sex_01 - 09 years, Male"=rowSums(TX_ML_df[91:92])) %>%
  relocate("  transferred out by Age/Sex_01 - 09 years, Male", .after="  transferred out by Age/Sex_(5 - 9) Years, Male") %>%
  mutate("  transferred out by Age/Sex_40 - 49 Years, Male"=rowSums(TX_ML_df[99:100])) %>%
  relocate("  transferred out by Age/Sex_40 - 49 Years, Male", .after="  transferred out by Age/Sex_45 - 49 Years, Male") %>%
  select(-c("  transferred out by Age/Sex_1­ - 4 years, Female", "  transferred out by Age/Sex_(5 - 9) Years, Female",
            "  transferred out by Age/Sex_40 - 44 Years, Female", "  transferred out by Age/Sex_45 - 49 Years, Female",
            "  transferred out by Age/Sex_1­ - 4 years, Male", "  transferred out by Age/Sex_(5 - 9) Years, Male",
            "  transferred out by Age/Sex_40 - 44 Years, Male", "  transferred out by Age/Sex_45 - 49 Years, Male"))
#IIT <3 months
TX_ML_df <- TX_ML_df %>% 
  mutate(" IIT After being on Treatment for <3 month by Age/Sex_01 - 09 years, Female"=rowSums(TX_ML_df[31:32])) %>%
  relocate(" IIT After being on Treatment for <3 month by Age/Sex_01 - 09 years, Female", .after=" IIT After being on Treatment for <3 month by Age/Sex_(5 - 9) Years, Female") %>%
  mutate(" IIT After being on Treatment for <3 month by Age/Sex_40 - 49 Years, Female"=rowSums(TX_ML_df[39:40])) %>%
  relocate(" IIT After being on Treatment for <3 month by Age/Sex_40 - 49 Years, Female", .after=" IIT After being on Treatment for <3 month by Age/Sex_45 - 49 Years, Female") %>%
  mutate(" IIT After being on Treatment for <3 month by Age/Sex_01 - 09 years, Male"=rowSums(TX_ML_df[99:100])) %>%
  relocate(" IIT After being on Treatment for <3 month by Age/Sex_01 - 09 years, Male", .after=" IIT After being on Treatment for <3 month by Age/Sex_(5 - 9) Years, Male") %>%
  mutate(" IIT After being on Treatment for <3 month by Age/Sex_40 - 49 Years, Male"=rowSums(TX_ML_df[107:108])) %>%
  relocate(" IIT After being on Treatment for <3 month by Age/Sex_40 - 49 Years, Male", .after=" IIT After being on Treatment for <3 month by Age/Sex_45 - 49 Years, Male") %>%
  select(-c(" IIT After being on Treatment for <3 month by Age/Sex_1­ - 4 years, Female", " IIT After being on Treatment for <3 month by Age/Sex_(5 - 9) Years, Female",
            " IIT After being on Treatment for <3 month by Age/Sex_40 - 44 Years, Female", " IIT After being on Treatment for <3 month by Age/Sex_45 - 49 Years, Female",
            " IIT After being on Treatment for <3 month by Age/Sex_1­ - 4 years, Male", " IIT After being on Treatment for <3 month by Age/Sex_(5 - 9) Years, Male",
            " IIT After being on Treatment for <3 month by Age/Sex_40 - 44 Years, Male", " IIT After being on Treatment for <3 month by Age/Sex_45 - 49 Years, Male"))
#IIT 3-5 months
TX_ML_df <- TX_ML_df %>% 
  mutate(" IIT After being on Treatment for 3-5 months by Age/Sex_01 - 09 years, Female"=rowSums(TX_ML_df[41:42])) %>%
  relocate(" IIT After being on Treatment for 3-5 months by Age/Sex_01 - 09 years, Female", .after=" IIT After being on Treatment for 3-5 months by Age/Sex_(5 - 9) Years, Female") %>%
  mutate(" IIT After being on Treatment for 3-5 months by Age/Sex_40 - 49 Years, Female"=rowSums(TX_ML_df[49:50])) %>%
  relocate(" IIT After being on Treatment for 3-5 months by Age/Sex_40 - 49 Years, Female", .after=" IIT After being on Treatment for 3-5 months by Age/Sex_45 - 49 Years, Female") %>%
  mutate(" IIT After being on Treatment for 3-5 months by Age/Sex_01 - 09 years, Male"=rowSums(TX_ML_df[107:108])) %>%
  relocate(" IIT After being on Treatment for 3-5 months by Age/Sex_01 - 09 years, Male", .after=" IIT After being on Treatment for 3-5 months by Age/Sex_(5 - 9) Years, Male") %>%
  mutate(" IIT After being on Treatment for 3-5 months by Age/Sex_40 - 49 Years, Male"=rowSums(TX_ML_df[115:116])) %>%
  relocate(" IIT After being on Treatment for 3-5 months by Age/Sex_40 - 49 Years, Male", .after=" IIT After being on Treatment for 3-5 months by Age/Sex_45 - 49 Years, Male") %>%
  select(-c(" IIT After being on Treatment for 3-5 months by Age/Sex_1­ - 4 years, Female", " IIT After being on Treatment for 3-5 months by Age/Sex_(5 - 9) Years, Female",
            " IIT After being on Treatment for 3-5 months by Age/Sex_40 - 44 Years, Female", " IIT After being on Treatment for 3-5 months by Age/Sex_45 - 49 Years, Female",
            " IIT After being on Treatment for 3-5 months by Age/Sex_1­ - 4 years, Male"," IIT After being on Treatment for 3-5 months by Age/Sex_(5 - 9) Years, Male",
            " IIT After being on Treatment for 3-5 months by Age/Sex_40 - 44 Years, Male", " IIT After being on Treatment for 3-5 months by Age/Sex_45 - 49 Years, Male"))

#IIT 6+ months
TX_ML_df <- TX_ML_df %>% 
  mutate(" IIT After being on Treatment for 6+ months by Age/Sex_01 - 09 years, Female"=rowSums(TX_ML_df[51:52])) %>%
  relocate(" IIT After being on Treatment for 6+ months by Age/Sex_01 - 09 years, Female", .after=" IIT After being on Treatment for 6+ months by Age/Sex_(5 - 9) Years, Female") %>%
  mutate(" IIT After being on Treatment for 6+ months by Age/Sex_40 - 49 Years, Female"=rowSums(TX_ML_df[59:60])) %>%
  relocate(" IIT After being on Treatment for 6+ months by Age/Sex_40 - 49 Years, Female", .after=" IIT After being on Treatment for 6+ months by Age/Sex_45 - 49 Years, Female") %>%
  mutate(" IIT After being on Treatment for 6+ months by Age/Sex_01 - 09 years, Male"=rowSums(TX_ML_df[115:116])) %>%
  relocate(" IIT After being on Treatment for 6+ months by Age/Sex_01 - 09 years, Male", .after=" IIT After being on Treatment for 6+ months by Age/Sex_(5 - 9) Years, Male") %>%
  mutate(" IIT After being on Treatment for 6+ months by Age/Sex_40 - 49 Years, Male"=rowSums(TX_ML_df[123:124])) %>%
  relocate(" IIT After being on Treatment for 6+ months by Age/Sex_40 - 49 Years, Male", .after=" IIT After being on Treatment for 6+ months by Age/Sex_45 - 49 Years, Male") %>%
  select(-c(" IIT After being on Treatment for 6+ months by Age/Sex_1­ - 4 years, Female", " IIT After being on Treatment for 6+ months by Age/Sex_(5 - 9) Years, Female",
            " IIT After being on Treatment for 6+ months by Age/Sex_40 - 44 Years, Female", " IIT After being on Treatment for 6+ months by Age/Sex_45 - 49 Years, Female",
            " IIT After being on Treatment for 6+ months by Age/Sex_1­ - 4 years, Male"," IIT After being on Treatment for 6+ months by Age/Sex_(5 - 9) Years, Male",
            " IIT After being on Treatment for 6+ months by Age/Sex_40 - 44 Years, Male", " IIT After being on Treatment for 6+ months by Age/Sex_45 - 49 Years, Male"))

#Refused Treatment 
TX_ML_df <- TX_ML_df %>% 
  mutate(" Refused (Stopped) Treatment by Age/Sex_01 - 09 years, Female"=rowSums(TX_ML_df[61:62])) %>%
  relocate(" Refused (Stopped) Treatment by Age/Sex_01 - 09 years, Female", .after=" Refused (Stopped) Treatment by Age/Sex_(5 - 9) Years, Female") %>%
  mutate(" Refused (Stopped) Treatment by Age/Sex_40 - 49 Years, Female"=rowSums(TX_ML_df[69:70])) %>%
  relocate(" Refused (Stopped) Treatment by Age/Sex_40 - 49 Years, Female", .after=" Refused (Stopped) Treatment by Age/Sex_45 - 49 Years, Female") %>%
  mutate(" Refused (Stopped) Treatment by Age/Sex_01 - 09 years, Male"=rowSums(TX_ML_df[123:124])) %>%
  relocate(" Refused (Stopped) Treatment by Age/Sex_01 - 09 years, Male", .after=" Refused (Stopped) Treatment by Age/Sex_(5 - 9) Years, Male") %>%
  mutate(" Refused (Stopped) Treatment by Age/Sex_40 - 49 Years, Male"=rowSums(TX_ML_df[131:132])) %>%
  relocate(" Refused (Stopped) Treatment by Age/Sex_40 - 49 Years, Male", .after=" Refused (Stopped) Treatment by Age/Sex_45 - 49 Years, Male") %>%
  select(-c(" Refused (Stopped) Treatment by Age/Sex_1­ - 4 years, Female", " Refused (Stopped) Treatment by Age/Sex_(5 - 9) Years, Female",
            " Refused (Stopped) Treatment by Age/Sex_40 - 44 Years, Female", " Refused (Stopped) Treatment by Age/Sex_45 - 49 Years, Female",
            " Refused (Stopped) Treatment by Age/Sex_1­ - 4 years, Male", " Refused (Stopped) Treatment by Age/Sex_(5 - 9) Years, Male",
            " Refused (Stopped) Treatment by Age/Sex_40 - 44 Years, Male", " Refused (Stopped) Treatment by Age/Sex_45 - 49 Years, Male"))



#change line spacing 
names(TX_ML_df)[-1] <- sub("\\r\\n", " ", names(TX_ML_df)[-1])
names(TX_ML_df)

TX_ML_df <- TX_ML_df %>% 
  rename(" Died by Age/Sex_<01 year, Female" = " Died by Age/Sex_< 1 year, Female", "  transferred out by Age/Sex_<01 year, Female" = 
           "  transferred out by Age/Sex_< 1 year, Female", " IIT After being on Treatment for <3 month by Age/Sex_<01 year, Female" =
           " IIT After being on Treatment for <3 month by Age/Sex_< 1 year, Female", " IIT After being on Treatment for 3-5 months by Age/Sex_<01 year, Female" =
           " IIT After being on Treatment for 3-5 months by Age/Sex_< 1 year, Female", " IIT After being on Treatment for 6+ months by Age/Sex_<01 year, Female" =
           " IIT After being on Treatment for 6+ months by Age/Sex_< 1 year, Female",  " Refused (Stopped) Treatment by Age/Sex_<01 year, Female" =
           " Refused (Stopped) Treatment by Age/Sex_< 1 year, Female", " Died by Age/Sex_<01 year, Male" = " Died by Age/Sex_< 1 year, Male",
           "  transferred out by Age/Sex_<01 year, Male" = "  transferred out by Age/Sex_< 1 year, Male", " IIT After being on Treatment for <3 month by Age/Sex_<01 year, Male" =
           " IIT After being on Treatment for <3 month by Age/Sex_< 1 year, Male", " IIT After being on Treatment for 3-5 months by Age/Sex_<01 year, Male" =
           " IIT After being on Treatment for 3-5 months by Age/Sex_< 1 year, Male", " IIT After being on Treatment for 6+ months by Age/Sex_<01 year, Male" =
           " IIT After being on Treatment for 6+ months by Age/Sex_< 1 year, Male", " Refused (Stopped) Treatment by Age/Sex_<01 year, Male" =
           " Refused (Stopped) Treatment by Age/Sex_< 1 year, Male")

#transpose columns from wide to long to semi-wide
TX_ML_df_pivots<- function(df, x)
{pivot_longer(df, contains(x),   
              names_to = c("disag", "age", "sex"),
              names_sep = "_|,", 
)
}

df_ML_long<- TX_ML_df_pivots(TX_ML_df, c("IIT", "Died by", "refused", "transferred")) 

df_ML_wider<-df_ML_long %>% 
  pivot_wider(names_from = disag, values_from=value)

#remove white space in age and sex column values
df_ML_wider$sex <- gsub('\\s+', '', df_ML_wider$sex)

#remove all unnecessary non numeric chars from age column
df_ML_wider$age <- gsub("[^<+0-9.-]", '', df_ML_wider$age)

#relocate age and sex
df_ML_wider <- df_ML_wider %>%
  relocate(age, .after = "DHIS2 HF Name") %>%
  relocate(sex, .after = "age")

#add age_type 
df_ML_wider <- df_ML_wider %>%
  mutate(age_type="trendsfine", .before='age') 

#rename columns
TX_ML_df_final <- df_ML_wider %>% 
  rename("psnu" = "DHIS2 District", "psnuuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support", "period" = "Period", "TX_ML_Interruption <3 Months Treatment_Now_R" =  " IIT After being on Treatment for <3 month by Age/Sex",
         "TX_ML_Interruption 3-5 Months Treatment_R" = " IIT After being on Treatment for 3-5 months by Age/Sex",
         "TX_ML_Interruption 6+ Months Treatment_R" = " IIT After being on Treatment for 6+ months by Age/Sex", "TX_ML_Died_Now_R" = " Died by Age/Sex",
         "TX_ML_Refused Stopped Treatment_Now_R" = " Refused (Stopped) Treatment by Age/Sex", "TX_ML_Transferred Out_Now_R" = "  transferred out by Age/Sex" )

#---------------------------------TX_RTT---------------------------------
#remove unnecessary rows from disag df
TX_RTT <- TX_RTT [, -grep("KP", colnames(TX_RTT))]
TX_RTT <- TX_RTT [, -grep("IIT", colnames(TX_RTT))]
TX_RTT <- TX_RTT [, -grep("MSM", colnames(TX_RTT))]
TX_RTT <- TX_RTT [, -grep("Sex Workers", colnames(TX_RTT))]
TX_RTT <- TX_RTT [, -grep("Sub Total", colnames(TX_RTT))]

#choose columns
TX_RTT_df <-TX_RTT %>% 
  select(c("DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#concat age cols
TX_RTT_df <- TX_RTT_df %>%
  mutate_at(c(10:39), as.numeric)

TX_RTT_df <- TX_RTT_df %>%
  mutate("01-09 Years, Female"=rowSums(TX_RTT_df[11:12])) %>%
  relocate("01-09 Years, Female", .after="(5 - 9) Years, Female") %>%
  mutate("40-49 Years, Female"=rowSums(TX_RTT_df[19:20])) %>%
  relocate("40-49 Years, Female", .after="45 - 49 Years, Female") %>%
  mutate("01-09 Years, Male"=rowSums(TX_RTT_df[26:27])) %>%
  relocate("01-09 Years, Male", .after="(5 - 9) Years, Male") %>%
  mutate("40-49 Years, Male"=rowSums(TX_RTT_df[34:35])) %>%
  relocate("40-49 Years, Male", .after="45 - 49 Years, Male") %>%
  select(-c("1­ - 4 years, Female" , "(5 - 9) Years, Female", "40 - 44 Years, Female", "45 - 49 Years, Female",
            "1­ - 4 years, Male", "(5 - 9) Years, Male", "40 - 44 Years, Male", "45 - 49 Years, Male"))

#rename columns
TX_RTT_df <- TX_RTT_df %>% 
  rename("psnu" = "DHIS2 District", "psnuuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support", "period" = "Period", "<01 year, Female" = "< 1 year, Female", "<01 year, Male" = "< 1 year, Male")

#create 1 column for <3, 3-5 and 6+ (collapsing male/female)  
TX_RTT_disag <- TX_RTT_df %>%
  rowwise() %>% 
  mutate("TX_RTT_ <3 Months Interruption"= sum(c_across(contains("<3"))),
         "TX_RTT_3-5 Months Interruption"= sum(c_across(contains("3-5"))),
         "TX_RTT_6+ Months Interruption"= sum(c_across(contains("6+"))),
         .keep = c("unused"))

#delete age/sex cols for time interrupted df
TX_RTT_disag <- select(TX_RTT_disag, -contains(c("Female", "Male")))

#remove <3, 3-5 and 6+ from TX_RTT_df (will add back later)
TX_RTT_df <- TX_RTT_df [, -grep("Experienced", colnames(TX_RTT_df))]

#transpose columns from wide to long
TX_RTT_df <- pivot_longer(TX_RTT_df, contains(c("Female", "Male")), names_to = "age", values_to = "TX_RTT_Now_R")

TX_RTT_df <- TX_RTT_df %>%
  relocate(age, .before = "indicatortype") 

#add Sex column
TX_RTT_df <- TX_RTT_df %>% 
  mutate(sex="NA",
         .after="age")

#fill Sex column with Male or Female
TX_RTT_df$sex <- ifelse(grepl("Female", TX_RTT_df$age), "Female","Male")

#remove all unnecessary non numeric chars from age column
TX_RTT_df$age <- gsub("[^<+0-9.-]", '', TX_RTT_df$age)

#add age_type column
TX_RTT_df <- TX_RTT_df %>%
  mutate(age_type="trendsfine", .before='age') 

#Join both dfs to combine indicators with age/sex component and without age/sex component
TX_RTT_df_final <- bind_rows(TX_RTT_df, TX_RTT_disag)

#Ensure Period values are only FY2022Q1
TX_RTT_df_final <- TX_RTT_df_final %>%
  mutate_at("period", str_replace, "Oct to Dec 2021", "FY22Q1")

TX_RTT_df_final$mech_code <- as.character(TX_RTT_df_final$mech_code)

#before initiating join, make sure age cols don't have white spaces so join initiates properly
TX_CURR_df_final$age <- gsub('\\s+', '', TX_CURR_df_final$age)
TX_NEW_df_final$age <- gsub('\\s+', '', TX_NEW_df_final$age)
TX_ML_df_final$age <- gsub('\\s+', '', TX_ML_df_final$age)
TX_RTT_df_final$age <- gsub('\\s+', '', TX_RTT_df_final$age)


#Left Join DFs on DATIM ID 
# add age, sex and age_type to the merge/join
#consider full_join from dplyr
#example below, but need to check

#add select function for each join that specifies only the most recent quarter 
df <- TX_CURR_df_final %>% 
  full_join(TX_NEW_df_final, 
            by=c("psnu","psnuuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "period", "age_type","age", "sex"))
df <- df %>% 
  full_join(TX_ML_df_final, 
            by=c("psnu","psnuuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "period", "age_type","age", "sex"))

df <- df %>% 
  full_join(TX_RTT_df_final, 
            by=c("psnu","psnuuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "period", "age_type","age", "sex"))#

df <- df %>% 
  rename("orgunituid"="DATIM ID")

# df_test_1 <- df %>%
#   group_by(age) %>%
#   filter(age == "40-44" | age == "45-49") %>%
#   summarise(age = paste(age, collapse = " "))
#   summarise_at(vars('TX_CURR_Now_R':'TX_RTT_6+ Months Interruption'), sum, na.rm = TRUE)
#   mutate(sum(c_across('TX_CURR_Now_R':'TX_RTT_6+ Months Interruption')[age == "40-44" | age == "45-49"], na.rm = T))
#  mutate(age = if_else(age == "01-09", " 01-09", age)) 


# sum_test <- df %>%
#   mutate(age = if_else(age == "1-4" | age == "5-9", "01-09", age)) %>%
#   mutate(age = if_else(age == "40-44" | age == "45-49", "40-49", age)) %>%
#   mutate(age = if_else(age == "50-54" | age == "55-59" | age == "60-64" | age == "65+", "50+", age)) %>%
#   mutate_at(c("TX_CURR_Now_R", "TX_NEW_Now_R","TX_ML_Interruption <3 Months Treatment_Now_R","TX_ML_Interruption 3-5 Months Treatment_R",
#               "TX_ML_Interruption 6+ Months Treatment_R","TX_ML_Died_Now_R","TX_ML_Refused Stopped Treatment_Now_R", "TX_ML_Transferred Out_Now_R",
#               "TX_RTT_Now_R", "TX_RTT_ <3 Months Interruption", "TX_RTT_3-5 Months Interruption", "TX_RTT_6+ Months Interruption"), as.numeric) %>%
#   group_by(orgunituid, mech_code, age, sex) %>%
#   summarise_at(vars('TX_CURR_Now_R':'TX_RTT_6+ Months Interruption'), sum, na.rm = TRUE) %>%
#   


#----Incorporate CoT dashboard----

#delete unusable cols and edit age 
# df_cot <- df_cot %>%
#   filter(df_cot$period != "FY22Q1" & df_cot$period != "FY22Q2") %>%
#   mutate(age = if_else(age == "44570", "01-09", age)) %>%
#   mutate(age = if_else(age == "44848", "10-14", age))

#join TX_NEW and TX_CURR targets onto df
# df_cot_previous <- df_cot %>%
#   filter(period == previous_qtr) %>%
#   select(c("orgunituid", "age", "sex", "TX_NEW_Now_T", "TX_CURR_Now_T")) %>%
#   filter(!is.na(TX_CURR_Now_T) | !is.na(TX_NEW_Now_T))
# # 
#  #add targets to df
# df_test <- df %>%
#    inner_join(df_cot_previous, 
#              by=c("orgunituid", "age", "sex"))

# df_filtered <- df %>%
#   select(c("orgunituid", "age", "sex"))
# 
# testing <- df %>%
#   inner_join(df_cot_fy22q1, 
#             by=c("orgunituid", "age", "sex"))


#create duplicate CoT df in order to replace TX_CURR/TX_NEW Now to Prev
df_cot_dup <- df_cot 
df_cot_dup <- df_cot_dup %>%
  select(-"TX_NEW_2Prev_R", -"TX_CURR_2Prev_R")

 
df_prev_q <- df_cot_dup %>%
  filter(period == previous_qtr) %>%
  rename ("TX_NEW_2Prev_R" = "TX_NEW_Prev_R", "TX_CURR_2Prev_R" = "TX_CURR_Prev_R")%>%
  rename("TX_NEW_Prev_R" = "TX_NEW_Now_R", "TX_CURR_Prev_R" = "TX_CURR_Now_R") %>%
  select(c(snu1, snuprioritization, sitetype, sitename, orgunituid, primepartner, age, sex, TX_NEW_Prev_R, TX_CURR_Prev_R, TX_NEW_Now_T, TX_CURR_Now_T, TX_CURR_2Prev_R, TX_NEW_2Prev_R))


#ensure indicator columns are numeric for both df and CoT
df <- df %>%
  mutate_at(c("TX_CURR_Now_R", "TX_NEW_Now_R","TX_ML_Interruption <3 Months Treatment_Now_R","TX_ML_Interruption 3-5 Months Treatment_R",
              "TX_ML_Interruption 6+ Months Treatment_R","TX_ML_Died_Now_R","TX_ML_Refused Stopped Treatment_Now_R", "TX_ML_Transferred Out_Now_R"), as.numeric)
df_cot <- df_cot %>%
  mutate_at(c("TX_CURR_Now_R", "TX_NEW_Now_R","TX_ML_Interruption <3 Months Treatment_Now_R","TX_ML_Interruption 3-5 Months Treatment_R",
              "TX_ML_Interruption 6+ Months Treatment_R","TX_ML_Died_Now_R","TX_ML_Refused Stopped Treatment_Now_R", "TX_ML_Transferred Out_Now_R",
              "TX_RTT_Now_R", "TX_RTT_ <3 Months Interruption", "TX_RTT_3-5 Months Interruption", "TX_RTT_6+ Months Interruption"), as.numeric)

df <- df %>%
  mutate(countryname="Uganda",
         .before="psnu") %>%
  mutate(operatingunit="Uganda",
         .before="countryname") %>%
  mutate_at("period", str_replace, "FY2022Q3", "FY22Q3")

#filter out the unnecessary duplicate rows 
df <- df %>%
  filter_at(vars(TX_CURR_Now_R, `TX_RTT_ <3 Months Interruption`), any_vars(!is.na(.)))

df_cot <- df_cot %>%
  mutate_at(("mech_code"), as.character)  

#join TX_NEW_Prev_R and TX_CURR_Prev_R onto df
df <- df %>%
  left_join(df_prev_q, 
            by=c("orgunituid", "age", "sex"))

df <- df %>%
  mutate("TX_ML_Interruption 3+ Months Treatment_Now_R"=NA,
                   .before = "TX_ML_Interruption 3-5 Months Treatment_R") %>%
  relocate(snu1:snuprioritization, .after=countryname) %>%
  relocate(sitetype:sitename, .after=psnuuid) %>%
  relocate(primepartner, .after=fundingagency) %>%
  relocate(TX_CURR_Now_T, .after = TX_CURR_Now_R) %>%
  relocate(TX_NEW_Prev_R, .after = TX_CURR_Now_T) %>%
  relocate(TX_NEW_Now_R, .after = TX_NEW_Prev_R) %>%
  relocate(TX_NEW_Now_T, .after = TX_NEW_Now_R) %>%
  relocate(TX_CURR_Prev_R, .after = period) 

#setup sub df with only 1 military aggregation
df_military<- df %>%
  #select(c("orgunituid", "mech_name", "facility", "TX_CURR_Prev_R", "age", "sex", "age_type")) %>%
  filter(mech_name == "URC_DOD_UPDF") 
df_military <-  df_military[-c(51:nrow(df_military)), ]
df_military <-  df_military %>%
  filter(age_type != 'trendsfine') 

#filter out all DOD sites
df <- df %>%
  filter(mech_name != "URC_DOD_UPDF")
  
# join dfs together so that military entries match what is in the CoT dashboard
df <- bind_rows(df, df_military)


df_cot <- df_cot %>%
  mutate_at(("mech_code"), as.character) %>%
  mutate_at(("TX_CURR_Prev_R"), as.numeric) %>%
  mutate_at(("TX_NEW_Prev_R"), as.numeric)

#structure/rename 21Q4 CoT to match structure of 22Q1
#df_cot <- df_cot %>%
#  rename("indicatortype"="indicator_type", "TX_ML_Died_Now_R"="TX_ML_No Contact Outcome - Died_Now_R", "TX_ML_Refused Stopped Treatment_Now_R"="TX_ML_No Contact Outcome - Refused Stopped Treatment_Now_R",
#         "TX_ML_Transferred Out_Now_R"="TX_ML_No Contact Outcome - Transferred Out_Now_R", 
#         "TX_ML_Interruption <3 Months Treatment_Now_R"="TX_ML_No Contact Outcome - Interruption in Treatment <3 Months Treatment_Now_R")

#df_cot <- df_cot %>%
#   relocate(TX_NEW_Now_R, .before=TX_CURR_Now_T) %>%
#   relocate(TX_NEW_Now_T, .after=TX_CURR_Now_T) %>%
#   relocate("TX_ML_Interruption <3 Months Treatment_Now_R", .after=TX_NEW_Now_T)
# 
# #cols did not exist prior to FY22
# df_cot <- df_cot %>%
#   mutate("TX_ML_Interruption 3-5 Months Treatment_R"=NA,
#          .after="TX_ML_Interruption <3 Months Treatment_Now_R") %>%
#   mutate("TX_ML_Interruption 6+ Months Treatment_R"=NA,
#          .after="TX_ML_Interruption 3-5 Months Treatment_R") %>%
#   mutate("TX_RTT_ <3 Months Interruption"=NA,
#          .after="TX_RTT_Now_R") %>%
#   mutate("TX_RTT_3-5 Months Interruption"=NA,
#          .after="TX_RTT_ <3 Months Interruption") %>%
#   mutate("TX_RTT_6+ Months Interruption"=NA,
#          .after="TX_RTT_3-5 Months Interruption")
# df_cot <- df_cot %>%
#   select(-"TX_ML_No Contact Outcome - Interruption in Treatment 3+ Months Treatment_Now_R")
   
#final append
df_final <- bind_rows(df_cot, df)

 # df_final_test <- df_final %>%
 #      filter(period == 'FY22Q1') %>%
 #      filter(age == "15+" | age == "<15") %>%
 #      mutate_at(("TX_CURR_Prev_R"), as.numeric)
 #   
 # sum(df_final_test$TX_CURR_Prev_R, na.rm = TRUE)
 # 
 # df_cot_test <- df_cot %>%
 #   filter(period == 'FY21Q4') %>%
 #   filter(age == "15+") #%>%
 #   #mutate_at(("TX_CURR_Prev_R"), as.numeric)
 # 
 # sum(df_cot_test$TX_CURR_Now_R, na.rm = TRUE)
  
#temporary solution so that dates aren't inputted in csv
df_final <- df_final %>%
  mutate(age = if_else(age == "01-09", " 01-09", age)) %>%
  mutate(age = if_else(age == "10-14", " 10-14", age))

  

#format for FY22Q2
#df_final <- df_final %>%
#   filter(df_final$period != "FY20Q1" & df_final$period != "FY20Q2" & df_final$period != "FY20Q3" & df_final$period != "FY20Q4")

#final adjustments
#df_final <- df_final %>%
#  relocate(TX_NEW_Prev_R, .after = TX_CURR_Now_T) %>%
#  relocate (TX_NEW_Now_R, .after = TX_NEW_Prev_R) %>%
#  mutate("TX_ML_Interruption 3+ Months Treatment_R" = rowSums(select(., "TX_ML_Interruption 3-5 Months Treatment_R":"TX_ML_Interruption 6+ Months Treatment_R")),
#          .before = "TX_ML_Interruption 3-5 Months Treatment_R") %>%
#  mutate_at("period", str_replace, "FY2022Q1", "FY22Q1")
#  select(-facilityprioritization)

#export file
write.csv(df_final, paste0("Dataout/CoT_Waterfall_DHIS2_", current_qtr,".csv"), row.names=F)

#memory.limit(size=20000)

#reload dataset into CoT Dashboard
#wb = loadWorkbook('Data/CoT Dashboard_FY21Q4_Clean_Uganda.xlsx')
#waterfall = read.xlsx(wb, sheet='Waterfall Data')
#writeData(wb, sheet='Waterfall Data', waterfall, startRow=2, colNames=FALSE)
#saveWorkbook(wb, 'CoT Dashboard_FY21Q4_Clean_Uganda.xlsx', overwrite = TRUE)