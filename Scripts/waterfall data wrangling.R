library(tidyverse)
library(openxlsx)
library(readxl)
library(reshape2)
library(data.table)

library(glamr)
library(tidyverse)


# A R Project is required for this folder path. If no R project, the full folder path of the user is required. 
fldr <- "Data"

file <- "8_TX Data - FY22Q1.xlsx" 
path <- file.path(fldr, '8_TX Data - FY22Q1.xlsx')

#this needs to be adjusted depending on how the collection of datasets will work together

path_in <- file.path(fldr, file)


# Create list of excel sheets and turn each into separate dataframes
sheet = excel_sheets(path_in)

mylist = lapply(setNames(sheet, sheet), function(x) read_excel(path_in, sheet=x))
names(mylist) <- sheet
list2env(mylist ,.GlobalEnv)



#reshape each df so column headers are in the right place
names(`TX_NEW `) <- `TX_NEW `[4,]
TX_NEW <- `TX_NEW `[-c(1:4),]

# names(TX_NEW) <- TX_NEW[4,]
# TX_NEW <- TX_NEW[-c(1:4),]

names(TX_CURR) <- TX_CURR[4,]
TX_CURR <- TX_CURR[-c(1:4),]

# names(TX_ML) <- TX_ML[3,]
# TX_ML <- TX_ML[-c(1:3),]

(hdr1 <- read_excel(path,
                    sheet = "TX_ML",
                    skip = 3,
                    n_max = 1,
                    .name_repair = "minimal") %>% 
    names())
(hdr2 <- read_excel(path,
                    sheet = "TX_ML",
                    skip = 4,
                    n_max = 1,
                    .name_repair = "minimal") %>% 
    names())
(clean_hdrs <- tibble(hdr1, hdr2)%>% 
    mutate(hdr1 = na_if(hdr1, "")) %>% 
    fill(hdr1) %>% 
    unite(hdr, c(hdr1, hdr2)) %>% 
    pull())
(TX_ML <- read_excel(path,
                  sheet = "TX_ML",
                  skip = 6,
                  col_names = clean_hdrs))


names(TX_RTT) <- TX_RTT[2,]
TX_RTT <- TX_RTT[-c(1:2),]


#---------------------------------TX_NEW---------------------------------

#choose columns 
TX_NEW_df <-TX_NEW %>% 
  select(c("Statistical Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
                     "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))   

#rename columns
TX_NEW_df <- TX_NEW_df %>% 
  rename("region" = "Statistical Region", "psnu" = "DHIS2 District", "psnuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
          "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support")

#transpose columns from wide to long
TX_NEW_df <- pivot_longer(TX_NEW_df, contains(c("Female", "Male")), names_to = "age", values_to = "TX_NEW_Now_R")


#delete all rows that have an age disagg equal to 0
#TX_NEW_df <- subset(TX_NEW_df,TX_NEW_df$values!= "0")
#duplicate rows based on specified age counts
#TX_NEW_df <- setDT(TX_NEW_df)[,.(count=1:values), TX_NEW_df]
#remove irrelevant columns
#TX_NEW_df <- select(TX_NEW_df, -"values")

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
#https://stackoverflow.com/questions/46153832/exact-match-with-grepl-r --- remove after reviewing
TX_NEW_df<-  TX_NEW_df %>% 
  mutate(age_type=ifelse(grepl(coarse_values, TX_NEW_df$age), "trendscoarse","trendsfine"))


#---------------------------------TX_CURR---------------------------------

#choose columns 
TX_CURR_df = TX_CURR[c("UAIS 2011_ Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency","COP  Mechanism name",
                       "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", "Unknown age Female", "Female\r\n<1yr",
                       "Female\r\n1-4yrs", "Female\r\n5-9yrs", "Female\r\n10-14 yrs", "Female\r\n15-19 yrs", "Female\r\n20-24 yrs",                                   
                       "Female\r\n25-29 yrs", "Female\r\n30-34 yrs", "Female\r\n35-39 yrs", "Female\r\n40-44 yrs", "Female\r\n45-49 yrs",
                       "Female\r\n 50-54 yrs", "Female\r\n 55-59 yrs", "Female\r\n 60-64 yrs", "Female\r\n 65+ yrs","Male\r\nUnknown age",
                       "Male\r\n<1yr", "Male\r\n1-4yrs", "Male\r\n5-9yrs", "Male\r\n10-14 yrs", "Male\r\n20-24 yrs", "Male\r\n25-29 yrs",                                     
                       "Male\r\n30-34 yrs", "Male\r\n35-39 yrs","Male\r\n40-44 yrs", "Male\r\n45-49 yrs", "Male\r\n 50-54 yrs",
                       "Male\r\n 55-59 yrs", "Male\r\n 60-64 yrs", "Male\r\n 65+ yrs", "<15  yrs Male", "<15  yrs  Female","15+ yrs  Male",
                       "15+ yrs  Female")]  

#rename columns
TX_CURR_df <- rename(TX_CURR_df, "region" = "UAIS 2011_ Region")
TX_CURR_df <- rename(TX_CURR_df, "psnu" = "DHIS2 District")
TX_CURR_df <- rename(TX_CURR_df, "psnuid" = "DHIS2 ID")
TX_CURR_df <- rename(TX_CURR_df, "fundingagency" = "COP US Agency")
TX_CURR_df <- rename(TX_CURR_df, "mech_name" = "COP  Mechanism name")
TX_CURR_df <- rename(TX_CURR_df, "mech_code" = "COP  Mechanism ID")
TX_CURR_df <- rename(TX_CURR_df, "facility" = "DHIS2 HF Name")
TX_CURR_df <- rename(TX_CURR_df, "indicatortype" = "Type of Support")


#transpose columns from wide to long
TX_CURR_df <- pivot_longer(TX_CURR_df, "Unknown age Female":"15+ yrs  Female", names_to = "age", values_to = "TX_CURR_Now_R")

#delete all rows that have an age disagg equal to 0
#TX_CURR_df <- subset(TX_CURR_df,TX_CURR_df$values!= "0")
#duplicate rows based on specified age counts
#TX_CURR_df<- setDT(TX_CURR_df)[,.(count=1:values), TX_CURR_df]

#add Sex column
TX_CURR_df <- TX_CURR_df %>% 
  mutate(sex="NA",
         .before="indicatortype")
TX_CURR_df <- TX_CURR_df %>%
  relocate(age, .after = sex)

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
coarse_values <- "<15|15+"
TX_CURR_df$age_type <- ifelse(grepl(coarse_values, TX_CURR_df$age), "trendscoarse","trendsfine")

#handle exception case for 15-19
TX_CURR_df <- TX_CURR_df %>%
  mutate(age_type = ifelse(age ==  "\r\n15-19 " , "trendsfine", age_type))


#---------------------------------TX_ML---------------------------------

#Creating new columns for TX_ML

#<3 months
TX_ML <- TX_ML %>% 
         rowwise() %>% 
         mutate(TX_ML_Interruption_Less_Than_3_Months_Treatment_Now_R = sum(c_across(c(72:96)), na.rm = T))
#3+ months
TX_ML <- TX_ML %>% 
  rowwise() %>% 
  mutate(TX_ML_Interruption_More_Than_3_Months_Treatment_Now_R = sum(c_across(c(97:146)), na.rm = T))
#3-5 months
TX_ML <- TX_ML %>% 
  rowwise() %>% 
  mutate(TX_ML_Interruption_3_To_5_Months_Treatment_Now_R = sum(c_across(c(97:121)), na.rm = T))
#6+ months
TX_ML <- TX_ML %>% 
  rowwise() %>% 
  mutate(TX_ML_Interruption_More_Than_6_Months_Treatment_Now_R = sum(c_across(c(122:146)), na.rm = T))
#Died Now R
TX_ML <- TX_ML %>% 
  rowwise() %>% 
  mutate(TX_ML_Died_Now_R = sum(c_across(c(22:46)), na.rm = T))
#Refused/Stopped Treatment R
TX_ML <- TX_ML %>% 
  rowwise() %>% 
  mutate(TX_ML_Refused_Stopped_Treatment_Now_R = sum(c_across(c(147:171)), na.rm = T))
#Transferred Out
TX_ML <- TX_ML %>% 
  rowwise() %>% 
  mutate(TX_ML_Transferred_Out_Now_R = sum(c_across(c(47:71)), na.rm = T))

#choose columns
TX_ML_df <- TX_ML[c("UAIS 2011_ Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency",
                  "COP  Mechanism name", "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period",
                  "TX_ML_Interruption_Less_Than_3_Months_Treatment_Now_R","TX_ML_Interruption_More_Than_3_Months_Treatment_Now_R", "TX_ML_Interruption_3_To_5_Months_Treatment_Now_R",
                  "TX_ML_Interruption_More_Than_6_Months_Treatment_Now_R", "TX_ML_Died_Now_R", "TX_ML_Refused_Stopped_Treatment_Now_R",
                  "TX_ML_Transferred_Out_Now_R")] 


#rename columns
TX_ML_df <- rename(TX_ML_df, "region" = "UAIS 2011_ Region")
TX_ML_df <- rename(TX_ML_df, "psnu" = "DHIS2 District")
TX_ML_df <- rename(TX_ML_df, "psnuid" = "DHIS2 ID")
TX_ML_df <- rename(TX_ML_df, "fundingagency" = "COP US Agency")
TX_ML_df <- rename(TX_ML_df, "mech_name" = "COP  Mechanism name")
TX_ML_df <- rename(TX_ML_df, "mech_code" = "COP  Mechanism ID")
TX_ML_df <- rename(TX_ML_df, "facility" = "DHIS2 HF Name")
TX_ML_df <- rename(TX_ML_df, "indicatortype" = "Type of Support")

  

#---------------------------------TX_RTT---------------------------------

#Creating new columns for TX_RTT 
#<3 months
TX_RTT <- TX_RTT %>% 
  rowwise() %>% 
  mutate(TX_RTT_Interruption_Less_Than_3_Months = sum(c_across(c(54:55)), na.rm = T))
#3-5 months
TX_RTT <- TX_RTT %>% 
  rowwise() %>% 
  mutate(TX_RTT_Interruption_3_to_5_Months = sum(c_across(c(56:57)), na.rm = T))
#6+ months
TX_RTT <- TX_RTT %>% 
  rowwise() %>% 
  mutate(TX_RTT_Interruption_More_Than_6_Months = sum(c_across(c(58:59)), na.rm = T))

#choose columns
TX_RTT_df <- TX_RTT[c("UAIS 2011_ Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
                    "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", "TOTAL", "TX_RTT_Interruption_Less_Than_3_Months",
                    "TX_RTT_Interruption_3_to_5_Months", "TX_RTT_Interruption_More_Than_6_Months")]

#rename columns
TX_RTT_df <- rename(TX_RTT_df, "region" = "UAIS 2011_ Region")
TX_RTT_df <- rename(TX_RTT_df, "psnu" = "DHIS2 District")
TX_RTT_df <- rename(TX_RTT_df, "psnuid" = "DHIS2 ID")
TX_RTT_df <- rename(TX_RTT_df, "fundingagency" = "COP US Agency")
TX_RTT_df <- rename(TX_RTT_df, "mech_name" = "COP  Mechanism name")
TX_RTT_df <- rename(TX_RTT_df, "mech_code" = "COP  Mechanism ID")
TX_RTT_df <- rename(TX_RTT_df, "facility" = "DHIS2 HF Name")
TX_RTT_df <- rename(TX_RTT_df, "indicatortype" = "Type of Support")
TX_RTT_df <- rename(TX_RTT_df, "TX_RTT_Now_R" = "TOTAL")


#Left Join DFs on DATIM ID
df = merge(x=TX_CURR_df, y=TX_NEW_df, id="DATIM ID", all.x=TRUE)
df = merge(x=df, y=TX_ML_df, id="DATIM ID", all.x=TRUE)
df = merge(x=df, y=TX_RTT_df, id="DATIM ID", all.x=TRUE)

#last bit of data manipulation 

#create countryname column 
df <- df %>%
  mutate(countryname="Uganda",
         .before="region")
#create operatingunit column
df <- df %>%
  mutate(operatingunit="Uganda",
         .before="countryname")

df <- df %>%
  relocate(age, .after = facility)
df <- df %>%
  relocate(sex, .after = age)
df <- rename(df, "period" = "Period")
