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

names(TX_CURR) <- TX_CURR[4,]
TX_CURR <- TX_CURR[-c(1:4),]

# names(TX_ML) <- TX_ML[3,]
# TX_ML <- TX_ML[-c(1:3),]

(hdr1 <- read_excel(path_in,
                    sheet = "TX_ML",
                    skip = 3,
                    n_max = 1,
                    .name_repair = "minimal") %>% 
    names())
(hdr2 <- read_excel(path_in,
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

#read in data frame skipping the header rows
(TX_ML <- read_excel(path_in,
                  sheet = "TX_ML",
                  range = cell_limits(c(5, 1), c(NA, NA)),
                  col_names = clean_hdrs))

#delete extra title text from indicators
names(TX_ML)[-1] <- sub("(.*:[^:]+).*:", "", names(TX_ML)[-1])

#delete NAs added to the other cols
names(TX_ML) = gsub("NA_", "", x = names(TX_ML))

#remove first row
TX_ML <- TX_ML[-1,]

#process needs to be repeated for TX_RTT

names(TX_RTT) <- TX_RTT[2,]
TX_RTT <- TX_RTT[-c(1:2),]


#---------------------------------TX_NEW---------------------------------
#drop columns with Unknown string
TX_NEW <- TX_NEW [, -grep("Unknown", colnames(TX_NEW))]

#choose columns 
TX_NEW_df <-TX_NEW %>% 
  select(c("Statistical Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
                     "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male")))) 

#delete and Male and Female cols
TX_NEW_df = select(TX_NEW_df, -c("Male", "Female"))

#rename columns
TX_NEW_df <- TX_NEW_df %>% 
  rename("region" = "Statistical Region", "psnu" = "DHIS2 District", "psnuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
          "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support")

#replace \r and \n with white space for age cols
names(TX_NEW_df)[-1] <- sub("\\r\\n", " ", names(TX_NEW_df)[-1])

#transpose columns from wide to long
TX_NEW_df <- pivot_longer(TX_NEW_df, contains(c("Female", "Male")), names_to = "age", values_to = "TX_NEW_Now_R")

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
TX_NEW_df<-  TX_NEW_df %>% 
  mutate(age_type=ifelse(grepl(coarse_values, TX_NEW_df$age), "trendscoarse","trendsfine"))


#---------------------------------TX_CURR---------------------------------

#remove rows with unknown age and ARVs to avoid unneccessary cols being pulled
TX_CURR <- TX_CURR [, -grep("Unknown", colnames(TX_CURR))]
TX_CURR <- TX_CURR [, -grep("ARVs", colnames(TX_CURR))]

#choose columns
TX_CURR_df <-TX_CURR %>% 
  select(c("UAIS 2011_ Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#delete and Male and Female cols
TX_CURR_df = select(TX_CURR_df, -c("Male", "Female"))

#rename columns
TX_CURR_df <- TX_CURR_df %>% 
  rename("region" = "UAIS 2011_ Region", "psnu" = "DHIS2 District", "psnuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support")

#transpose columns from wide to long
TX_CURR_df <- pivot_longer(TX_CURR_df, "Unknown age Female":"15+ yrs  Female", names_to = "age", values_to = "TX_CURR_Now_R")

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
TX_ML <- TX_ML [, -grep("Unknown", colnames(TX_ML))]

#choose columns
TX_ML_df <-TX_ML %>% 
  select(c("UAIS 2011_ Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#rename columns
TX_ML_df <- TX_ML_df %>% 
  rename("region" = "UAIS 2011_ Region", "psnu" = "DHIS2 District", "psnuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support")

#change line spacing 
names(TX_ML_df)[-1] <- sub("\\r\\n", " ", names(TX_ML_df)[-1])

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
