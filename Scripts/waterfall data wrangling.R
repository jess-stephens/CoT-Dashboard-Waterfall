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

#from FY22Q3
# names(`TX_NEW `) <- `TX_NEW `[4,]
# TX_NEW <- `TX_NEW `[-c(1:4),]

#reshape each df so column headers are in the right place
names(TX_NEW) <- TX_NEW[4,]
TX_NEW <- TX_NEW[-c(1:4),]

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
                     range = cell_limits(c(3, 1), c(NA, NA)),
                     col_names = clean_hdrs_RTT))

TX_RTT <- TX_RTT[-1,]

#remove irrelevant column strings for columns that contain MER
names(TX_RTT)[-1] <- ifelse(grepl("MER:", names(TX_RTT)[-1]), sub(".*_", "", names(TX_RTT)[-1]), names(TX_RTT)[-1])

#delete NAs added to the other cols
names(TX_RTT) = gsub("NA_", "", x = names(TX_RTT))


#---------------------------------TX_NEW---------------------------------
#remove rows with all NAS
TX_NEW <- TX_NEW[,colSums(is.na(TX_NEW))<nrow(TX_NEW)]
#drop columns with Unknown string
TX_NEW <- TX_NEW [, -grep("Unknown", colnames(TX_NEW))]

#choose columns 
TX_NEW_df <- TX_NEW %>% 
  select(c("DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
                     "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male")))) 

#change indicator cols to numeric + dropping Male, Female columns
TX_NEW_df <- TX_NEW_df %>%
  mutate_at(c(10:length(TX_NEW_df)), as.numeric) %>%
  select(-c("Male", "Female"))  

#separate age cols with no space between age band and year
colnames(TX_NEW_df) <- gsub("([0-9])([yY])", "\\1 \\2", colnames(TX_NEW_df))

#Rename cols
 TX_NEW_df <- TX_NEW_df %>% 
   rename("psnu" = "DHIS2 District", "psnuuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
          "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
          "indicatortype" = "Type of Support", "period" = "Period", "Male\r\n<15 yrs" = "<15  yrs Male", "Male\r\n15+ yrs" = "15+ yrs  Male",
          "Female\r\n<15 yrs" = "<15  yrs  Female", "Female\r\n15+ yrs" = "15+ yrs  Female")

list_TX_NEW <- c("yr", "yrs")

## Construct the regular expression
pat_NEW <- paste0("\\b(", paste0(list_TX_NEW, collapse="|"), ")\\b")  

#Pivot wide to long
TX_NEW_df_long <- TX_NEW_df %>% 
  pivot_longer(contains(c("Female", "Male")),
               names_to=c("sex", "age"),
               names_sep= "_|\r",
               values_to="TX_NEW_Now_R") %>%
  mutate(age=gsub(pat_NEW,"", age)) %>%
  mutate(across(where(is.character), str_trim)) %>%
  mutate(age=case_when(age=="<1" ~ '<01',
                       age== "1-4" ~ ' 01-09',
                       age== "5-9" ~ ' 01-09',
                       age=="10-14" ~ ' 10-14',
                       age=="15-19" ~ '15-19',
                       age=="20-24" ~ '20-24',
                       age=="25-29" ~ '25-29',
                       age=="30-34" ~ '30-34',
                       age=="35-39" ~ '35-39',
                       age=="40-44" ~ "40-49",
                       age=="45-49" ~ "40-49",
                       age=="<15" ~ '<15',
                       age=="15+" ~ '15+',
                       TRUE ~ "50+")) %>%
  drop_na(TX_NEW_Now_R)

#sum 01-09 and 40-49 age band rows
TX_NEW_df_grouped <- TX_NEW_df_long %>%
  group_by(psnu, psnuuid, `DATIM ID`, fundingagency, mech_name, mech_code, facility, indicatortype,
           period, sex, age) %>%
  summarize(TX_NEW_Now_R = sum(TX_NEW_Now_R)) %>%
  ungroup()

#create age_type  
coarse_values <- "^<15|^15+"
TX_NEW_df_final <- TX_NEW_df_grouped %>% 
  relocate(age, .before = sex) %>%
  mutate(age_type=ifelse(grepl(coarse_values, TX_NEW_df_grouped$age), "trendscoarse","trendsfine"), .before="age") %>%
  mutate(age_type = ifelse(age == '15-19', 'trendsfine', age_type))

#---------------------------------TX_CURR---------------------------------
#remove rows with all NAS
TX_CURR <- TX_CURR[,colSums(is.na(TX_CURR))<nrow(TX_CURR)]

#delete out of place Male and Female cols
TX_CURR <- select(TX_CURR, -c("Male", "Female", "15+ Years, Female", "15+ Years, Male", contains(c("Unknown", "ARVs", 
                                                                                   "Sex Workers", "<20", "20+", "weight", "DTG"))))
#choose columns
TX_CURR_df <- TX_CURR %>%
  select(c("DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#change indicator cols to numeric
TX_CURR_df <- TX_CURR_df %>%
  mutate_at(c(10:length(TX_CURR_df)), as.numeric)

#separate age cols with no space between age band and year
colnames(TX_CURR_df) <- gsub("([0-9])([yY])", "\\1 \\2", colnames(TX_CURR_df))

#rename columns
TX_CURR_df <- TX_CURR_df %>% 
  rename("psnu" = "DHIS2 District", "psnuuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support", "period" = "Period", "Male\r\n<15 yrs" = "<15  yrs Male", "Male\r\n15+ yrs" = "15+ yrs  Male",
         "Female\r\n<15 yrs" = "<15  yrs  Female", "Female\r\n15+ yrs" = "15+ yrs  Female")

list_TX_CURR <- c("yr", "yrs")

## Construct the regular expression
pat_CURR <- paste0("\\b(", paste0(list_TX_CURR, collapse="|"), ")\\b")  


TX_CURR_df_long <- TX_CURR_df %>% 
  pivot_longer(contains(c("Female", "Male")),
               names_to=c("sex", "age"),
               names_sep= "_|\r",
               values_to="TX_CURR_Now_R") %>%
  mutate(age=gsub(pat_CURR,"", age)) %>%
  mutate(across(where(is.character), str_trim)) %>%
  mutate(age=case_when(age=="<1" ~ '<01',
                       age== "1-4" ~ "01-09",
                       age== "5-9" ~ "01-09",
                       age=="10-14" ~ '10-14',
                       age=="15-19" ~ '15-19',
                       age=="20-24" ~ '20-24',
                       age=="25-29" ~ '25-29',
                       age=="30-34" ~ '30-34',
                       age=="35-39" ~ '35-39',
                       age=="40-44" ~ "40-49",
                       age=="45-49" ~ "40-49",
                       age=="<15" ~ '<15',
                       age=="15+" ~ '15+',
                       TRUE ~ "50+")) %>%
  drop_na(TX_CURR_Now_R)

#sum 01-09, 40-49 and 50+ age bands
TX_CURR_df_grouped <- TX_CURR_df_long %>%
  group_by(psnu, psnuuid, `DATIM ID`, fundingagency, mech_name, mech_code, facility, indicatortype,
           period, sex, age) %>%
  summarize(TX_CURR_Now_R = sum(TX_CURR_Now_R)) %>%
  ungroup()

TX_CURR_df_final <- TX_CURR_df_grouped %>% 
     relocate(age, .before = sex) %>%
     mutate(age_type=ifelse(grepl(coarse_values, TX_CURR_df_grouped$age), "trendscoarse","trendsfine"), .before="age") %>%
     mutate(age_type = ifelse(age == '15-19', 'trendsfine' ,age_type))

#---------------------------------TX_ML---------------------------------
#remove unwanted columns
TX_ML <- select(TX_ML, -c(contains(c("Unknown", "Numerator:", "resulting in", "(FSW)", "(MSM)"))))

#choose columns
TX_ML_df <- TX_ML %>% 
  select(c("DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#change indicator cols to numeric
TX_ML_df <- TX_ML_df %>%
  mutate_at(c(10:length(TX_ML_df)), as.numeric)

list_TX_ML <- c("year", "years", "Year", "Years" )

## Construct the regular expression
pat_ML <- paste0("\\b(", paste0(list_TX_ML, collapse="|"), ")\\b")    

TX_ML_df_long <- TX_ML_df %>% 
               pivot_longer(contains(c("Female", "Male")),
               names_to=c("disag", "age", "sex"),
               names_sep= "_|,",
               values_to="value") %>% 
               mutate(disag=gsub("by Age/Sex", "",disag ), 
               age=gsub(pat_ML,"", age)) %>% 
               mutate(across(where(is.character), str_trim)) %>% 
               mutate(age=case_when(age=="< 1" ~ '<01',
                                    age=="1-4" ~ '01-09',
                                    age=="5 - 9" ~ '01-09',
                                    age=="10-14" ~ '10-14',
                                    age=="15 - 19" ~ "15-19",
                                    age=="20-24" ~ '20-24',
                                    age=="25-29" ~ '25-29',
                                    age=="30-34" ~ '30-34',
                                    age=="35-39" ~ '35-39',
                                    age=="40 - 44" ~ "40-49",
                                    age=="45 - 49" ~ "40-49",
                                    TRUE ~ "50+")) %>%
              drop_na(value)

df_ML_wider<-TX_ML_df_long %>% 
  pivot_wider(names_from = disag, values_from=value, values_fn = sum)

#relocate age and sex
df_ML_wider <- df_ML_wider %>%
  relocate(age, .after = "DHIS2 HF Name") %>%
  relocate(sex, .after = "age") %>%
  mutate(age_type="trendsfine", .before='age')

#rename columns
TX_ML_df_final <- df_ML_wider %>% 
  rename("psnu" = "DHIS2 District", "psnuuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support",  "period"="Period",
         "TX_ML_Interruption <3 Months Treatment_Now_R" =  "IIT After being on Treatment for <3 month",
         "TX_ML_Interruption 3-5 Months Treatment_R" = "IIT After being on Treatment for 3-5 months",
         "TX_ML_Interruption 6+ Months Treatment_R" = "IIT After being on Treatment for 6+ months", 
         "TX_ML_Died_Now_R" = "Died",
         "TX_ML_Refused Stopped Treatment_Now_R" = "Refused (Stopped) Treatment", 
         "TX_ML_Transferred Out_Now_R" = "transferred out" )

#---------------------------------TX_RTT---------------------------------
#remove unwanted columns
TX_RTT <- select(TX_RTT, -c(contains(c("KP", "IIT", "MSM", "Sex Workers", "Sub Total"))))

#choose columns
TX_RTT_df <-TX_RTT %>% 
  select(c("DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#change age cols to numeric
TX_RTT_df <- TX_RTT_df %>%
  mutate_at(c(10:length(TX_RTT_df)), as.numeric)

#create 1 column for <3, 3-5 and 6+ (collapsing male/female)  
TX_RTT_disag <- TX_RTT_df %>%
  mutate_at(c("Experienced treatment interruption of <3 months before returning to treatment_Female", "Experienced treatment interruption of 3-5 months before returning to treatment_Female",
              "Experienced treatment interruption of 6+ months before returning to treatment_Female", "Experienced treatment interruption of <3 months before returning to treatment_Male",
              "Experienced treatment interruption of 3-5 months before returning to treatment_Male", "Experienced treatment interruption of 6+ months before returning to treatment_Male"), ~replace_na(.,0)) %>%
  rowwise() %>% 
  mutate("TX_RTT_ <3 Months Interruption"= sum(c_across(contains("<3"))),   
         "TX_RTT_3-5 Months Interruption"= sum(c_across(contains("3-5"))),
         "TX_RTT_6+ Months Interruption"= sum(c_across(contains("6+"))),
         .keep = c("unused")) %>%
  select(-contains(c("Female", "Male")))
  

#remove <3, 3-5 and 6+ from TX_RTT_df (will add back later)
TX_RTT_df <- TX_RTT_df [, -grep("Experienced", colnames(TX_RTT_df))]

list_RTT <- c("year", "years", "Year", "Years" )

## Construct the regular expression
pat_RTT <- paste0("\\b(", paste0(list_RTT, collapse="|"), ")\\b")  

TX_RTT_df_long <- TX_RTT_df %>% 
  pivot_longer(contains(c("Female", "Male")),
               names_to=c("age", "sex"),
               names_sep= "_|,",
               values_to="TX_RTT_Now_R") %>%
  mutate(age=gsub(pat_RTT,"", age)) %>%
  mutate(across(where(is.character), str_trim)) %>%
  mutate(age=case_when(age=="< 1" ~ '<01',
                       age=="1-4" ~ '01-09',
                       age=="5 - 9" ~ '01-09',
                       age=="10-14" ~ '10-14',
                       age=="15 - 19" ~ "15-19",
                       age=="20-24" ~ '20-24',
                       age=="25-29" ~ '25-29',
                       age=="30-34" ~ '30-34',
                       age=="35-39" ~ '35-39',
                       age=="40 - 44" ~ "40-49",
                       age=="45 - 49" ~ "40-49",
                       TRUE ~ "50+")) %>%
  drop_na(TX_RTT_Now_R)

#Join both dfs to combine indicators with age/sex component and without age/sex component
TX_RTT_df_long <- bind_rows(TX_RTT_df_long, TX_RTT_disag)

TX_RTT_df_long <- TX_RTT_df_long %>% 
  rename("psnu" = "DHIS2 District", "psnuuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support",  "period"="Period")

#sum 01-09 and 40-49 age band rows
TX_RTT_df_long <- TX_RTT_df_long %>%
  group_by(psnu, psnuuid, `DATIM ID`, fundingagency, mech_name, mech_code, facility, indicatortype,
           period, age, sex, `TX_RTT_ <3 Months Interruption`, `TX_RTT_3-5 Months Interruption`, `TX_RTT_6+ Months Interruption`) %>%
  summarize(TX_RTT_Now_R = sum(TX_RTT_Now_R)) %>%
  ungroup()

#add age_type column
TX_RTT_df_final <- TX_RTT_df_long %>%  
  mutate(age_type="trendsfine", .before='age') %>%
  relocate(TX_RTT_Now_R, .before = `TX_RTT_ <3 Months Interruption`) 

TX_RTT_df_final$mech_code <- as.character(TX_RTT_df_final$mech_code)


#Join DFs on DATIM ID 
df <- TX_CURR_df_final %>% 
  full_join(TX_NEW_df_final, 
            by=c("psnu","psnuuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "period", "age_type","age", "sex"))
df <- df %>% 
  full_join(TX_ML_df_final, 
            by=c("psnu","psnuuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "period", "age_type","age", "sex"))

df <- df %>% 
  full_join(TX_RTT_df_final, 
            by=c("psnu","psnuuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "period", "age_type","age", "sex"))

df <- df %>% 
  rename("orgunituid"="DATIM ID")

#----Incorporate CoT dashboard----

#create duplicate CoT df in order to replace TX_CURR/TX_NEW Now to Prev
df_cot_dup <- df_cot 
df_cot_dup <- df_cot_dup %>%
  select(-"TX_NEW_2Prev_R", -"TX_CURR_2Prev_R")
 
df_prev_q <- df_cot_dup %>%
  filter(period == previous_qtr) %>%
  rename("TX_NEW_2Prev_R" = "TX_NEW_Prev_R", "TX_CURR_2Prev_R" = "TX_CURR_Prev_R")%>%
  rename("TX_NEW_Prev_R" = "TX_NEW_Now_R", "TX_CURR_Prev_R" = "TX_CURR_Now_R") %>%
  select(c(snu1, snuprioritization, sitetype, sitename, orgunituid, primepartner, mech_code, age, sex, TX_NEW_Prev_R, TX_CURR_Prev_R, TX_NEW_Now_T, TX_CURR_Now_T, TX_CURR_2Prev_R, TX_NEW_2Prev_R)) 

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
  mutate_at("period", str_replace, "FY2023Q1", current_qtr) #automate at some point


#join TX_NEW_Prev_R and TX_CURR_Prev_R onto df
df <- df %>%
 left_join(df_prev_q, 
            by=c("mech_code", "orgunituid", "age", "sex")) %>%
  distinct()


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
  relocate(TX_CURR_Prev_R, .after = sex) %>%
  mutate_at(c("TX_NEW_Prev_R", "TX_CURR_Prev_R", "TX_NEW_Now_T", "TX_CURR_Now_T", "TX_CURR_2Prev_R", "TX_NEW_2Prev_R"), as.numeric)

df_cot <- df_cot %>%
  mutate_at(c("TX_NEW_Prev_R", "TX_CURR_Prev_R", "TX_NEW_Now_T", "TX_CURR_Now_T", "TX_CURR_2Prev_R", "TX_NEW_2Prev_R"), as.numeric)

#final append
df_final <- bind_rows(df_cot, df)

#temporary solution so that dates aren't inputted in csv
df_final <- df_final %>%
  mutate(age = if_else(age == "01-09", " 01-09", age)) %>%
  mutate(age = if_else(age == "10-14", " 10-14", age))

df_final <- df_final %>%
  filter(period != "FY21Q1") %>%
  filter(period != "FY21Q2") %>%
  filter(period != "FY21Q3") %>%
  filter(period != "FY21Q4") %>%
  filter(period != "FY22Q1") %>%
  filter(period != "FY22Q2") %>%
  filter(period != "FY22Q3")
  
#export file
write.csv(df_final, paste0("Dataout/CoT_Waterfall_DHIS2_", current_qtr,".csv"), row.names=F)

#reload dataset into CoT Dashboard
#wb = loadWorkbook('Data/CoT Dashboard_FY21Q4_Clean_Uganda.xlsx')
#waterfall = read.xlsx(wb, sheet='Waterfall Data')
#writeData(wb, sheet='Waterfall Data', waterfall, startRow=2, colNames=FALSE)
#saveWorkbook(wb, 'CoT Dashboard_FY21Q4_Clean_Uganda.xlsx', overwrite = TRUE)