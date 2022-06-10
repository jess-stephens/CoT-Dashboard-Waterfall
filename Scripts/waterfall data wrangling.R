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

#remove rows with unknown age and ARVs to avoid unneccessary cols being pulled
TX_CURR <- TX_CURR [, -grep("Unknown", colnames(TX_CURR))]
TX_CURR <- TX_CURR [, -grep("ARVs", colnames(TX_CURR))]
TX_CURR <- TX_CURR [, -grep("Sex Workers", colnames(TX_CURR))]

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
  select(c("UAIS 2011_ Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#change line spacing 
names(TX_ML_df)[-1] <- sub("\\r\\n", " ", names(TX_ML_df)[-1])
names(TX_ML_df)


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
  rename("region" = "UAIS 2011_ Region", "psnu" = "DHIS2 District", "psnuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support", "TX_ML_Interruption <3 Months Treatment_Now_R" =  " IIT After being on Treatment for <3 month by Age/Sex",
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
  select(c("UAIS 2011_ Region", "DHIS2 District", "DHIS2 ID", "DATIM ID", "COP US Agency", "COP  Mechanism name",
           "COP  Mechanism ID", "DHIS2 HF Name", "Type of Support", "Period", contains(c("Female","Male"))))

#rename columns
TX_RTT_df <- TX_RTT_df %>% 
  rename("region" = "UAIS 2011_ Region", "psnu" = "DHIS2 District", "psnuid" = "DHIS2 ID", "fundingagency" = "COP US Agency",
         "mech_name" = "COP  Mechanism name", "mech_code" = "COP  Mechanism ID", "facility" = "DHIS2 HF Name", 
         "indicatortype" = "Type of Support")

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
  mutate_at("Period", str_replace, "Oct to Dec 2021", "FY2022Q1")

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
df <- TX_CURR_df_final %>% 
  full_join(TX_NEW_df_final, 
            by=c("region", "psnu","psnuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "Period", "age_type","age", "sex"))

df <- df %>% 
  full_join(TX_ML_df_final, 
            by=c("region", "psnu","psnuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "Period", "age_type","age", "sex"))

df <- df %>% 
  full_join(TX_RTT_df_final, 
            by=c("region", "psnu","psnuid","DATIM ID", "fundingagency","mech_name", "mech_code","facility","indicatortype", "Period", "age_type","age", "sex"))

#need to repeat for other df

# df = merge(x=TX_CURR_df, y=TX_NEW_df, id="DATIM ID", all.x=TRUE)
# df = merge(x=df, y=TX_ML_df, id="DATIM ID", all.x=TRUE)
# df = merge(x=df, y=TX_RTT_df, id="DATIM ID", all.x=TRUE)

#last bit of data manipulation 

#create countryname column 
#df <- df %>%
 # mutate(countryname="Uganda",
  #       .before="region")
#create operatingunit column
#df <- df %>%
#  mutate(operatingunit="Uganda",
#       .before="countryname")
#import msd
#left_join msd to df to include snu1, etc missing columns
