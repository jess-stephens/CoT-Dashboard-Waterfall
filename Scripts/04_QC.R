# Project: Uganda CoT Interagency Dashboard
# Script: 04_QC
# Developers: Alex Brun (DOD), Jessica Stephens (USAID)
# Use: To integrate Uganda MoH/METS quarterly data into ICPI/CoT CoOP Continuity of Treatment
# Dashboard, in advance of the DATIM release to encourage more timely data usage. 


str(df)

df2<-df %>% 
mutate_at(c("TX_CURR_Now_R", "TX_NEW_Now_R","TX_ML_Interruption <3 Months Treatment_Now_R","TX_ML_Interruption 3-5 Months Treatment_R",
           "TX_ML_Interruption 6+ Months Treatment_R","TX_ML_Died_Now_R","TX_ML_Refused Stopped Treatment_Now_R", "TX_ML_Transferred Out_Now_R"), as.numeric) %>% 
  clean_names()

str(df2)



install.packages("psych")
library(psych)

describeBy(df2, group=c(df2$age, df2$sex))


df2<-df %>% 
  group_by(c(age,sex)) %>% 
  mutate(sum("TX_RTT_6+ Months Interruption"))

df3<-df2 %>% 
group_by(age, sex, .groups) %>% 
  summarise(tx_rtt_6_months_interruption_sum = sum(tx_rtt_6_months_interruption))

str(df2)

df3<-df2 %>% 
  group_by(age, sex) %>% 
  # summarise(across(where(is.numeric), list(sum=sum)))
summarise(across(where(is.numeric), list(sum = sum)))
df3



df3<-df2 %>% 
  group_by(age, sex) %>% 
  summarise(TX_RTT_6+_Months_Interruption = sum("TX_RTT_6+ Months Interruption"))



write.csv(df2, "C:/Users/jstephens/Documents/ICPI/CoT/CoT Dashboard Waterfall/CoT-Dashboard-Waterfall/Dataout/QC.csv")