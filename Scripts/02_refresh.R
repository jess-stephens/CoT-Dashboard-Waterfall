# Project: Uganda CoT Interagency Dashboard
# Script: 02_refresh
# Developers: Alex Brun (DOD), Jessica Stephens (USAID)
# Use: To integrate Uganda MoH/METS quarterly data into ICPI/CoT CoOP Continuity of Treatment
# Dashboard, in advance of the DATIM release to encourage more timely data usage. 

###########################################
###########################################
###### Instructions ######################
###########################################
###########################################

###### Download Data ###########################################
# Access most recent quarter CoT Dashboard here: https://pepfar.sharepoint.com/sites/ICPI/Products/Forms/AllItems.aspx?csf=1&web=1&e=kbLh7w&cid=5bed6b68%2Dea64%2D4173%2D8ecf%2Dd536c0de9385&FolderCTID=0x01200073B328CBB9171E4DA6242FF8A712DFE9&id=%2Fsites%2FICPI%2FProducts%2FICPI%20Approved%20Tools%20%28Most%20Current%20Versions%29%2FTreatment%20Dashboards%2FICPI%20Continuity%20of%20Treatment%20Dashboard&viewid=22678384%2D0fc3%2D42c5%2Dbd7a%2Ddbdbbaa8c083
# Download most recent dashboard for Uganda. 
# Open Dashboard and resave as .xlsx (.xlbx not easily manipulated)
########################################################################


###### Update folder paths ###########################################
# A R Project is required for this folder path. 
# If using GitHub, an R project will automatically be associated.
# If no R project, the full folder path of the user is required. To do so, add before "Data" when identifying "fldr" 
########################################################################
# First, add raw data and CoT Dashboard from previous quarter to the Data folder
# save data to the "Data" folder (should be in R project) and tell R to search there
fldr <- "Data" 
# update MoH/METS in process file name below, especially if name format changes
file <- "8_TX Data - FY22Q1.xlsx" 
# update CoT Dashboard file name below (change FYXXQX). 
main_file <- "Continuity in Treatment Dashboard_FY21Q4_Clean_Uganda.xlsx" 

########  Update Quarters ################################################################
################################################################
# Each quarter, current_qtr and last_qtr but be updated by user/refresh team
current_qtr="FY21Q1"
last_qtr="FY20Q4"

########   Munge Data ################################################################
################################################################
# Run function which cleans the MoH/METS inprocess data and integrates it with prior quarters data
source("Scripts/waterfall data wrangling.R")

#######################################################################  
# Integrate raw data into CoT Excel Dashboard
################################################################
