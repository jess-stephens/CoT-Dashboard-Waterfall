######## ONLY RUN THIS SCRIPT THE FIRST TIME USING THIS CODE ON YOUR COMPUTERS #####

# Project: Uganda CoT Interagency Dashboard
# Script: 00_setup
# Developers: Alex Brun (DOD), Jessica Stephens (USAID)
# Use: To integrate Uganda MoH/METS quarterly data into ICPI/CoT CoOP Continuity of Treatment
# Dashboard, in advance of the DATIM release to encourage more timely data usage. 
 

# Instructions:

# Install packages
install.packages(c("tidyverse","openxlsx", "readxl","reshape2", "data.table", "usethis"))
install.packages("remotes")
remotes::install_github("USAID-OHA-SI/glamr", build_vignettes = TRUE)

# Create standard folder structure
library(glamr)
si_setup()

# Add raw data to the Data folder
# This step needs to be done physically, outside of R