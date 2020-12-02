# ---- Program Header ----
# Project: 
# Program Name: 
# Author: 
# Created: 
# Purpose: 
# Revision History:
# Date        Author        Revision
# 

# ---- Initialize Libraries ----

library(dplyr)
library(lubridate)
library(haven)
library(openxlsx)
library(tidyverse)
library(labelled)

# ---- Load Functions ----
setwd("C:/Users/JordanMyer/Desktop/New OneDrive/Emanate Life Sciences/DM - Inflammatix - Documents/INF-04/11. Clinical Progamming/11.3 Production Reports/11.3.3 Input Files")
source("../11.3.1  R Production Programs/INF Global Functions.R")
cleaner()
source("../11.3.1  R Production Programs/INF Global Functions.R")


patProfSubjNum <- "04-102-002"


Nth.del <- function(dataframe, n, k) {
  dataframe[,-(seq(n,to=ncol(dataframe),by=k))]
}

queryCheck <- function(form){
  df <- openWithoutAdj%>%filter(SUBJECT.ID==patProfSubjNum&FORM==form)
  return(nrow(df))
}
renameWithLabel <- function(df){
  oldNames <- colnames(df)
  for (i in 1:ncol(df)){
    names(df)[names(df)==oldNames[i]] <- as.character(var_label(df[i]))
  }
  return(df)
}

generateFullListing <- function(importListing){
  df <- importListing%>%
    filter(SubjectStatus=="Enrolled"&SubjectID==patProfSubjNum)%>%
    select(c(2,3,5,7:(ncol(importListing))))
  finalDF <- renameWithLabel(Nth.del(df,5,2))
  finalDF <- finalDF[,-c(ncol(finalDF))]
  return(finalDF)
}

fullListingWithQueries <- function(importFile, sheetName, queryForm,queryPrefix,fileOutput){
  
  generatedListing <- generateFullListing(importFile)
  generatedListing$"SubjVisitTogether" <- paste(generatedListing$`Subject ID`,generatedListing$Visit)
  colnames(generatedListing) <- gsub(" ","",colnames(generatedListing))
  
  filteredOWA <- openWithoutAdj%>%
    filter(FORM==queryForm)%>%
    mutate(VARIABLE=gsub(queryPrefix,"",.$VARIABLE))%>%
    mutate(SubjVisitTogether=paste(SUBJECT.ID,Visit))
  
 
  addWorksheet(wb, sheetName=sheetName)
  
  writeData(wb, sheet=sheetName, x=generatedListing,withFilter = TRUE)
  
  # define style
  red_style <- createStyle(fgFill="#F9B5B5")
  headerStyle <- createStyle(
    fontSize = 14, fontColour = "#FFFFFF", halign = "center",
    fgFill = "#4F81BD", border = "TopBottom", borderColour = "#4F81BD"
  )
  
  #Loop to add style

  rowsWithQueries <- which(generatedListing$SubjVisitTogether%in%filteredOWA$SubjVisitTogether)
  if(length(rowsWithQueries)!=0){
  for(i in 1:length(rowsWithQueries)){
    SubjVisitPairing <- generatedListing$SubjVisitTogether[rowsWithQueries[i]]
    allFields <- filteredOWA%>%filter(SubjVisitTogether==SubjVisitPairing)%>%pull(VARIABLE)
    y <- which(colnames(generatedListing)%in%allFields)
    addStyle(wb, sheet=sheetName, style=red_style, rows=rowsWithQueries[i]+1, col=y, gridExpand=TRUE) # +1 for header line
  }}
  addStyle(wb, sheet=sheetName, style=headerStyle, rows=1,cols=1:ncol(generatedListing), gridExpand=TRUE) # +1 for header line
  setColWidths(wb,sheet=sheetName,cols=1:ncol(generatedListing),widths = "auto")
  #Save
  
  }


allQueries <- read.xlsx(MostRecentFile("Medrio Reports/",".*Medrio_QueryExport_LIVE_Inflammatix_INF_04.*xlsx$","ctime"))

openWithoutAdj <- allQueries%>%
  filter(STATUS=="Open")%>%
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")

wb <- createWorkbook()

#Generate IC ----
importIC <- read_sas("EDC/DS_IC.sas7bdat")

fullListingWithQueries(importFile = importIC, sheetName = "Informed Consent",
                       queryForm = "Informed Consent",
                       queryPrefix = "dsic")

#Generate DEM ----
importDEM <- read_sas("EDC/DM.sas7bdat")

fullListingWithQueries(importFile = importDEM, sheetName = "Demographics",
                       queryForm = "Demographics",
                       queryPrefix = "dm")

#Generate VS ----
importVS <- read_sas("EDC/VS.sas7bdat")

fullListingWithQueries(importFile = importVS, sheetName = "Vital Signs",
                       queryForm = "Vital Signs",
                       queryPrefix = "vs")


#Generate MH ----
importMH <- read_sas("EDC/MH.sas7bdat")

fullListingWithQueries(importFile = importMH, sheetName = "Medical History",
                       queryForm = "Medical History",
                       queryPrefix = "mh",fileOutput = "../11.3.4 Output Files/Patient Profiles/MH_Listing_With_Queries.xlsx")

#Generate IE ----
importMH <- read_sas("EDC/IE.sas7bdat")

fullListingWithQueries(importFile = importMH, sheetName = "Inclusion Exclusion",
                       queryForm = "Inclusion Exclusion",
                       queryPrefix = "ie")

#Generate TOE ----
importTOE <- read_sas("EDC/QS_TOE.sas7bdat")

fullListingWithQueries(importFile = importTOE, sheetName = "TOE Questionnaire",
                       queryForm = "Time-of-Enrollment Provider Questionnaire",
                       queryPrefix = "qstoe")


#Generate CL Sample COllection ----
importSampleCollection <- read_sas("EDC/LB_COLL.sas7bdat")

fullListingWithQueries(importFile = importSampleCollection, sheetName = "Sample Collection - Central Lab",
                       queryForm = "Sample Collection - Central Lab",
                       queryPrefix = "lbcoll")

#Generate Lab Results - Site ----
importLBSITE<- read_sas("EDC/LB_SITE.sas7bdat")

fullListingWithQueries(importFile = importLBSITE, sheetName = "Laboratory Results - Site",
                       queryForm = "Laboratory Results - Site",
                       queryPrefix = "lbsite")


#Generate IM ----
importIM <- read_sas("EDC/IM.sas7bdat")

fullListingWithQueries(importFile = importIM, sheetName = "Imaging",
                       queryForm = "Imaging",
                       queryPrefix = "im")

#Generate FUP ----
importFUP <- read_sas("EDC/FU.sas7bdat")

fullListingWithQueries(importFile = importFUP, sheetName = "Phone Call",
                       queryForm = "Phone Call",
                       queryPrefix = "fu")

#Generate SD ----
importSD <- read_sas("EDC/DS_SD.sas7bdat")

fullListingWithQueries(importFile = importSD, sheetName = "Subject Disposition",
                       queryForm = "Subject Disposition",
                       queryPrefix = "dssd")


#Generate LOS ----
importLOS <- read_sas("EDC/HO.sas7bdat")

fullListingWithQueries(importFile = importLOS, sheetName = "Length of Stay",
                       queryForm = "Length of Stay",
                       queryPrefix = "ho",fileOutput = "../11.3.4 Output Files/Patient Profiles/LOS_Listing_With_Queries.xlsx")




fileOutput = paste("../11.3.4 Output Files/Patient Profiles/",patProfSubjNum,".xlsx",sep="")
saveWorkbook(wb,fileOutput,overwrite = TRUE)

