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

Nth.del <- function(dataframe, n, k) {
  dataframe[,-(seq(n,to=ncol(dataframe),by=k))]
}

queryCheck <- function(subject,form){
  df <- openWithoutAdj%>%filter(SUBJECT.ID==subject&FORM==form)
  return(nrow(df))
}
renameWithLabel <- function(df){
  oldNames <- colnames(df)
  for (i in 1:ncol(df)){
    names(df)[names(df)==oldNames[i]] <- as.character(var_label(df[i]))
  }
  return(df)
}
addQueryCount <- function(importForm, formName){
  
  FormSubjects <- importForm$SubjectID
  queryNumberList <- c()
  subjectNumberList <- c()
  for (i in 1:nrow(importForm)){
    queryNumberList[length(queryNumberList)+1] <- queryCheck(FormSubjects[i],formName)
    #subjectNumberList[length(subjectNumberList)+1] <- Forms
  }
  queryListDF <- data.frame(SubjectID=FormSubjects,NumQueries=queryNumberList)
  IEWithQueryCount <- left_join(importIE,queryListDF)
  return(IEWithQueryCount)
}
generateFullListing <- function(importListing,subjectNumber){
  df <- importListing%>%
    filter(SubjectStatus=="Enrolled"&SubjectID==subjectNumber)%>%
    select(c(2,3,5,7:(ncol(importListing))))
  
  finalDF <- renameWithLabel(Nth.del(df,5,2))
  finalDF <- finalDF[,-c(ncol(finalDF))]
  return(finalDF)
}

fullListingWithQueries <- function(importFile, sheetName, queryForm,queryPrefix,fileOutput,subjectNumber){
  
  generatedListing <- generateFullListing(importFile)
  generatedListing$"SubjVisitTogether" <- paste(generatedListing$`Subject ID`,generatedListing$Visit)
  colnames(generatedListing) <- gsub(" ","",colnames(generatedListing))
  
  filteredOWA <- openWithoutAdj%>%
    filter(FORM==queryForm)%>%
    mutate(VARIABLE=gsub(queryPrefix,"",.$VARIABLE))%>%
    mutate(SubjVisitTogether=paste(SUBJECT.ID,Visit))
  
  wb <- createWorkbook()
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
  for(i in 1:length(rowsWithQueries)){
    SubjVisitPairing <- generatedListing$SubjVisitTogether[rowsWithQueries[i]]
    allFields <- filteredOWA%>%filter(SubjVisitTogether==SubjVisitPairing)%>%pull(VARIABLE)
    y <- which(colnames(generatedListing)%in%allFields)
    addStyle(wb, sheet=sheetName, style=red_style, rows=rowsWithQueries[i]+1, col=y, gridExpand=TRUE) # +1 for header line
  }
  addStyle(wb, sheet=sheetName, style=headerStyle, rows=1,cols=1:ncol(generatedListing), gridExpand=TRUE) # +1 for header line
  setColWidths(wb,sheet=sheetName,cols=1:ncol(generatedListing),widths = "auto")
  #Save
  saveWorkbook(wb, fileOutput, overwrite=TRUE)
  }


allQueries <- read.xlsx(MostRecentFile("Medrio Reports/",".*Medrio_QueryExport_LIVE_Inflammatix_INF_04.*xlsx$","ctime"))

openWithoutAdj <- allQueries%>%
  filter(STATUS=="Open")%>%
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")

#Generate TOE ----
importTOE <- read_sas("EDC/QS_TOE.sas7bdat")

fullListingWithQueries(importFile = importTOE, sheetName = "TOE Questionnaire",
                       queryForm = "Time-of-Enrollment Provider Questionnaire",
                       queryPrefix = "qstoe",fileOutput = "../11.3.4 Output Files/CDRP Listings/TOE_Listing_With_Queries.xlsx",
                       subjectNumber = "04-102-002")

#Generate MBSITE ----
importMBSITE <- read_sas("EDC/MB_SITE.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Microbiology Results - Site")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importMBSITE, sheetName = "Microbiology Results - Site",
                       queryForm = "Microbiology Results - Site",
                       queryPrefix = "mbsite",fileOutput = "../11.3.4 Output Files/CDRP Listings/MBSITE_Listing_With_Queries.xlsx")


#Generate IM ----
importIM <- read_sas("EDC/IM.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Imaging")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importIM, sheetName = "Imaging",
                       queryForm = "Imaging",
                       queryPrefix = "im",fileOutput = "../11.3.4 Output Files/CDRP Listings/IM_Listing_With_Queries.xlsx")

#Generate FUP ----
importFUP <- read_sas("EDC/FU.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Phone Call")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importFUP, sheetName = "Phone Call",
                       queryForm = "Phone Call",
                       queryPrefix = "fu",fileOutput = "../11.3.4 Output Files/CDRP Listings/FUP_Listing_With_Queries.xlsx")

#Generate IE ----
importIE <- read_sas("EDC/IE.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Inclusion Exclusion")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importIE, sheetName = "Inclusion Exclusion",
                       queryForm = "Inclusion Exclusion",
                       queryPrefix = "ie",fileOutput = "../11.3.4 Output Files/CDRP Listings/IE_Listing_With_Queries.xlsx")


#Generate IC ----
importIC <- read_sas("EDC/DS_IC.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Informed Consent")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importIC, sheetName = "Informed Consent",
                       queryForm = "Informed Consent",
                       queryPrefix = "dsic",fileOutput = "../11.3.4 Output Files/CDRP Listings/IC_Listing_With_Queries.xlsx")


#Generate DEM ----
importDEM <- read_sas("EDC/DM.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Demographics")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importDEM, sheetName = "Demographics",
                       queryForm = "Demographics",
                       queryPrefix = "dm",fileOutput = "../11.3.4 Output Files/CDRP Listings/DEM_Listing_With_Queries.xlsx")

#Generate VS ----
importVS <- read_sas("EDC/VS.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Vital Signs")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importVS, sheetName = "Vital Signs",
                       queryForm = "Vital Signs",
                       queryPrefix = "vs",fileOutput = "../11.3.4 Output Files/CDRP Listings/VS_Listing_With_Queries.xlsx")

#Generate Sample COllection ----
importSampleCollection <- read_sas("EDC/LB_COLL.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Sample Collection - Central Lab")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importSampleCollection, sheetName = "Sample Collection - Central Lab",
                       queryForm = "Sample Collection - Central Lab",
                       queryPrefix = "lbcoll",fileOutput = "../11.3.4 Output Files/CDRP Listings/Sample_Collection_Listing_With_Queries.xlsx")

#Generate MH ----
importMH <- read_sas("EDC/MH.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Medical History")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importMH, sheetName = "Medical History",
                       queryForm = "Medical History",
                       queryPrefix = "mh",fileOutput = "../11.3.4 Output Files/CDRP Listings/MH_Listing_With_Queries.xlsx")


#Generate SD ----
importSD <- read_sas("EDC/DS_SD.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Subject Disposition")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importSD, sheetName = "Subject Disposition",
                       queryForm = "Subject Disposition",
                       queryPrefix = "dssd",fileOutput = "../11.3.4 Output Files/CDRP Listings/SD_Listing_With_Queries.xlsx")

#Generate LOS ----
importLOS <- read_sas("EDC/HO.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Length of Stay")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importLOS, sheetName = "Length of Stay",
                       queryForm = "Length of Stay",
                       queryPrefix = "ho",fileOutput = "../11.3.4 Output Files/CDRP Listings/LOS_Listing_With_Queries.xlsx")

#Generate Lab Results - Site ----
importLBSITE<- read_sas("EDC/LB_SITE.sas7bdat")
unique(allQueries$FORM)
unique(allQueries%>%filter(FORM=="Laboratory Results - Site")%>%pull(VARIABLE))

fullListingWithQueries(importFile = importLBSITE, sheetName = "Laboratory Results - Site",
                       queryForm = "Laboratory Results - Site",
                       queryPrefix = "lbsite",fileOutput = "../11.3.4 Output Files/CDRP Listings/Lab_Res_Listing_With_Queries.xlsx")



#Generate EW ----
importEW <- read_sas("EDC/DS_EW.sas7bdat")
df <- importEW%>%
  select(c(2,3,5,7:(ncol(importEW)-3)))
generatedListing <- renameWithLabel(df)

wb <- createWorkbook()
addWorksheet(wb, sheetName="Early Withdrawl")
writeData(wb, sheet="Early Withdrawl", x=generatedListing,withFilter = TRUE)

# define style
red_style <- createStyle(fgFill="#F9B5B5")
headerStyle <- createStyle(
  fontSize = 14, fontColour = "#FFFFFF", halign = "center",
  fgFill = "#4F81BD", border = "TopBottom", borderColour = "#4F81BD"
)

addStyle(wb, sheet="Early Withdrawl", style=headerStyle, rows=1,cols=1:ncol(generatedListing), gridExpand=TRUE) # +1 for header line
setColWidths(wb,sheet=1,cols=1:ncol(generatedListing),widths = "auto")
#Save
saveWorkbook(wb, "../11.3.4 Output Files/CDRP Listings/EW_Listing_With_Queries.xlsx", overwrite=TRUE)

