###### BEFORE RUNNING THIS CODE FOR THE FIRST TIME ON A NEW DEVICE, PLEASE MAKE SURE ALL PACKAGES USED HERE ARE INSTALLED #####
###### ADDITIONALLY, PLEASE MAKE SURE YOU HAVE JAVA INSTALLED ON YOUR COMPUTER: https://www.oracle.com/java/technologies/downloads/ #####

# Load Libraries
library(tidyverse)
library(readxl)
library(lubridate)
library(officer)
library(svDialogs)
library(openxlsx)
library(rvest)
library(chromote)
library(RSelenium)

#-----------------------CHECK THE CODE BELOW BEFORE RUNNING THIS SCRIPT-----------------------
# current semester
currentSemester = "Fall 2024"
journalURLs = data.frame(journal = c("CDPS","CPS","SPPS"),
       url = c("https://browzine.com/libraries/1892/journals/5012/issues/current", 
    "https://browzine.com/libraries/1892/journals/32462/issues/current",
    "https://browzine.com/libraries/1892/journals/7502/issues/current"))
#----------------------------------------------------------------------------------------------


# get current submission date from user
submissionDate = as.Date(dlg_input(message = "Enter current submission deadline (mm/dd)", default = "")$res, "%m/%d")


#----------------------- GET JOURNAL INFO ---------------------------

# Start RSelenium driver
driver <- rsDriver(browser = "firefox",
       port = 4555L,
       verbose = FALSE,
       chromever = NULL)

# Initialize an empty dataframe to store journal information
journalInfo = data.frame(journalIndex = character(), 
       journal = character(), 
       volume = double(), 
       issue = double(), 
       article = character())

# Loop through each journal URL
for (i in 1:nrow(journalURLs)) {
  # extract the client for readability of the code to follow
  remote_driver <- driver[["client"]]
  
  # Set URL
  url <- as.character(journalURLs[i,"url"])
  
  # Navigate to the webpage
  remote_driver$navigate(url)
  
  # make R pause execution of further code to let the page load
  Sys.sleep(10) 
  
  # Get code from the webpage
  page_source <- remote_driver$getPageSource()[[1]]
  thisJournal <- read_html(page_source)
  
  # Extract the title of the journal
  title <- thisJournal %>%
    html_element("h1") %>%
    html_text2()
  
  # Store the article names
  articles <- thisJournal %>% 
    html_elements("section.article-list-item-content-block") %>% 
    html_text2() %>%
    str_extract("[^.\n]+")
  
  # Extract volume and issue text
  text <- thisJournal %>%
    html_elements("h2") %>%
    html_text2()
  
  # Store volume and issue number
  vol <- str_extract(text[1], "[^Vol. ]+")
  iss <- str_extract(text[1], "(?<=Issue ).*")
  
  # Loop through each article and add to journalInfo dataframe
  for (j in 1:length(articles)) {
    journalInfo <- journalInfo %>%
      add_row(journalIndex = journalURLs[i, "journal"],
              journal = title,
              volume = as.double(vol),
              issue = as.double(iss),
              article = articles[j])
  }
  
  # make R pause execution of further code to let the page load
  Sys.sleep(3) 
}

# close connection
driver$server$stop()

#--------------------------------------------------------------------

#load in the QuestionPro excel file
QPro = read_excel('QuestionPro.xlsx', sheet = 'Raw Data')

# Create the dataframe for MASTER excel file
QPro_master = QPro %>%
  # select only completed surveys
  filter(`Response Status` == 'Completed' & is.na(`First Name`) == F) %>% 
  # select only columns of interest
  select('Timestamp (mm/dd/yyyy)','First Name', 'Last Name', 'UTA Email Address',
     'Student Sona ID', 'Journal Name', 'Volume Number', 'Issue Number',
     'Article Name:',
     `The summary must be at least 500-1000 words. Do NOT put in paragraph indentations or extra spacing.`) %>%
  # rename columns
  rename('Article Name' = 'Article Name:',
     'Summary' = `The summary must be at least 500-1000 words. Do NOT put in paragraph indentations or extra spacing.`) %>% 
  # create new columns
  mutate(`Correct Journal` = "",
     `Correct Volume` = "",
     `Correct Issue` = "",
     `Correct Article` = "",
     `>500 words` = lengths(strsplit(Summary, "(?!['-])(\\W+)", perl = T)),
     `Summarized the Article` = "",
     "Readable" = "",
     `No Plagiarism` = "",
     `Give Credit?` = "",
     "Notes" = "") %>%
  arrange(`Timestamp (mm/dd/yyyy)`)

# add submission deadline column to the data frame
`Submisison Deadline` = format(submissionDate, "%b %d %Y")
QPro_master = cbind(`Submisison Deadline`, QPro_master) 

# replace NAs with 999
QPro_master[is.na(QPro_master)] = 999

# check if journal name, volume, and issue, are correct
# This script checks if the journal name, volume, issue, and article name in the QPro_master dataframe
# match the corresponding information in the journalInfo dataframe. It updates the QPro_master dataframe
# with "x" if the information is correct.

# Loop through each row in the QPro_master dataframe
for (i in 1:nrow(QPro_master)) {
  # Loop through each row in the journalInfo dataframe
  for (j in 1:nrow(journalInfo)) {
    # Check if the journal name matches
    if (QPro_master$`Journal Name`[i] == journalInfo$journal[j]) {
      # Mark the journal name as correct
      QPro_master$`Correct Journal`[i] = "x"
      # Check if the volume number matches
      if (QPro_master$`Volume Number`[i] == journalInfo$volume[j]) {
        # Mark the volume number as correct
        QPro_master$`Correct Volume`[i] = "x"
        # Check if the issue number matches (converted to double for comparison)
        if (as.double(QPro_master$`Issue Number`[i]) == journalInfo$issue[j]) {
          # Mark the issue number as correct
          QPro_master$`Correct Issue`[i] = "x"
          # Check if the article name matches
          if (QPro_master$`Article Name`[i] == journalInfo$article[j]) {
            # Mark the article name as correct
            QPro_master$`Correct Article`[i] = "x"
          }
        }
      }
    }
  }
}

# Check if file already exists
# If it exist, get the data from it so there are no duplicates when writing
# back into it
# If it doesn't exist, create a new file
masterExcelName = paste("MASTER Paper Summary Credit Sheet - ",currentSemester, ".xlsx", sep = "")
if (file.exists(masterExcelName) == T) {
  wb_master = loadWorkbook(masterExcelName)
  df_master = readWorkbook(wb_master)
  # find the last submission date
  # arrange the df by submission date
  df_master = df_master %>% arrange(`Timestamp.(mm/dd/yyyy)`)
  lastSubmissionDate = mdy(df_master[nrow(df_master),1])
  # keep only rows of people who submitted for this submission date
  QPro_master = QPro_master %>%  
  filter(lastSubmissionDate < floor_date(QPro_master$`Timestamp (mm/dd/yyyy)`, "day")
       & mdy(`Submisison Deadline`) >= floor_date(QPro_master$`Timestamp (mm/dd/yyyy)`, "day"))
  # bind master and QPro_master
  #bound_master = bind_rows(master, QPro_master)
  # create styling for the excel file
  # style for font of the header
  #headerStyle = createStyle(fontSize = 11, halign = "center", valign = "center",
  #                         textDecoration = "bold", wrapText = T)
  # style for manual input columns
  #manualColumnsStyle = createStyle(border = "TopBottomLeftRight")
  # create some vars to use
  worksheetName = "Grading Sheet"
  #colNum = ncol(bound_master)
  # create workbook and worksheets
  masterWB = createWorkbook()
  addWorksheet(masterWB, worksheetName)
  #add our data into the empty workbook
  writeData(wb_master, sheet = worksheetName, QPro_master, startRow = nrow(df_master)+2, colNames = F)
  # freeze top row
  #freezePane(masterWB, worksheetName, firstRow = T)
  # add filter
  #addFilter(masterWB, worksheetName, rows = 1, cols = 1:colNum)
  # add style
  #addStyle(masterWB, sheet = worksheetName, headerStyle, 
  #        rows = 1, cols = 1:colNum, gridExpand = TRUE)
  #addStyle(masterWB, sheet = worksheetName, manualColumnsStyle, 
  #        rows = 2:1000, cols = 12:21, gridExpand = TRUE)
  # add the current submissions to the master file
  saveWorkbook(wb_master, file = masterExcelName, overwrite = T, returnValue = F)
} else {
  # keep only rows of people who submitted for this submission date
  QPro_master = QPro_master %>%  
  filter(mdy(`Submisison Deadline`) >= floor_date(QPro_master$`Timestamp (mm/dd/yyyy)`, "day"))
  # create styling for the excel file
  # style for font of the header
  headerStyle = createStyle(fontSize = 11, halign = "center", valign = "center",
              textDecoration = "bold", wrapText = T)
  # style for manual input columns
  manualColumnsStyle = createStyle(border = "TopBottomLeftRight")
  # create some vars to use
  worksheetName = "Grading Sheet"
  colNum = ncol(QPro_master)
  # create workbook and worksheets
  master = createWorkbook()
  addWorksheet(master, worksheetName)
  #add our data into the empty workbook
  writeData(master, sheet = worksheetName, QPro_master)
  # freeze top row
  freezePane(master, worksheetName, firstRow = T)
  # add filter
  addFilter(master, worksheetName, rows = 1, cols = 1:colNum)
  # add style
  addStyle(master, sheet = worksheetName, headerStyle, 
       rows = 1, cols = 1:colNum, gridExpand = TRUE)
  addStyle(master, sheet = worksheetName, manualColumnsStyle, 
       rows = 2:1000, cols = 12:21, gridExpand = TRUE)
  # save the workbook
  saveWorkbook(master, file = masterExcelName, overwrite = T, returnValue = F)
}  

# Assign QPro_master to master
master = QPro_master 


#-----------------------Create separate word docs for student summaries-----------------------

# Create a folder for current submissions
# Create a name for the folder
submissionFolder = paste(as.character(format(submissionDate, "%b %d")), "PAPER SUMMARIES")
# Create the folder
dir.create(file.path(getwd(), submissionFolder))

# Set current and submission directories
currDir = getwd()
subDir = file.path(getwd(), submissionFolder)

# create a list of folders to look for students' previous submissions
folderList = list.dirs(full.names = F, recursive = F)

# create word docs for each student's submission
# Create a list of unique student names
studentList = c(unique(paste(master$`First Name`, master$`Last Name`)))

# Loop through each student in the student list
for (student in studentList) {
  # Loop through each folder in the folder list
  for (folder in folderList) {
    # Check if there are any files in the current directory that match the student's name
    if (length(list.files(path = ".", pattern = student)) > 0) {
      # If there are, add the student's name to the student list, repeated for the number of matching files
      studentList = c(studentList, rep(student, (length(list.files(path = ".", pattern = student)))))
    }
  }
}

# Loop through each student's submission and create a word document
for (i in 1:nrow(master)) {  # Loop through each row in the master dataframe
  studentName = paste(master$`First Name`[i], master$`Last Name`[i])  # Create the student's full name
  print(studentName)  # Print the student's name for debugging purposes
  if (sum(studentList == studentName) > 1) {  # Check if the student has multiple submissions
    read_docx() %>%  # Create a new Word document
      body_add_par(master$Summary[i]) %>%  # Add the student's summary to the document
      print(target = paste(subDir, "/", studentName, " - ", sum(studentList == studentName), ".docx", sep = ""))  # Save the document with a unique name
    studentList = c(studentList, studentName)  # Add the student's name to the student list
  } else {  # If the student has only one submission
    read_docx() %>%  # Create a new Word document
      body_add_par(master$Summary[i]) %>%  # Add the student's summary to the document
      print(target = paste(subDir, "/", studentName, ".docx", sep = ""))  # Save the document with the student's name
    studentList = c(studentList, studentName)  # Add the student's name to the student list
  }
}
