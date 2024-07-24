# Survey-Project - An R code that extracts data from survey question. 
#Requirements to be met 
#This script should run on an automated schedule at 12am EST everday (bonus)
#It should loop through each line row within this document "C:/Users/user/Downloads/Standard Indicator Form.xlsx" and remembers so that it doesn't loop over it again once the three files (remember each row is a different "project") so files should be different) "C:/Users/user/Downloads/Training - Training Table Template.xlsx", "C:/Users/user/Downloads/Disaggregations.xlsx", "C:/Users/user/OneDrive/Downloads/Indicators.xlsx") have been generated for that project.

#The first thing it should do is create a new folder that is the value of column I, that puts these 3 files there for that specific row: "C:/Users/user/Downloads/Training - Training Table Template.xlsx", "C:/Users/user/Downloads/Disaggregations.xlsx", "C:/Users/user/OneDrive/Downloads/Indicators.xlsx") once the script actions below are complete:

#If column J has the value "Yes", then import this sheet for that project. "C:/Users/user/Downloads/Training - Training Table Template.xlsx"

#If column J has a value of "No" or it is empty, then skip pass the step.

#If column L from "Standard Indicator Form.xlsx" has a value of "In-person training" or "Virtual training" then delete column O from "Training - Training Table Template.xlsx"

#If column M from "Standard Indicator Form.xlsx" has a value of "In-person training" or "Virtual training" then delete column P from "Training - Training Table Template.xlsx"

#If column N from "Standard Indicator Form.xlsx" has a value of "One country" then delete column E from "Training - Training Table Template.xlsx"

#If column N from "Standard Indicator Form.xlsx" has a value of "Multiple countries" then delete column D from "Training - Training Table Template.xlsx"

#If column Q from "Standard Indicator Form.xlsx" has a value of "Only topic only" then delete columns R and S from "Training - Training Table Template.xlsx"

#If column Q from "Standard Indicator Form.xlsx" has a value of "Multiple topics" then delete column Q from "Training - Training Table Template.xlsx"

#If column Y from "Standard Indicator Form.xlsx" is empty or has value of "Certificate of Completion" then delete column Y from ".Training - Training Table Template.xlsx"

#If column W from "Standard Indicator Form.xlsx" has value of "Anonymous" then delete columns H, I, J, L, M, and N from "Training - Training Table Template.xlsx"

#Now create a function called disaggregations() that takes the variable input of column and disaggregation_suffix

#The disaggregation suffixes are "Course Topic", "Course Name", "Participant Sector", "Participant Agency/Office/Organization", and "Certification Type".

#The disaggregation suffixes should be parsed together with the value that is in column I of the "Standard Indicator Form" with a -, so for example the full disaggregation name should look like this "23GR0024-AMEWestBank-JLA-03012023 - Course Topic" for the first row in Standard Indicator Form" workbook.

#For "Course Topic" if a value exists in column P of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column P of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside it otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Course Name" if a value exists in column S of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column P of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside it otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Participant Sector" if a value exists in column T of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column P of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside it otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Participant Agency/Office/Organization" if a value exists in column U of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column P of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside it otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Certificate Type" if a value exists in column Y of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column P of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside unless it says, "Certificate of Completion" in column Y of the "Standard Indicator Form" otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Equipment Type" if a value exists in column AL of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column AL of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside unless it says, "Certificate of Completion" in column Y of the "Standard Indicator Form" otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Beneficiary Type" if a value exists in column AM of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column AL of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside unless it says, "Certificate of Completion" in column Y of the "Standard Indicator Form" otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Type of Goods Seized" if a value exists in column AS of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column AL of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside unless it says, "Certificate of Completion" in column Y of the "Standard Indicator Form" otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Interdiction Location Type" if a value exists in column AT of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column AL of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside unless it says, "Certificate of Completion" in column Y of the "Standard Indicator Form" otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#For "Search Type" if a value exists in column AV of the "Standard Indicator Form" then print the full disaggregation name in column A of "Disaggregations.xlsx" workbook into the next available row and print the value of column AL of the "Standard Indicator Form" into column B of "Disaggregations.xlsx" right beside unless it says, "Certificate of Completion" in column Y of the "Standard Indicator Form" otherwise don't put the disaggregation in "Disaggregations.xlsx" at all.

#Now If column P has any of these keywords: "narcotics", "wildlife", "environmental products", "persons", "arms", "money laundering", "counterfeiting", "smuggling", "cybercrime", "cyber", "trafficking", "laundering", "counter", "money", "TOC", "International" then copy row 2 of IndicatorsMaster.xlsx": C:/Users/user/OneDrive/Documents/IndicatorsMaster.xlsx into "Indicators.xlsx in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#Now If column P has any of these keywords: "cybercrime" then copy row 19 of IndicatorsMaster.xlsx": C:/Users/user/OneDrive/Documents/IndicatorsMaster.xlsx into "Indicators.xlsx in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#If Column T has the keyword "Justice (prosecutors and court officials)" then copy row 3 of IndicatorsMaster.xlsx": C:/Users/user/OneDrive/Documents/IndicatorsMaster.xlsx into "Indicators.xlsx in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#If Column T has the keyword "Government Officials" and column P has the keyword "Anti-corruption" then copy row 3 of IndicatorsMaster.xlsx": C:/Users/user/OneDrive/Documents/IndicatorsMaster.xlsx into "Indicators.xlsx in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#If Column T has the keyword "Civil Society" or "Business / Private sector" and column P has the keyword "Anti-corruption" then copy row 3 of IndicatorsMaster.xlsx": C:/Users/user/OneDrive/Documents/IndicatorsMaster.xlsx into "Indicators.xlsx in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#IF column V of "Standard Indicator Form" includes any of these strings:

#"Average difference in pre- and post-test scores" - copy row 10 of "IndicatorsMasters.xlsx" into next available row into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#"Average score given by participants on how relevant they feel the training" - copy row 7 of "IndicatorsMasters.xlsx" into next available row into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#"Average score given by participants on their satisfaction with the course materials" - copy row 8 of "IndicatorsMasters.xlsx" into next available row into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#"Average score given by participants on their satisfaction with the instructor" - copy row 9 of "IndicatorsMasters.xlsx" into next available row into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#"Percentage who pass the post-test" - copy row 11 of "IndicatorsMasters.xlsx" into next available row into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#"Percentage who pass the pass/fail assessment" - copy row 12 of "IndicatorsMasters.xlsx" into next available row into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#If column Z of "Standard Indicator Form" has a value, then copy row 13 of "IndicatorsMaster.xlsx" into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#If column AA of "Standard Indicator Form" has a value, then copy row 14 of "IndicatorsMaster.xlsx" into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"

#If column AH of "Standard Indicator Form" has a value of "Yes" then import C:/Users/user/Downloads/Equipment - Equipment Table Template.xlsx and import row 15 of "IndicatorsMaster.xlsx" into "Indicators.xlsx", if it is a value of "No" then don't

#If column AQ of "Standard Indicator Form" has a value of "Yes" then import C:/Users/userl/Downloads/interdiction - Interdiction Table Template.xlsx, if it is a value of "No" then don't

#Import rows 16, 17, 18, 19, and 20 of "IndicatorsMaster.xlsx" into "Indicators.xlsx" in the next available rows, paste row I of "Standard Indicator Form" into Row B of "Indicators.xlsx" and row S of "Indicators.xlsx" with the prefix "P/"


This is the code

library(readxl)
library(openxlsx)
library(dplyr)

# Define file paths
standard_indicator_path <- "C:/Users/Admin/Downloads/Standard Indicator Form.xlsx"
training_template_path <- "C:/Users/Admin/Downloads/Training - Training Table Template.xlsx"
disaggregations_path <- "C:/Users/user/Downloads/Disaggregations.xlsx"
indicators_path <- "C:/Users/Admin/Downloads/Indicators.xlsx"
indicators_master_path <- "C:/Users/Admin/Downloads/Indicatorsmaster.xlsx"
equipment_template_path <- "C:/Users/user/Downloads/Equipment - Equipment Table Template.xlsx"
interdiction_template_path <- "C:/Users/user/Downloads/interdiction - Interdiction Table Template.xlsx"

# Read data
indicator_data <- read_excel(standard_indicator_path)
indicators_master <- read_excel(indicators_master_path, sheet = 1)

# Print column names to ensure they match expected values
print(colnames(indicator_data))

# Function to remove columns by name, handling non-existent columns gracefully
remove_columns <- function(df, cols_to_remove) {
  cols_to_remove <- cols_to_remove[cols_to_remove %in% colnames(df)]
  if (length(cols_to_remove) > 0) {
    df <- df %>% select(-all_of(cols_to_remove))
  }
  return(df)
}

# Create new folder named after the value in the project code column
# Rename columns
colnames(indicator_data) <- c(
  "ID", "Start_time", "Completion_time", "Email", "Name", 
  "Read_Note", "Contact_Name", "Contact_Email", "Project Code", 
  "Formal Training", "Training_Info", "Training Mode", "Training Type", 
  "Country Scope", "Participant_Countries", "Course Topic", 
  "Topic Scope", "Course_List", "Course Name", "Participant Sector", 
  "Participant Agency/Office/Organization", "Evaluation_Indicators", "Participant Anonymity", 
  "Certifications", "Certificate Type", "Total_Participants", 
  "Total_Events", "In_Person_Events", "Event_Info", "International_Participants", 
  "GPP_Funding", "Event_Types", "Total_Events_Target", "Equipment_Donation", 
  "Equipment_Donation_Info", "Donation_Countries","Benefit_Countries","Equipment Type", 
  "Beneficiary Type", "Total_Donation_Value", "Construction_Activities", 
  "Construction_Locations", "Interdiction_Activities", "Interdiction_Info", 
  "Type of Goods Seized", "Interdiction Location Type", "Specific_Locations", 
  "Search Type", "Track_Searches", "Search_Target", 
  "Track_Flags", "Flag_Target", "Wildlife_Trafficking", 
  "Arrests_Prosecutions", "Arrests_Target", "Anti_Corruption_Indicators", 
  "Anti_Corruption_Target", "Improved_Case_Management", "Case_Management_Target", 
  "Legal_Aid_Indicators", "Legal_Aid_Target"
)

# Create new folder named after the value in the project code column
project_name <- as.character(indicator_data[i, "Project Code"])

# Sanitize project name to remove invalid characters
project_name <- gsub("[^a-zA-Z0-9_\\-]", "_", project_name)

if (is.na(project_name) || project_name == "") {
  print(paste("Invalid project name in row:", i))
  next
}

print(paste("Creating folder:", project_name))
dir.create(project_name, showWarnings = FALSE, recursive = TRUE)

# Copy template files to the new folder
# Create the directory
dir.create(project_name, showWarnings = FALSE, recursive = TRUE)

# Define file paths
training_template_path <- "C:/Users/Admin/Downloads/Training - Training Table Template.xlsx"
disaggregations_path <- "C:/Users/Admin/Downloads/Disaggregations.xlsx"
indicators_path <- "C:/Users/Admin/Downloads/Indicators.xlsx"
indicators_master_path <- "C:/Users/Admin/Downloads/Indicatorsmaster.xlsx"
equipment_template_path <- "C:/Users/Admin/Downloads/Equipment - Equipment Table Template.xlsx"
interdiction_template_path <- "C:/Users/Admin/Downloads/Interdiction - Interdiction Table Template.xlsx"


# Check if the files exist
file_paths <- c(
  training_template_path,
  disaggregations_path,
  indicators_path,
  equipment_template_path,
  interdiction_template_path
)

for (path in file_paths) {
  if (!file.exists(path)) {
    print(paste("File does not exist:", path))
  }
}

# Check if the directory exists
if (!dir.exists(project_name)) {
  dir.create(project_name, recursive = TRUE)
}

# Copy template files to the new folder
file_paths_to_copy <- c(
  training_template_path,
  disaggregations_path,
  indicators_path,
  equipment_template_path,
  interdiction_template_path
)

file_destinations <- c(
  file.path(project_name, "Training - Training Table Template.xlsx"),
  file.path(project_name, "Disaggregations.xlsx"),
  file.path(project_name, "Indicators.xlsx"),
  file.path(project_name, "Equipment - Equipment Table Template.xlsx"),
  file.path(project_name, "interdiction - Interdiction Table Template.xlsx")
)

for (i in seq_along(file_paths_to_copy)) {
  success <- file.copy(file_paths_to_copy[i], file_destinations[i], overwrite = TRUE)
  if (!success) {
    print(paste("Failed to copy file:", file_paths_to_copy[i], "to", file_destinations[i]))
  } else {
    print(paste("Successfully copied file:", file_paths_to_copy[i], "to", file_destinations[i]))
  }
}
# Read the copied training template
training_template <- read_excel(file.path(project_name, "Training - Training Table Template.xlsx"))
disaggregations_file <- file.path(project_name, "Disaggregations.xlsx")
indicators_file <- file.path(project_name, "Indicators.xlsx")
equipment_template <- file.path(project_name, "Equipment - Equipment Table Template.xlsx")
interdiction_template <- file.path(project_name, "interdiction - Interdiction Table Template.xlsx")

# Define the remove_columns function if it's not already defined
remove_columns <- function(df, cols) {
  df <- df[, !(names(df) %in% cols)]
  return(df)
}


# Conditions for training results
library(openxlsx)

# Define the path to your data and templates
indicator_data_path <- "C:/Users/Admin/Downloads/Standard Indicator Form.xlsx"
training_template_path <- "C:/Users/Admin/Downloads/Training - Training Table Template.xlsx"
output_dir <- "C:/Users/Admin/Downloads/Training_Results"

# Load the indicator data
indicator_data <- read.xlsx(indicator_data_path)
training_data <- read.xlsx(training_template_path, sheet = 1)

# Rename columns
colnames(indicator_data) <- c(
  "ID", "Start_time", "Completion_time", "Email", "Name", 
  "Read_Note", "Contact_Name", "Contact_Email", "Project Code", 
  "Formal Training", "Training_Info", "Training Mode", "Training Type", 
  "Country Scope", "Participant_Countries", "Course Topic", 
  "Topic Scope", "Course_List", "Course Name", "Participant Sector", 
  "Participant Agency/Office/Organization", "Evaluation_Indicators", "Participant Anonymity", 
  "Certifications", "Certificate Type", "Total_Participants", 
  "Total_Events", "In_Person_Events", "Event_Info", "International_Participants", 
  "GPP_Funding", "Event_Types", "Total_Events_Target", "Equipment_Donation", 
  "Equipment_Donation_Info", "Donation_Countries","Benefit_Countries","Equipment Type", 
  "Beneficiary Type", "Total_Donation_Value", "Construction_Activities", 
  "Construction_Locations", "Interdiction_Activities", "Interdiction_Info", 
  "Type of Goods Seized", "Interdiction Location Type", "Specific_Locations", 
  "Search Type", "Track_Searches", "Search_Target", 
  "Track_Flags", "Flag_Target", "Wildlife_Trafficking", 
  "Arrests_Prosecutions", "Arrests_Target", "Anti_Corruption_Indicators", 
  "Anti_Corruption_Target", "Improved_Case_Management", "Case_Management_Target", 
  "Legal_Aid_Indicators", "Legal_Aid_Target"
)

# Print renamed columns
print(colnames(indicator_data))
print(colnames(training_data))

# Define column indexes
project_code_index <- which(colnames(indicator_data) == "Project Code")
formal_training_index <- which(colnames(indicator_data) == "Formal Training")
training_type_index <- which(colnames(indicator_data) == "Training Type")
training_mode_index <- which(colnames(indicator_data) == "Training Mode")
country_scope_index <- which(colnames(indicator_data) == "Country Scope")
topic_scope_index <- which(colnames(indicator_data) == "Topic Scope")
certifications_index <- which(colnames(indicator_data) == "Certificate Type")
participant_anonymity_index <- which(colnames(indicator_data) == "Participant Anonymity")

# Function to remove columns
remove_columns <- function(template, columns) {
  colnames_to_remove <- intersect(colnames(template), columns)
  if (length(colnames_to_remove) > 0) {
    template <- template[, !(colnames(template) %in% colnames_to_remove)]
  }
  return(template)
}

# Process each row of the indicator_data
# Loop through each row in indicator_data
for (i in 1:nrow(indicator_data)) {
  row_data <- indicator_data[i, ]
  project_code_index <- row_data[["Project Code"]]
  
  
  
  # Load the training template
  if (file.exists(training_template_path)) {
    training_template <- read.xlsx(training_template_path)
  } else {
    print(paste("Training template file not found:", training_template_path))
    next
  }
  
  # Print the initial state of the template
  print(paste("Initial template columns for row", i, ":", paste(colnames(training_template), collapse = ", ")))
  
  
  if (!is.na(indicator_data[i, formal_training_index]) && indicator_data[i, formal_training_index] == "Yes") {
    # Process Training Table Template based on column values
    project_code_index <- which(colnames(indicator_data) == "Project Code")
    formal_training_index <- which(colnames(indicator_data) == "Formal Training")
    training_type_index <- which(colnames(indicator_data) == "Training Type")
    training_mode_index <- which(colnames(indicator_data) == "Training Mode")
    country_scope_index <- which(colnames(indicator_data) == "Country Scope")
    topic_scope_index <- which(colnames(indicator_data) == "Topic Scope")
    certifications_index <- which(colnames(indicator_data) == "Certificate Type")
    participant_anonymity_index <- which(colnames(indicator_data) == "Participant Anonymity")
    
    
    # Debugging print statements
    print(paste("Processing col:###", j))
    print(paste("Training Type:", indicator_data[i, training_type_index]))
    print(paste("Training Mode:", indicator_data[i, training_mode_index]))
    print(paste("Country Scope:", indicator_data[i, country_scope_index]))
    print(paste("Topic Scope:", indicator_data[i, topic_scope_index]))
    print(paste("Certifications:", indicator_data[i, certifications_index]))
    print(paste("Participant Anonymity:", indicator_data[i, participant_anonymity_index]))
    
    if (!is.na(indicator_data[i, training_type_index]) && indicator_data[i, training_type_index] %in% c("In-person training", "Virtual training")) {
      training_template <- remove_columns(training_template, c("Instruction.Mode.-.Instructor-led.or.Self-paced"))
    }
    if (!is.na(indicator_data[i, training_mode_index]) && indicator_data[i, training_mode_index] %in% c("In-person training", "Virtual training")) {
      print(paste("**** matched ***"))
      training_template <- remove_columns(training_template, c("Course.Mode.-.In-person,.Virtual,.or.Hybrid"))
    }
    if (!is.na(indicator_data[i, country_scope_index]) && indicator_data[i, country_scope_index] == "One country") {
      training_template <- remove_columns(training_template, c("Participant.Work.Location.-.Country"))
    }
    if (!is.na(indicator_data[i, country_scope_index]) && indicator_data[i, country_scope_index] == "Multiple countries") {
      training_template <- remove_columns(training_template, c("Participant.Work.Location.-.[Admin.1]"))
    }
    if (!is.na(indicator_data[i, topic_scope_index]) && indicator_data [i, topic_scope_index] == "One topic only") {
      training_template <- remove_columns(training_template, c("Course.Topic.-.Primary", "Course.Topic.-.Secondary"))
    }
    if (!is.na(indicator_data[i, topic_scope_index]) && indicator_data[i, topic_scope_index] == "Multiple topics") {
      training_template <- remove_columns(training_template, c("Course.Topic"))
    }
    if (is.na(indicator_data[i, certifications_index]) || indicator_data[i, certifications_index] == "Certificates of completion") {
      training_template <- remove_columns(training_template, c("Certification.Received"))
    }
    if (!is.na(indicator_data[i, participant_anonymity_index]) && indicator_data[i, participant_anonymity_index] == "Anonymous") {
      training_template <- remove_columns(training_template, c("Pre-Test.Score", "Post-Test.Score", "Assessment.-.Pass.or.Fail", "Course.Relevance.Score", "Course.Satisfaction.Score", "Instructor.Satisfaction.Score"))
    }
    # Print the modified state of the template
    print(paste("Modified template columns for row", i, ":", paste(colnames(training_template), collapse = ", ")))
    
    # Save the modified training template
    output_file <- file.path(output_dir, paste0(indicator_data[i, project_code_index], "_Training_Table_Template.xlsx"))
    write.xlsx(training_template, output_file, overwrite = TRUE)
    print(paste("Saved modified training template to:", output_file))
  } else {
    # Skip processing if formal training is not "Yes"
    print(paste("Skipping row", i, "as Formal Training is not 'Yes'"))
  }
}

# Load required libraries
library(readxl)
library(openxlsx)
library(dplyr)

# Define file paths
standard_indicator_path <- "C:/Users/Admin/Downloads/Standard Indicator Form.xlsx"
output_directory <- "C:/Users/Admin/Downloads/Disaggregations_Results"

# Create the output directory if it does not exist
if (!dir.exists(output_directory)) {
  dir.create(output_directory)
}

# Read data
indicator_data <- read_excel(standard_indicator_path)

# Rename columns
colnames(indicator_data) <- c(
  "ID", "Start_time", "Completion_time", "Email", "Name", 
  "Read_Note", "Contact_Name", "Contact_Email", "Project Code", 
  "Formal Training", "Training_Info", "Training Mode", "Training Type", 
  "Country Scope", "Participant_Countries", "Course Topic", 
  "Topic Scope", "Course_List", "Course Name", "Participant Sector", 
  "Participant Agency/Office/Organization", "Evaluation_Indicators", "Participant Anonymity", 
  "Certifications", "Certificate Type", "Total_Participants", 
  "Total_Events", "In_Person_Events", "Event_Info", "International_Participants", 
  "GPP_Funding", "Event_Types", "Total_Events_Target", "Equipment_Donation", 
  "Equipment_Donation_Info", "Donation_Countries","Benefit_Countries","Equipment Type", 
  "Beneficiary Type", "Total_Donation_Value", "Construction_Activities", 
  "Construction_Locations", "Interdiction_Activities", "Interdiction_Info", 
  "Type of Goods Seized", "Interdiction Location Type", "Specific_Locations", 
  "Search Type", "Track_Searches", "Search_Target", 
  "Track_Flags", "Flag_Target", "Wildlife_Trafficking", 
  "Arrests_Prosecutions", "Arrests_Target", "Anti_Corruption_Indicators", 
  "Anti_Corruption_Target", "Improved_Case_Management", "Case_Management_Target", 
  "Legal_Aid_Indicators", "Legal_Aid_Target"
)

# Print renamed columns
print(colnames(indicator_data))

# Define the columns of interest for disaggregation
columns_of_interest <- c(
  "Course Topic", "Course Name", "Participant Sector", 
  "Participant Agency/Office/Organization", "Certificate Type", 
  "Equipment Type", "Beneficiary Type", "Type of Goods Seized", 
  "Interdiction Location Type", "Search Type"
)

# Function for disaggregations using row data and column names
generate_disaggregation_file <- function(row_data) {
  project_code <- row_data[["Project Code"]]
  disaggregation_results <- data.frame()
  
  disaggregations <- function(column_name, disaggregation_suffix, value_column) {
    if (!is.na(row_data[[value_column]]) && 
        !(value_column == "Certificate Type" && row_data[[value_column]] == "Certificates of completion")) {
      full_name <- paste(project_code, "-", disaggregation_suffix)
      disaggregation_data <- data.frame(
        Disaggregation = full_name, 
        Disaggregation_Categories = row_data[[value_column]]
      )
      disaggregation_results <<- bind_rows(disaggregation_results, disaggregation_data)
    }
  }
  
  # Apply disaggregations for each column of interest
  for (column in columns_of_interest) {
    disaggregations(column, column, column)
  }
  
  # Only save if there are disaggregation results
  if (nrow(disaggregation_results) > 0) {
    # Define file path for the project-specific disaggregations file
    disaggregation_file_path <- file.path(output_directory, paste0(project_code, "_Disaggregations.xlsx"))
    
    # Create a new workbook and add a sheet named "Disaggregations"
    workbook <- createWorkbook()
    addWorksheet(workbook, "Disaggregations")
    writeData(workbook, sheet = "Disaggregations", disaggregation_results)
    
    # Save the workbook
    saveWorkbook(workbook, disaggregation_file_path, overwrite = TRUE)
  }
}

# Process each row of the indicator data
for (i in 1:nrow(indicator_data)) {
  row_data <- indicator_data[i, ]
  generate_disaggregation_file(row_data)
}

message("Disaggregations saved to: ", output_directory)

# End of Disaggregations



#Start of Indicators
# Load Libraries
library(readxl)
library(writexl)
library(dplyr)

# Define paths
standard_indicator_path <- "C:/Users/Admin/Downloads/Standard Indicator Form.xlsx"
indicators_master_path <- "C:/Users/Admin/Downloads/Indicatorsmaster.xlsx"
output_directory <- "C:/Users/Admin/Downloads/Indicator_Results"

# Create the output directory if it does not exist
if (!dir.exists(output_directory)) {
  dir.create(output_directory)
}

# Read data
indicator_data <- read_excel(standard_indicator_path)
indicators_master <- read_excel(indicators_master_path)

# Rename columns
colnames(indicator_data) <- c(
  "ID", "Start_time", "Completion_time", "Email", "Name", 
  "Read_Note", "Contact_Name", "Contact_Email", "Project Code", 
  "Formal Training", "Training_Info", "Training Mode", "Training Type", 
  "Country Scope", "Participant_Countries", "Course Topic", 
  "Topic Scope", "Course_List", "Course Name", "Participant Sector", 
  "Participant Agency/Office/Organization", "Evaluation_Indicators", "Participant Anonymity", 
  "Certifications", "Certificate Type", "Total_Participants", 
  "Total_Events", "In_Person_Events", "Event_Info", "International_Participants", 
  "GPP_Funding", "Event_Types", "Total_Events_Target", "Equipment_Donation", 
  "Equipment_Donation_Info", "Donation_Countries", "Benefit_Countries", "Equipment Type", 
  "Beneficiary Type", "Total_Donation_Value", "Construction_Activities", 
  "Construction_Locations", "Interdiction_Activities", "Interdiction_Info", 
  "Type of Goods Seized", "Interdiction Location Type", "Specific_Locations", 
  "Search Type", "Track_Searches", "Search_Target", 
  "Track_Flags", "Flag_Target", "Wildlife_Trafficking", 
  "Arrests_Prosecutions", "Arrests_Target", "Anti_Corruption_Indicators", 
  "Anti_Corruption_Target", "Improved_Case_Management", "Case_Management_Target", 
  "Legal_Aid_Indicators", "Legal_Aid_Target"
)

# Print renamed columns for debugging
print(colnames(indicator_data))

# Initialize the output dataframe
indicators <- tibble()

# Keywords checks for Indicators
keywords <- c("narcotics", "wildlife", "environmental products", "persons", "arms", "money laundering", "counterfeiting", "smuggling", "cybercrime", "cyber", "trafficking", "laundering", "counter", "money", "TOC", "International")

# Loop through each row in indicator_data
for (i in 1:nrow(indicator_data)) {
  row_data <- indicator_data[i, ]
  project_code <- row_data[["Project Code"]]
  
  # Initialize a dataframe for the current project
  project_indicators <- tibble()
  
  # General keyword check
  if (any(sapply(keywords, function(kw) grepl(kw, row_data[["Course Topic"]], ignore.case = TRUE)))) {
    project_indicators <- bind_rows(project_indicators, indicators_master[1, ])
  }
  
  # Additional specific keyword checks for Indicators
  if (grepl("cybercrime", row_data[["Course Topic"]], ignore.case = TRUE)) {
    project_indicators <- bind_rows(project_indicators, indicators_master[18, ])
  }
  
  if (grepl("Justice (prosecutors and court officials)", row_data[["Participant Sector"]], ignore.case = TRUE)) {
    project_indicators <- bind_rows(project_indicators, indicators_master[2, ])
  }
  
  if (grepl("Government Officials", row_data[["Participant Sector"]], ignore.case = TRUE) && grepl("Anti-corruption", row_data[["Course Topic"]], ignore.case = TRUE)) {
    project_indicators <- bind_rows(project_indicators, indicators_master[3, ])
  }
  
  if ((grepl("Civil Society", row_data[["Participant Sector"]], ignore.case = TRUE) || grepl("Business / Private sector", row_data[["Participant Sector"]], ignore.case = TRUE)) && grepl("Anti-corruption", row_data[["Course Topic"]], ignore.case = TRUE)) {
    project_indicators <- bind_rows(project_indicators, indicators_master[4, ])
  }
  
  if (grepl("Average difference in pre- and post-test scores", row_data[["Evaluation_Indicators"]])) {
    project_indicators <- bind_rows(project_indicators, indicators_master[9, ])
  }
  
  if (grepl("Average score given by participants on how relevant they feel the training", row_data[["Evaluation_Indicators"]])) {
    project_indicators <- bind_rows(project_indicators, indicators_master[6, ])
  }
  
  if (grepl("Average score given by participants on their satisfaction with the course materials", row_data[["Evaluation_Indicators"]])) {
    project_indicators <- bind_rows(project_indicators, indicators_master[7, ])
  }
  
  if (grepl("Average score given by participants on their satisfaction with the instructor", row_data[["Evaluation_Indicators"]])) {
    project_indicators <- bind_rows(project_indicators, indicators_master[8, ])
  }
  
  if (grepl("Percentage who pass the post-test", row_data[["Evaluation_Indicators"]])) {
    project_indicators <- bind_rows(project_indicators, indicators_master[10, ])
  }
  
  if (grepl("Percentage who pass the pass/fail assessment", row_data[["Evaluation_Indicators"]])) {
    project_indicators <- bind_rows(project_indicators, indicators_master[11, ])
  }
  
  # Check column Z
  if (!is.na(row_data[["Total_Participants"]])) {
    project_indicators <- bind_rows(project_indicators, indicators_master[12, ])
  }
  
  # Check column AA
  if (!is.na(row_data[["Total_Events"]])) {
    project_indicators <- bind_rows(project_indicators, indicators_master[13, ])
  }
  
  # Check column AH
  if (!is.na(row_data[["Equipment_Donation"]]) && row_data[["Equipment_Donation"]] == "Yes") {
    # Define file path for the project-specific Equipment Table Template
    equipment_file_path <- file.path(output_directory, paste0(project_code, "_Equipment - Equipment Table Template.xlsx"))
    
    # Copy the Equipment Table Template to the project-specific file if it does not exist
    if (!file.exists(equipment_file_path)) {
      file.copy("C:/Users/Admin/Downloads/Equipment - Equipment Table Template.xlsx", equipment_file_path, overwrite = TRUE)
    }
    
    # Add row 15 of IndicatorsMaster.xlsx to project_indicators
    project_indicators <- bind_rows(project_indicators, indicators_master[14, ])
    
    # Write the project_indicators to the output file
    project_file_path <- file.path(output_directory, paste0(project_code, "_Indicators.xlsx"))
    write_xlsx(list(
      Indicators = project_indicators,
      Equipment = read_excel("C:/Users/Admin/Downloads/Equipment - Equipment Table Template.xlsx")
    ), project_file_path)
    
    # Print message for debugging
    cat("Saved file with Equipment Table Template:", project_file_path, "\n")
  } else {
    equipment_file_path <- file.path(output_directory, paste0(project_code, "_Equipment - Equipment Table Template.xlsx"))
    if (file.exists(equipment_file_path)) {
      file.remove(equipment_file_path)
    }
  }
  
  # Check column AQ
  if (!is.na(row_data[["Interdiction_Activities"]]) && row_data[["Interdiction_Activities"]] == "Yes") {
    # Define file path for the project-specific Interdiction Table Template
    interdiction_file_path <- file.path(output_directory, paste0(project_code, "_Interdiction - Interdiction Table Template.xlsx"))
    
    # Copy the Interdiction Table Template to the project-specific file if it does not exist
    if (!file.exists(interdiction_file_path)) {
      file.copy("C:/Users/Admin/Downloads/Interdiction - Interdiction Table Template.xlsx", interdiction_file_path, overwrite = TRUE)
    }
  }
  # Add rows 16, 17, 18, 19, and 20 of IndicatorsMaster.xlsx to project_indicators
  project_indicators <- bind_rows(project_indicators, indicators_master[15:19, ])
  
  # Only save if there are indicators
  if (nrow(project_indicators) > 0) {
    # Copy row "I" from Standard Indicator Form to row "B" in project_indicators
    if (ncol(row_data) >= 9 && ncol(project_indicators) >= 2) {
      project_indicators[[2]] <- row_data[[9]]
    }
    
    # Copy row "S" in project_indicators with prefix "P/"
    if (ncol(project_indicators) >= 19) {
      project_indicators[[19]] <- paste0("P/", project_indicators[[19]])
    }
    
    # Define file path for the project-specific indicators file
    project_file_path <- file.path(output_directory, paste0(project_code, "_Indicators.xlsx"))
    
    # Save the workbook
    write_xlsx(project_indicators, project_file_path)
    
    # Print message for debugging
    cat("Saved file:", project_file_path, "\n")
  }
}

print(paste("Completed row:", i))


print("Completed processing all rows")
#End of Indicators


print(paste("Completed row:", i))


print("Completed processing all rows")

#Schedule the document 
library(taskscheduleR)

# Define the path to the R script
script_path <- "C:/Users/Admin/Downloads/Project_script.R"

# Format the start date as "mm/dd/yyyy"
start_date <- format(Sys.Date(), "%m/%d/%Y")

# Schedule the task
taskscheduler_create(taskname = "INL_Project_Processing",
                     rscript = script_path,
                     schedule = "DAILY",
                     starttime = "00:00",
                     startdate = start_date,  # Use formatted start date
                     modifier = 1,
                     Rexe = file.path(R.home("bin"), "Rscript.exe"))
