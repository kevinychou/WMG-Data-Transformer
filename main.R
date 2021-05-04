library(readxl)
library(writexl)

library(plyr)
library(tibble)
library(tidyverse)

library(rJava)
library(xlsxjars)
library(xlsx)

## INPUT DATA ##

xl_data <- "C:\\Users\\kevin\\Documents\\WMG_Input_Output\\MRSS #3400s Data.xlsx" #Input excel name here
template <- "C:\\Users\\kevin\\Documents\\WMG_Input_Output\\Template.xlsx" #Input 'template' sheet for columns
template_SF <- "C:\\Users\\kevin\\Documents\\WMG_Input_Output\\Template_SF.xlsx"#Input 'template_SF' sheet for SF columns

xl_data <- "C:\\Users\\kevin\\Documents\\WMG_Input_Output\\Testing.xlsx" 

# Makes initial template dataframe

df_FP <- as.data.frame(matrix(ncol = 4, nrow = 0))
df_RE <- as.data.frame(matrix(ncol = 4, nrow = 0))
df_SF <- as.data.frame(matrix(ncol = 4, nrow = 0))
colnames(df_FP) <- c("Code", "Family", "Measure", "File Name")
colnames(df_RE) <- c("Code", "Family", "Measure", "File Name")
colnames(df_SF) <- c("Code", "Family", "Measure", "File Name")

# Make initial excel spreadsheets
wb <-createWorkbook(type="xlsx")
sheet <- createSheet(wb, sheetName ="Sheet_1")

saveWorkbook(wb, "MRSS FP Output.xlsx")
saveWorkbook(wb, "MRSS RE Output.xlsx")
saveWorkbook(wb, "MRSS SF Output.xlsx")

# List of Sheets
sheet_names <- excel_sheets(path = xl_data) # List of sheet names
sheet_list <- lapply(sheet_names, function(x){
  read_excel(path = xl_data, sheet = x, col_names = FALSE)}) # List containing data in each sheet


# Defining functions
get_freq <- function(category, df) {
  value = df %>%
    filter(...3 == category) %>%
    filter(...4 != "NA") %>%
    select(...4)
  
  if (nrow(value) == 0) {
    value = 0
  }
  
  return(sum(as.numeric(value[[1]])))
}

get_dur <- function(category, df) {
  value = df %>%
    filter(...3 == category) %>%
    filter(...5 != "NA") %>%
    select(...5)
  
  if (nrow(value) == 0) {
    value = 0
  }
  
  return(sum(as.numeric(value[[1]])))
}


# Adding new sheet data




for (sheet in sheet_list) {
  
  if (str_detect(sheet[1,2], "FP") | str_detect(sheet[1,2], "RE")) {
  
    new_df <- read_excel(path = template)
    
    new_df$`File Name` = sheet[[1,2]]
    new_df$Family = substr(sheet[1,2], start = 1, stop = 4) # Add family code in
    
    # Gaze
    new_df$`Measure`[1] = get_freq("Looks at infant", sheet) 
    new_df$`Measure`[2] = get_dur("Looks at infant", sheet)
    new_df$`Measure`[3] = get_freq("Looks at object", sheet) 
    new_df$`Measure`[4] = get_dur("Looks at object", sheet)
    new_df$`Measure`[5] = get_freq("Avert", sheet) 
    new_df$`Measure`[6] = get_dur("Avert", sheet)
    new_df$`Measure`[7] = get_freq("Avert gaze from game", sheet) 
    new_df$`Measure`[8] = get_dur("Avert gaze from game", sheet)
    
    # Proximity
    new_df$`Measure`[9] = get_freq("Lean in", sheet) 
    new_df$`Measure`[10] = get_dur("Lean in", sheet)
    new_df$`Measure`[11] = get_freq("Nose to nose", sheet) 
    new_df$`Measure`[12] = get_dur("Nose to nose", sheet)
    
    # Touch
    new_df$`Measure`[13] = (sheet %>% 
      filter(...2 == "Yes Touch") %>%
      select(...4)) [[1]]
    
    # Elicits
    new_df$`Measure`[14] = get_freq("Wave", sheet) 
    new_df$`Measure`[15] = get_dur("Wave", sheet)
    new_df$`Measure`[16] = get_freq("Making noise using hands or fingers", sheet) 
    new_df$`Measure`[17] = get_dur("Making noise using hands or fingers", sheet)
    new_df$`Measure`[18] = get_freq("Reposition self", sheet) ##CHECK THIS CODE NAME
    new_df$`Measure`[19] = get_dur("Reposition self", sheet)
    new_df$`Measure`[20] = get_freq("Blow", sheet) 
    new_df$`Measure`[21] = get_dur("Blow", sheet)
    new_df$`Measure`[22] = get_freq("Use object", sheet) 
    new_df$`Measure`[23] = get_dur("Use object", sheet)
    
    # Vocalisations
    new_df$`Measure`[24] = get_freq("Praises_Compliments_Positive vocalisation", sheet) 
    new_df$`Measure`[25] = get_freq("Positive attributions to infant", sheet)
    new_df$`Measure`[26] = get_freq("Positive infant state description", sheet) 
    new_df$`Measure`[27] = get_freq("Positive description of own state", sheet)
    new_df$`Measure`[28] = get_freq("Criticises/Negative Vocalisation", sheet) 
    new_df$`Measure`[29] = get_freq("Negative attributions to infant", sheet)
    new_df$`Measure`[30] = get_freq("Negative infant state description", sheet) 
    new_df$`Measure`[31] = get_freq("Negative description of own state", sheet)
    new_df$`Measure`[32] = get_freq("Directs infant to self", sheet) 
    new_df$`Measure`[33] = strtoi(get_freq("Describing objects", sheet)) + strtoi(get_freq("Labelling infant actions", sheet))
    new_df$`Measure`[34] = get_freq("Sings", sheet)
    new_df$`Measure`[35] = get_dur("Sings", sheet) 
    new_df$`Measure`[36] = get_freq("Mimics Infant", sheet)
    new_df$`Measure`[37] = get_dur("Mimics Infant", sheet)
    new_df$`Measure`[38] = get_freq("Mouth noises", sheet)
    new_df$`Measure`[39] = get_dur("Mouth noises", sheet) #### USE TESTING TO FIX STUFF WITH FREQ AND DURATION 
    
    if (str_detect(sheet[1,2], "FP")) {
      df_FP = rbind(df_FP, new_df)
    }
    else {
      df_RE = rbind(df_RE, new_df)
    }  
    
  } else if (str_detect(sheet[1,2], "SF")) {
    
    new_df <- read_excel(path = template_SF)
    
    new_df$`File Name` = sheet[[1,2]]
    new_df$Family = substr(sheet[1,2], start = 1, stop = 4)
    
    # Gaze
    new_df$`Measure`[1] = get_freq("Looks at infant", sheet) 
    new_df$`Measure`[2] = get_dur("Looks at infant", sheet)
    new_df$`Measure`[3] = get_freq("Looks at object", sheet) 
    new_df$`Measure`[4] = get_dur("Looks at object", sheet)
    new_df$`Measure`[5] = get_freq("Avert", sheet) 
    new_df$`Measure`[6] = get_dur("Avert", sheet)
    
    # Still Face Breaches
    new_df$`Measure`[7] = get_freq("Still Face Touch", sheet) 
    new_df$`Measure`[8] = get_dur("Still Face Touch", sheet)
    new_df$`Measure`[9] = get_freq("Still Face Elicit", sheet) 
    new_df$`Measure`[10] = get_dur("Still Face Elicit", sheet)
    new_df$`Measure`[11] = get_freq("Still Face Vocalisation", sheet) 
    new_df$`Measure`[12] = get_dur("Still Face Vocalisation", sheet)
    new_df$`Measure`[13] = get_freq("Still Face Vocalisation", sheet) 
    new_df$`Measure`[14] = get_dur("Still Face Vocalisation", sheet)
    
    
    df_SF = rbind(df_SF, new_df)
    
  } else {
    next
  }
  
} 


# Print out datasets #

target = "C:\\Users\\kevin\\Documents\\WMG_Input_Output\\MRSS FP Output.xlsx"
write_xlsx(df_FP, target)
target = "C:\\Users\\kevin\\Documents\\WMG_Input_Output\\MRSS RE Output.xlsx"
write_xlsx(df_RE, target)
target = "C:\\Users\\kevin\\Documents\\WMG_Input_Output\\MRSS SF Output.xlsx"
write_xlsx(df_SF, target)



# WIDE FORMAT: df <- read_excel(path = xl_data, sheet = "Template_2") 
# df <- read_excel(path = xl_data, sheet = "Header") #NOTE: Need to include Header title in each excel

# new_df = new_df %>% 
#   select(Family) %>% 
#   mutate(Family = substr(test_df1[1,2], start = 1, stop = 4))

# Wide Format Code
# df %>% add_row(
#   'Family Code' = substr(sheet[1,2], start = 1, stop = 4), # FAMILY ID: Takes the numbers from Observation ID
#   ### <Add Coder Name Code Here>
#   'Look at infant (freq)' = #
#   'Look at infant (dur)' =
#   colnames(df)[5] =
#   colnames(df)[6] = # Look at infant (freq)
#   colnames(df)[7] = # Look at infant (dur)
#   colnames(df)[8] =
#   colnames(df)[9] = # Look at infant (freq)
#   colnames(df)[10] = # Look at infant (dur)
#   colnames(df)[11] =
#   colnames(df)[12] = # Look at infant (freq)
#   colnames(df)[13] = # Look at infant (dur)
#   colnames(df)[14] =
# )
