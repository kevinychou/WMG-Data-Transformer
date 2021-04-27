library(readxl)
library(writexl)
library(plyr)
library(tibble)
library(tidyverse)

xl_data <- "C:\\Users\\kevin\\Desktop\\MRSS #3400s Data.xlsx" # "MRSS #3400s Data.xlsx", "Test_Set.xlsx" 
template <- "C:\\Users\\kevin\\Desktop\\Template.xlsx"

# Makes initial template dataframe
# WIDE FORMAT: df <- read_excel(path = xl_data, sheet = "Template_2") 
# df <- read_excel(path = xl_data, sheet = "Header") #NOTE: Need to include Header title in each excel
df <- as.data.frame(matrix(ncol = 4, nrow = 0))
colnames(df) <- c("Code", "Family", "Rater 1", "Name of Coders")

# List of Sheets
sheet_names <- excel_sheets(path = xl_data) # List of sheet names
sheet_list <- lapply(sheet_names, function(x){read_excel(path = xl_data, sheet = x, col_names = FALSE)}) # List containing data in each sheet


# Defining functions
get_freq <- function(category, df) {
  value = df %>%
    filter(...3 == category) %>%
    select(...4)
  
  if (nrow(value) == 0) {
    value = 0
  }
  
  return(value[[1]])
}

get_dur <- function(category, df) {
  value = df %>%
    filter(...3 == category) %>%
    select(...5)
  
  if (nrow(value) == 0) {
    value = 0
  }
  
  return(value[[1]])
}


# Adding new sheet data



for (sheet in sheet_list) {
  
  if (str_detect(sheet[1,2], "FP")) { ## Check FP/SF/RE
  
    new_df <- read_excel(path = template)
    
    new_df$Family = substr(sheet[1,2], start = 1, stop = 4) # Add family code in
    
    # Gaze
    new_df$`Rater 1`[1] = get_freq("Looks at infant", sheet) 
    new_df$`Rater 1`[2] = get_dur("Looks at infant", sheet)
    new_df$`Rater 1`[3] = get_freq("Looks at object", sheet) 
    new_df$`Rater 1`[4] = get_dur("Looks at object", sheet)
    new_df$`Rater 1`[5] = get_freq("Avert", sheet) 
    new_df$`Rater 1`[6] = get_dur("Avert", sheet)
    new_df$`Rater 1`[7] = get_freq("Avert gaze from game", sheet) 
    new_df$`Rater 1`[8] = get_dur("Avert gaze from game", sheet)
    
    # Proximity
    new_df$`Rater 1`[9] = get_freq("Lean in", sheet) 
    new_df$`Rater 1`[10] = get_dur("Lean in", sheet)
    new_df$`Rater 1`[11] = get_freq("Nose to nose", sheet) 
    new_df$`Rater 1`[12] = get_dur("Nose to nose", sheet)
    
    # Touch
    new_df$`Rater 1`[13] = (sheet %>% 
      filter(...2 == "Yes Touch") %>%
      select(...4)) [[1]]
    
    # Elicits
    new_df$`Rater 1`[14] = get_freq("Wave", sheet) 
    new_df$`Rater 1`[15] = get_dur("Wave", sheet)
    new_df$`Rater 1`[16] = get_freq("Making noise using hands or fingers", sheet) 
    new_df$`Rater 1`[17] = get_dur("Making noise using hands or fingers", sheet)
    new_df$`Rater 1`[18] = get_freq("Reposition self", sheet) ##CHECK THIS CODE NAME
    new_df$`Rater 1`[19] = get_dur("Reposition self", sheet)
    new_df$`Rater 1`[20] = get_freq("Blow", sheet) 
    new_df$`Rater 1`[21] = get_dur("Blow", sheet)
    
    # Vocalisations
    new_df$`Rater 1`[22] = get_freq("Praises_Compliments_Positive vocalisation", sheet) 
    new_df$`Rater 1`[23] = get_freq("Positive attributions to infant", sheet)
    new_df$`Rater 1`[24] = get_freq("Positive infant state description", sheet) 
    new_df$`Rater 1`[25] = get_freq("Positive description of own state", sheet)
    new_df$`Rater 1`[26] = get_freq("Criticises/Negative Vocalisation", sheet) 
    new_df$`Rater 1`[27] = get_freq("Negative attributions to infant", sheet)
    new_df$`Rater 1`[28] = get_freq("Negative infant state description", sheet) 
    new_df$`Rater 1`[29] = get_freq("Negative description of own state", sheet)
    new_df$`Rater 1`[30] = get_freq("Directs infant to self", sheet) 
    new_df$`Rater 1`[31] = get_freq("Sings", sheet)
    new_df$`Rater 1`[32] = get_dur("Sings", sheet) 
    new_df$`Rater 1`[33] = get_freq("Describing objects", sheet)##UNSURE IF INFANT ACTIONS IS NEEDED
    new_df$`Rater 1`[34] = get_freq("Mimics Infant", sheet)
    new_df$`Rater 1`[35] = get_freq("Mouth noises", sheet)
    
    df = rbind(df, new_df)
  } else {
    next
  }
  
} 


# target = "Output_Data.xlsx"
target = "C:\\Users\\kevin\\Desktop\\Output_Data.xlsx"

# Print out data
write_xlsx(df, target)





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
