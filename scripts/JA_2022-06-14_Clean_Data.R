#--------------------------------------------------------#
#           Read in excel data and clean
#--------------------------------------------------------#

# Imports
library(dplyr)
library(readxl)

# CHANGE ME
batch_num = 5
compound_type = "Dioxins"
# Reading file is annoying because we need forward slash '/' not backslash '\'
# Option 1: readline then convert with gsub
x = readline() # REMOVE QUOTES
read_file = gsub("\\\\", "/", x)
# Option 2: https://onlinetexttools.com/replace-text
# read_file = "D:/Summer DRSI/Pollutant Concentration Data/Donor R322/Donor R322_Adipose_OC Pesticides_Part 1.xlsx"

# Load the excel file
data = read_excel(read_file) %>% as.data.frame(data)

# Function for cleaning the data. I used grepl for string matching
cleanData = function(data) {
  # Get specific columns
  new_df = data %>% select(c(SAMPLE_NO, COMPOUND, LAB_FLAG, CONC_FOUND, DETECTION_LIMIT, UNIT, SAMPLE_SIZE, SAMPLE_SIZE_UNIT))
  
  # Merge sample size with units
  new_df$SAMPLE_SIZE = paste(new_df$SAMPLE_SIZE, new_df$SAMPLE_SIZE_UNIT)
  # Delete sample_size_unit
  new_df = subset(new_df, select = -c(SAMPLE_SIZE_UNIT))
  
  # Delete all TEQ
  new_df = new_df %>% filter(!grepl('TEQ', COMPOUND))
  
  # Delete all '% Recovery'
  new_df = new_df %>% filter(!grepl('% Recovery', UNIT))
  
  # Delete 4000islets
  new_df = new_df %>% filter(!grepl('4000', SAMPLE_NO))
  
  # Delete 1500islets
  new_df = new_df %>% filter(!grepl('1500', SAMPLE_NO))
  
  # Add batch number to 'Lab Blank' row
  new_df = new_df %>% mutate(SAMPLE_NO = replace(SAMPLE_NO, SAMPLE_NO == "Lab Blank", paste("Lab Blank Batch", batch_num, compound_type)))

  new_df
}

# Function for sorting
sortData = function(new_df) {
   # Ordering by custom factor
   if (compound_type == "Dioxins") {
     compound_list = c('1,2,3,7,8,9-HXCDD (225)', '2,3,7,8-TCDF (225)', '2,3,7,8-TCDD','1,2,3,7,8-PECDD','1,2,3,4,7,8-HXCDD','1,2,3,6,7,8-HXCDD','1,2,3,7,8,9-HXCDD','1,2,3,4,6,7,8-HPCDD','OCDD','2,3,7,8-TCDF','1,2,3,7,8-PECDF','2,3,4,7,8-PECDF','1,2,3,4,7,8-HXCDF','1,2,3,6,7,8-HXCDF','1,2,3,7,8,9-HXCDF','2,3,4,6,7,8-HXCDF','1,2,3,4,6,7,8-HPCDF','1,2,3,4,7,8,9-HPCDF','OCDF','TOTAL TETRA-DIOXINS','TOTAL PENTA-DIOXINS','TOTAL HEXA-DIOXINS','TOTAL HEPTA-DIOXINS','TOTAL TETRA-FURANS','TOTAL PENTA-FURANS','TOTAL HEXA-FURANS','TOTAL HEPTA-FURANS')
   } else if (compound_type == "Pesticides") {
     compound_list = c("Hexachlorobenzene","HCH, alpha","HCH, beta","HCH, gamma","HCH, delta","Heptachlor","Aldrin","Heptachlor Epoxide","Chlordane, oxy-","Chlordane, gamma (trans)","Chlordane, alpha (cis)","Nonachlor, trans-","Nonachlor, cis-","alpha-Endosulphan","beta-Endosulphan","Dieldrin","Endrin","Endrin Aldehyde","Endrin Ketone","Endosulphan Sulphate","2,4'-DDE","4,4'-DDE","2,4'-DDD","4,4'-DDD","2,4'-DDT","4,4'-DDT","Methoxychlor","Mirex")
   } else if (compound_type == "PCB") {
     compound_list = c("3,3',4,4'-TeCB","3,4,4',5-TeCB","2,3,3',4,4'-PeCB","2,3,4,4',5-PeCB","2,3',4,4',5-PeCB","2',3,4,4',5-PeCB","3,3',4,4',5-PeCB", "2,2',4,4',6,6'-HxCB", "2,3,3',4,4',5-HxCB", "2,3,3',4,4',5'-HxCB","2,3',4,4',5,5'-HxCB","3,3',4,4',5,5'-HxCB","2,2',3,3',4,4',5-HpCB","2,2',3,4,4',5,5'-HpCB","2,3,3',4,4',5,5'-HpCB","2,3,3',4',5,5',6-HpCB")
   }
   
   new_df$COMPOUND <- factor(new_df$COMPOUND, levels = compound_list)
   new_df <- new_df[order(new_df$SAMPLE_NO, new_df$COMPOUND),]
   
   new_df
 }

new_df = cleanData(data)
new_df = sortData(new_df)

#--------------------------------------------------------#
#                    Export excel data
#--------------------------------------------------------#

# Imports
library(xlsx)

# Write to excel
writeFile = paste("JA_", Sys.Date(), "_CleanedData_Batch", batch_num, "_", compound_type, ".xlsx", sep="")
write.xlsx(new_df, writeFile, sheetName = "Cleaned Data", col.names = TRUE, row.names = FALSE, append = FALSE, showNA = FALSE)
