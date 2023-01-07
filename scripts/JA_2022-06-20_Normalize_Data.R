#--------------------------------------------------------#
#                       Imports  
#--------------------------------------------------------#
library(readxl)
library(dplyr)
library(tidyr)
library(tidyverse)
library(xlsx)

#--------------------------------------------------------#
#                Read in excel data
#--------------------------------------------------------#

# Read in data and change NA to 0
data = read_excel("JA_2022-06-14_Compiled_Data.xlsx") %>% mutate(CONC_FOUND = replace_na(CONC_FOUND, 0))

#--------------------------------------------------------#
#               Subtract lab blanks
#--------------------------------------------------------#

# Compound list
dioxins_list = c('1,2,3,7,8,9-HXCDD (225)', '2,3,7,8-TCDF (225)', '2,3,7,8-TCDD','1,2,3,7,8-PECDD','1,2,3,4,7,8-HXCDD','1,2,3,6,7,8-HXCDD','1,2,3,7,8,9-HXCDD','1,2,3,4,6,7,8-HPCDD','OCDD','2,3,7,8-TCDF','1,2,3,7,8-PECDF','2,3,4,7,8-PECDF','1,2,3,4,7,8-HXCDF','1,2,3,6,7,8-HXCDF','1,2,3,7,8,9-HXCDF','2,3,4,6,7,8-HXCDF','1,2,3,4,6,7,8-HPCDF','1,2,3,4,7,8,9-HPCDF','OCDF','TOTAL TETRA-DIOXINS','TOTAL PENTA-DIOXINS','TOTAL HEXA-DIOXINS','TOTAL HEPTA-DIOXINS','TOTAL TETRA-FURANS','TOTAL PENTA-FURANS','TOTAL HEXA-FURANS','TOTAL HEPTA-FURANS')
dioxins_sum_list = c('1,2,3,7,8,9-HXCDD (225)', '2,3,7,8-TCDF (225)', '2,3,7,8-TCDD','1,2,3,7,8-PECDD','1,2,3,4,7,8-HXCDD','1,2,3,6,7,8-HXCDD','1,2,3,7,8,9-HXCDD','1,2,3,4,6,7,8-HPCDD','OCDD','2,3,7,8-TCDF','1,2,3,7,8-PECDF','2,3,4,7,8-PECDF','1,2,3,4,7,8-HXCDF','1,2,3,6,7,8-HXCDF','1,2,3,7,8,9-HXCDF','2,3,4,6,7,8-HXCDF','1,2,3,4,6,7,8-HPCDF','1,2,3,4,7,8,9-HPCDF','OCDF')
PCB_list = c("Hexachlorobenzene","HCH, alpha","HCH, beta","HCH, gamma","HCH, delta","Heptachlor","Aldrin","Heptachlor Epoxide","Chlordane, oxy-","Chlordane, gamma (trans)","Chlordane, alpha (cis)","Nonachlor, trans-","Nonachlor, cis-","alpha-Endosulphan","beta-Endosulphan","Dieldrin","Endrin","Endrin Aldehyde","Endrin Ketone","Endosulphan Sulphate","2,4'-DDE","4,4'-DDE","2,4'-DDD","4,4'-DDD","2,4'-DDT","4,4'-DDT","Methoxychlor","Mirex")
pesticides_list = c("3,3',4,4'-TeCB","3,4,4',5-TeCB","2,3,3',4,4'-PeCB","2,3,4,4',5-PeCB","2,3',4,4',5-PeCB","2',3,4,4',5-PeCB","3,3',4,4',5-PeCB", "2,2',4,4',6,6'-HxCB", "2,3,3',4,4',5-HxCB","2,3,3',4,4',5'-HxCB","2,3',4,4',5,5'-HxCB","3,3',4,4',5,5'-HxCB","2,2',3,3',4,4',5-HpCB","2,2',3,4,4',5,5'-HpCB","2,3,3',4,4',5,5'-HpCB","2,3,3',4',5,5',6-HpCB")

Normalize_Batch = function(batch) {
  
  # Separate tissues for this batch
  tissues = batch %>% filter(!grepl('Lab', SAMPLE_NO)) %>% group_split(SAMPLE_NO)
  
  # Normalize tissue function
  Normalize_Tissue = function(tissue) {
    #------- Dioxins -------#
    normalized_dioxin_df = tissue %>% filter(COMPOUND %in% dioxins_list)
    lab_blank_df = batch %>% filter(grepl('Lab Blank', SAMPLE_NO) & COMPOUND %in% dioxins_list)
    normalized_dioxin_df$CONC_FOUND = normalized_dioxin_df$CONC_FOUND - lab_blank_df$CONC_FOUND 
    # if negative concentration, set to 0
    normalized_dioxin_df$CONC_FOUND[normalized_dioxin_df$CONC_FOUND<0] <- 0
    
    #------- PCB -------#
    normalized_PCB_df = tissue %>% filter(COMPOUND %in% PCB_list)
    lab_blank_df = batch %>% filter(grepl('Lab Blank', SAMPLE_NO) & COMPOUND %in% PCB_list)
    normalized_PCB_df$CONC_FOUND = normalized_PCB_df$CONC_FOUND - lab_blank_df$CONC_FOUND
    # if negative concentration, set to 0
    normalized_PCB_df$CONC_FOUND[normalized_PCB_df$CONC_FOUND<0] <- 0
    
    #------- Pesticides -------#
    normalized_pest_df = tissue %>% filter(COMPOUND %in% pesticides_list)
    lab_blank_df = batch %>% filter(grepl('Lab Blank', SAMPLE_NO) & COMPOUND %in% pesticides_list)
    normalized_pest_df$CONC_FOUND = normalized_pest_df$CONC_FOUND - lab_blank_df$CONC_FOUND
    # if negative concentration, set to 0
    normalized_pest_df$CONC_FOUND[normalized_pest_df$CONC_FOUND<0] <- 0
    
    #------- Combine all compounds -------#
    normalized_tissue = bind_rows(normalized_dioxin_df, normalized_PCB_df, normalized_pest_df)
    
    #------- Append totals to bottom  -------#
    normalized_tissue = normalized_tissue %>% add_row(SAMPLE_NO = tissue[1,]$SAMPLE_NO, 
                                                            COMPOUND = "TOTAL CONCENTRATION - DIOXIN", 
                                                            CONC_FOUND = sum(normalized_dioxin_df %>% filter(COMPOUND %in% dioxins_sum_list) %>% select(CONC_FOUND))
    )
    normalized_tissue = normalized_tissue %>% add_row(SAMPLE_NO = tissue[1,]$SAMPLE_NO, 
                                                      COMPOUND = "TOTAL CONCENTRATION - PCB", 
                                                      CONC_FOUND = sum(normalized_PCB_df %>% filter(COMPOUND %in% PCB_list) %>% select(CONC_FOUND))
    )
    normalized_tissue = normalized_tissue %>% add_row(SAMPLE_NO = tissue[1,]$SAMPLE_NO, 
                                                        COMPOUND = "TOTAL CONCENTRATION - PESTICIDE", 
                                                        CONC_FOUND = sum(normalized_pest_df %>% filter(COMPOUND %in% pesticides_list) %>% select(CONC_FOUND))
    )
  }
  normalized_tissue_list = lapply(tissues, Normalize_Tissue)
  
  # Convert list to dataframe
  normalized_tissue_df = data.frame(Reduce(rbind, normalized_tissue_list))
  
  return(normalized_tissue_df)
}

# Separate data into batches
batch_1 = data %>% filter(grepl('R322|Batch 1', SAMPLE_NO))
batch_2 = data %>% filter(grepl('R361|R362|Batch 2', SAMPLE_NO))
batch_3 = data %>% filter(grepl('R382|R385|R389|R391|Batch 3', SAMPLE_NO))
batch_4 = data %>% filter(grepl('R393|R395|R399|R400|R401|R402|Batch 4', SAMPLE_NO))
batch_5 = data %>% filter(grepl('R408|R411|R412|R413|R417|R419|R421|Batch 5', SAMPLE_NO))
batch_list = list(batch_1, batch_2, batch_3, batch_4, batch_5)

# Normalize batches together
normalized_batches = lapply(batch_list, Normalize_Batch)

# Normalize by batches. Mainly for testing
# normalized_batch_1 = Normalize_Batch(batch_1) # 142 - good
# normalized_batch_2 = Normalize_Batch(batch_2) # 284 - good
# normalized_batch_3 = Normalize_Batch(batch_3) # 576 - good
# normalized_batch_4 = Normalize_Batch(batch_4) # 792 - good
# normalized_batch_5 = Normalize_Batch(batch_5) # 616 - good

#--------------------------------------------------------#
#             Export data into new spreadsheet
#--------------------------------------------------------#

# Bundle dataframes
# new_df = as.data.frame(rbind(normalized_batch_1, normalized_batch_2, normalized_batch_3, normalized_batch_4, normalized_batch_5))
new_df = data.frame(Reduce(rbind,normalized_batches))

# OPTION 1 - Write to excel w/o formatting. Use "=search" and condtional formatting
writeFile = paste("JA_", Sys.Date(), "_Normalized_Data.xlsx", sep="")
write.xlsx(new_df, writeFile, sheetName = "Normalized Data", col.names = TRUE, row.names = FALSE, append = FALSE, showNA = FALSE)

# OPTION 2 - Write to excel w/ highlights
library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, "containsText")
writeData(wb = wb, sheet = "containsText", x = new_df)
color <- createStyle(fgFill = "#FFFF00") #YELLOW 
addStyle(wb = wb, sheet = "containsText", style = color, rows = which(grepl(new_df$COMPOUND, pattern="TOTAL CONCENTRATION")) + 1, cols = 1:ncol(new_df), gridExpand = T)
openxlsx::saveWorkbook(wb = wb, file = "test.xlsx", overwrite = TRUE)

