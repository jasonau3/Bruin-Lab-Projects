library("readxl")
library(xlsx)
library(Hmisc)
library(tidyverse)
library(ppcor)

# read the sheet "C8898"
file = read_xlsx("2022-10-14_nanostring human islet R samples in final format (3).xlsx", sheet = "human islet R sample C8898")

# set the dataframe
df = data.frame(file)

# set the colnames 
colnames(df) = file[1,]

# delete first row
df = df[-c(1),]

# remove any NA rows since they are completely empty
df = df %>% drop_na()

# set the rownames
df = data.frame(df, row.names = 1)

# convert to numeric dataframe
covariance_df = sapply(df, as.numeric)
sapply(covariance_df, mode)

# set the rownames again
rownames(covariance_df) = rownames(df)

# put genes as columns
covariance_df = as.data.frame(t(covariance_df))

# remove any unneccessary dataframes
rm(file)
rm(df)

# # create a covariance matrix
# covar_CDH1 = cov(covariance_df, covariance_df$CDH1)
# covar = cov(covariance_df)
# # convert covariance to correlation 
# res = cov2cor(covar)
# 
# # generate p values from correlation using Hmisc
# res.p = Hmisc::rcorr(res, type="pearson")$P

# partial correlation analysis
genes_to_investigate = c("AHR", "ARNT", "CYP1A1", "G6PC2", "HIF1A", "LDHA", "VEGFA", "CYP2E1", "CYP3A4")

# make a dataframe for the correlation and p value
res.df = data.frame()

# loop through in a manner such as AHR with ARNT, CYP1A1, G6PC2.... then ARNT with CYPA1, G6PC2
# so it won't do AHR with ARNT then later ARNT with AHR
for (x in 1:length(genes_to_investigate)) {
  for (y in x:length(genes_to_investigate)) {
    # set the gene to investiagte
    gene_x = genes_to_investigate[x]
    gene_y = genes_to_investigate[y]
    # if the genes are equal, don't investigate
    if (gene_x == gene_y) {
      next
    }
    # calculate the partial correlation account for CDH1
    res = pcor.test(covariance_df[gene_x], covariance_df[gene_y], covariance_df$CDH1, method="spearman")
    # print the correlation and p value
    print(c(res$estimate, res$p.value))
    # save the result into res.df 
    res.df = rbind(res.df, c(res$estimate, res$p.value, paste(gene_x, "with", gene_y, "controlling for CDH1")))
  }
}

# rename the df nicely :)
names(res.df)[1] = "partial correlation estimate"
names(res.df)[2] = "p value"
names(res.df)[3] = "description"

# export to excel
write.xlsx(as.data.frame(res.df), file="JA_11-04-2022_partial_correlations_analysis.xlsx", sheetName = "correlation estimates", row.names = F)
