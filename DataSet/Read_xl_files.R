## Run Instructions
# 1. Keep all the xlsx and xls files in same directory
# 2. Copy the data direcory name and replace on "data_path" (ref- code line 8)
# 3. After changing the path, select all and run.
# 4. Merged data will be ready and cosolidated data will be stored in Global environment named as "Master_data"

## Set data file path - 2016 data
data_path="C:/Users/dubrangala/OneDrive - VMware, Inc/Project_PropModel/Materials/R Case Study 3 xlsx/Xlsx_data/data"
setwd(data_path)

## Just run the function - Dont change anything inside
read_excel_allsheets <- function(filename) {
  sheets <- readxl::excel_sheets(filename)
  mydata = data.frame()
  
    k=0
  cname = c()
  for(j in sheets) {
    k=k+1
    
    if (k==1){
      x = as.data.frame(readxl::read_excel(filename, sheet = j,col_names = T))
      x$SheetRef = j
      cname = c(cname,names(x))
    } else
    {
      x = as.data.frame(readxl::read_excel(filename, sheet = j,col_names = F))
      x$SheetRef = j
      names(x) = cname
    }
    mydata <- rbind(mydata,x)
    rm(x)
  }
 
return(mydata)
}

## Read Files from the path

xls_file <- dir(pattern = "xls")
xlsx <- dir(pattern = "xlsx")
xls <- xls_file[!xls_file %in% xlsx] 

## Join All xls file together
xls_data <- NULL
n=0
for(j in xls)
{
  n=n+1
  sheet_data <- read_excel_allsheets(j)
    fnam = stringr::str_replace(j,".xls","")
  sheet_data$Filename <- fnam
  xls_data <- rbind(xls_data,sheet_data)
  rm(sheet_data)
  if(n==length(xls)) {message("Data successfully extracted from 'xls' and consolidated together","\n")
    tab=table(xls_data$SheetRef,xls_data$Filename);return(print(tab))}
}

## Join All xls file together
xlsx_data <- NULL
n=0
for(j in xlsx)
{
  n=n+1
  fnam = stringr::str_replace(j,".xlsx","")
  sheet_data <- read_excel_allsheets(j)
  sheet_data$Filename <- fnam
  xlsx_data <- rbind(xlsx_data,sheet_data)
  rm(sheet_data)
  if(n==length(xlsx)) {message("Data successfully extracted from 'xlsx' and consolidated together","\n")
    tab=table(xlsx_data$SheetRef,xlsx_data$Filename);return(print(tab))}
}


## Mapping the column names
names_xls <- names(xls_data)
names_xls <- stringr::str_replace_all(names_xls," ","")
names(xls_data) <- names_xls

names_xlsx <- names(xlsx_data)
names_xlsx <- stringr::str_replace_all(names_xlsx," ","")
names(xlsx_data) <- names_xlsx

match_nam <-  intersect(names(xlsx_data),names(xls_data))
name_union <- union(names(xlsx_data),names(xls_data))

## Check manually to time column names, some names are common but R wont match properly becouse in between we can found sepcial cahracters
## for example in xls one variable name is "RetailTY" the same varible in xlsx is "Retail TY"
## Correct this manually before loading the data
Na_col_xls <- name_union[!name_union %in% names(xls_data)]
Na_col_xlsx <- name_union[!name_union %in% names(xlsx_data)]

if(length(Na_col_xls)>0){xls_data[,Na_col_xls]<-NA}
if(length(Na_col_xlsx)>0){xlsx_data[,Na_col_xlsx]<-NA}

## Join xls and xlsx file 
Master_data <- rbind(xls_data,xlsx_data)

# Checking the number of files merged properly
table(Master_data$Filename)

## Adding Retialer column into master
Master_data[,"Retailer"] <- "Panda"