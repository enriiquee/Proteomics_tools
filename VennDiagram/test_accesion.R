######################################################################
#TEST_ACCESION: 
#This script creates a xml file with merged data and Venn Diagram plot. 
######################################################################

#Load packages
if (!require('VennDiagram')) install.packages('VennDiagram')
library('VennDiagram')
if (!require('readxl')) install.packages('readxl')
library('readxl')
if (!require('WriteXLS')) install.packages('WriteXLS')
library('WriteXLS')

#Suppress warnings globally
options(warn = -1)
print("Ejecutando... Puede tardar un poco. ")


#Read files from directory. 
files_glob <- (Sys.glob("*.xlsx")) 


#Conditionals
if (length(files_glob) == 1){
  print("There is only 1 xlsx file")
  
} else if (length(files_glob) == 2){
  #Load excel file
  
  df1 <- read_excel(file.path(getwd(), files_glob[1]))
  df2 <- read_excel(file.path(getwd(), files_glob[2]))
  
  #Combine data frame using reduce function
  df_final <- Reduce(function(x, y) merge (x, y, by = c("Name", "Accession"), all = TRUE), list(df1, df2))
  
  #Export data frame to excel file.
  WriteXLS(df_final, ExcelFileName = "export2.xls", SheetNames = NULL, perl = "perl",
           verbose = FALSE, Encoding = c("UTF-8", "latin1", "cp1252"),
           row.names = FALSE, col.names = TRUE,
           AdjWidth = FALSE, AutoFilter = FALSE, BoldHeaderRow = FALSE,
           na = "",
           FreezeRow = 0, FreezeCol = 0,
           envir = parent.frame())
  

  #Dissable .log files
  futile.logger::flog.threshold(futile.logger::ERROR, name = "VennDiagramLogger")
  
  #Function ven.diagram and grid. 
  pdf("mygraph2.pdf")
  
  #Create a Venn diagram
  v <- venn.diagram(list(Muestra1=df1$Accession,Muestra2=df2$Accession),
                    fill = c("red", "blue"),
                    cat.cex = 1.5, cex=1.5,
                    filename=NULL)
  # have a look at the default plot
  grid.newpage()
  grid.draw(v)
  garbage <- dev.off()
  
} else if (length(files_glob) == 3){
  #Load excel file
  df1 <- read_excel(file.path(getwd(), files_glob[1]))
  df2 <- read_excel(file.path(getwd(), files_glob[2]))
  df3 <- read_excel(file.path(getwd(), files_glob[3]))

  #Combine data frame using reduce function
  df_final <- Reduce(function(x, y) merge (x, y, by = c("Name", "Accession"), all = TRUE), list(df1, df2, df3))
  
  #Export data frame to table.
  WriteXLS(df_final, ExcelFileName = "export3.xls", SheetNames = NULL, perl = "perl",
           verbose = FALSE, Encoding = c("UTF-8", "latin1", "cp1252"),
           row.names = FALSE, col.names = TRUE,
           AdjWidth = FALSE, AutoFilter = FALSE, BoldHeaderRow = FALSE,
           na = "",
           FreezeRow = 0, FreezeCol = 0,
           envir = parent.frame())
  
  
  #Create a Venn Diagram: 
  #Dissable .log files
  futile.logger::flog.threshold(futile.logger::ERROR, name = "VennDiagramLogger")
  
  #Function ven.diagram and grid. 
  pdf("mygraph3.pdf")
  v <- venn.diagram(list(Muestra1=df1$Accession,Muestra2=df2$Accession, Muestra3=df3$Accession),
                    fill = c("red", "blue", "green"),
                    cat.cex = 1.5, cex=1.5,
                    filename=NULL)
  # have a look at the default plot
  grid.newpage()
  grid.draw(v)
  garbage <- dev.off()
  
} else if (length(files_glob) == 4){
  df1 <- read_excel(file.path(getwd(), files_glob[1]))
  df2 <- read_excel(file.path(getwd(), files_glob[2]))
  df3 <- read_excel(file.path(getwd(), files_glob[3]))
  df4 <- read_excel(file.path(getwd(), files_glob[4]))
  
  
  #Combine data frame using reduce function
  df_final <- Reduce(function(x, y) merge (x, y, by = c("Name", "Accession"), all = TRUE), list(df1, df2, df3, df4))
  
  #Export data frame to table.
  WriteXLS(df_final, ExcelFileName = "export4.xls", SheetNames = NULL, perl = "perl",
           verbose = FALSE, Encoding = c("UTF-8", "latin1", "cp1252"),
           row.names = FALSE, col.names = TRUE,
           AdjWidth = FALSE, AutoFilter = FALSE, BoldHeaderRow = FALSE,
           na = "",
           FreezeRow = 0, FreezeCol = 0,
           envir = parent.frame())
  
  
  #Create a Venn Diagram: 
  #Dissable .log files
  futile.logger::flog.threshold(futile.logger::ERROR, name = "VennDiagramLogger")
  
  #Function ven.diagram and grid. 
  pdf("mygraph4.pdf")
  v <- venn.diagram(list(Muestra1=df1$Accession,Muestra2=df2$Accession,Muestra3=df3$Accession,Muestra4=df4$Accession),
                    fill = c("red", "blue", "green", "yellow"),
                    cat.cex = 1.5, cex=1.5,
                    filename=NULL)
  # have a look at the default plot
  grid.newpage()
  grid.draw(v)
  garbage <- dev.off()
  
}else{
  df1 <- read_excel(file.path(getwd(), files_glob[1]))
  df2 <- read_excel(file.path(getwd(), files_glob[2]))
  df3 <- read_excel(file.path(getwd(), files_glob[3]))
  df4 <- read_excel(file.path(getwd(), files_glob[4]))
  df5 <- read_excel(file.path(getwd(), files_glob[5]))
  
  #Combine data frame using reduce function
  df_final <- Reduce(function(x, y) merge (x, y, by = c("Name", "Accession"), all = TRUE), list(df1, df2, df3, df4, df5))
  
  #Export data frame to table.
  WriteXLS(df_final, ExcelFileName = "export5.xls", SheetNames = NULL, perl = "perl",
           verbose = FALSE, Encoding = c("UTF-8", "latin1", "cp1252"),
           row.names = FALSE, col.names = TRUE,
           AdjWidth = FALSE, AutoFilter = FALSE, BoldHeaderRow = FALSE,
           na = "",
           FreezeRow = 0, FreezeCol = 0,
           envir = parent.frame())
  
  #Create a Venn Diagram: 
  #Dissable .log files
  futile.logger::flog.threshold(futile.logger::ERROR, name = "VennDiagramLogger")
  
  #Function ven.diagram and grid. 
  pdf("mygraph5.pdf")
  v <- venn.diagram(list(Muestra1=df1$Accession,Muestra2=df2$Accession,Muestra3=df3$Accession,Muestra4=df4$Accession, Muestra5=df5$Accession),
                    fill = c("red", "blue", "green", "yellow", "black"),
                    cat.cex = 1.5, cex=1.5,
                    filename=NULL)
  # have a look at the default plot
  grid.newpage()
  grid.draw(v)
  garbage <- dev.off()
  

}



