setwd("E://R//datamodel")
library(xlsx)
library(plyr)

# open Excel file
fileInput <- read.xlsx("Data Model.xlsx", sheetName = "Entities", startRow = 2, colIndex = c(1:11), encoding = "UTF-8")

# Send output to file
fileOutput <- file("entities_fragment.config", encoding = "UTF-8")
sink(fileOutput)

#
vecColnames <- c("inScope", "attributeName", "description", "permittedValues", 
                 "entityName", "fieldName", "fieldCaption", "fieldType", "fieldLength", "isNullable", "enumerationName")

colnames(fileInput) <- vecColnames

# replace yes/no in column "is nullable" with true/false
fileInput$isNullable <- sapply(fileInput$isNullable, function(x) gsub("yes", "true", x, ignore.case = TRUE))
fileInput$isNullable <- sapply(fileInput$isNullable, function(x) gsub("no", "false", x, ignore.case = TRUE))

fileInput <- arrange(fileInput, entityName)

cat("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n")
cat("<config>\n")
cat("\t<entity name=\"record\">\n")
    
    
lastDmEntity <- ""
currentDmRow <- 0
for (currentDmEntity in fileInput[, "entityName"]) {
    if (!is.na(currentDmEntity)) {
        currentDmRow <- currentDmRow + 1
        if (currentDmEntity != lastDmEntity) {
            if (currentDmRow > 1) {
                cat("\t\t\t</fieldgroup>\n")
                cat("\t\t</entity>\n")
            }
            lastDmEntity <- currentDmEntity
            cat(paste("\t\t<entityName=\"", lastDmEntity, "\">", "\n", sep = ""))
            cat("\t\t\t<fieldgroup>\n")
        }
        if (tolower(fileInput[currentDmRow, "fieldType"]) == "enum") {
            outputLine <- paste("\t\t\t\t<fieldName=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                           "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                              "type=\"enum\" ", 
                                              "enum=\"", fileInput[currentDmRow, "enumerationName"], "\">", "\n", sep = "")
        }
        else if (tolower(fileInput[currentDmRow, "fieldType"]) == "boolean") {
            outputLine <- paste("\t\t\t\t<fieldName=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                           "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                              "type=\"boolean\" ", 
                                          "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" ", 
                                              "size=\"", "1", "\">", "\n", sep = "")
        }
        else {
            outputLine <- paste("\t\t\t\t<fieldName=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                           "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                              "type=\"", fileInput[currentDmRow, "fieldType"], "\" ", 
                                          "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" ", 
                                              "size=\"", fileInput[currentDmRow, "fieldLength"], "\">", "\n", sep = "")
        }
        cat(outputLine)
    }
}

#last elements are processed
cat("\t\t\t</fieldgroup>\n")
cat("\t\t</entity>\n")
cat("\t</entity>\n")
cat("</config>\n")

# Write output file
sink()
sink(NULL)
