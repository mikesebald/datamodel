setwd("E://R//datamodel")
library(xlsx)
library(plyr)

# open Excel file
fileInput <- read.xlsx("Data Model.xlsx", sheetName = "Entities", startRow = 2, colIndex = c(1:14), encoding = "UTF-8")

# Send output to file
fileOutput <- file("entities.config", encoding = "UTF-8")
sink(fileOutput)

#
vecColnames <- c("inScope", "attributeName", "description", "permittedValues", "providedBySource",
                 "entityName", "fieldName", "fieldCaption", "fieldType", "fieldLength", "isNullable", 
                 "enumerationName", "compareOptions", "comment")

colnames(fileInput) <- vecColnames

# replace yes/no in column "is nullable" with true/false
fileInput$isNullable <- sapply(fileInput$isNullable, function(x) gsub("yes", "true", x, ignore.case = TRUE))
fileInput$isNullable <- sapply(fileInput$isNullable, function(x) gsub("no", "false", x, ignore.case = TRUE))

#do not sort by entity
#fileInput <- arrange(fileInput, entityName)

cat("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n")
cat("<config>\n")
cat("\t<entity name=\"record\">\n")
    
    
lastDmEntity <- ""
currentDmRow <- 0
for (currentDmEntity in fileInput[, "entityName"]) {
    if (!is.na(currentDmEntity)) {
        currentDmRow <- currentDmRow + 1
        if (currentDmEntity != lastDmEntity) {
            if (currentDmRow > 1)
                cat("\t\t</entity>\n")
             
            lastDmEntity <- currentDmEntity
            cat(paste("\t\t<entity name=\"", lastDmEntity, "\" caption=\"", lastDmEntity, "\">", "\n", sep = ""))
        }

        fieldName <- fileInput[currentDmRow, "fieldName"]
        fieldCaption <- fileInput[currentDmRow, "fieldCaption"]
        fieldType <- tolower(fileInput[currentDmRow, "fieldType"])
        fieldLength <- fileInput[currentDmRow, "fieldLength"]
        isNullable <- fileInput[currentDmRow, "isNullable"]
        enumerationName <- fileInput[currentDmRow, "enumerationName"]
        
        if (is.na(fileInput[currentDmRow, "enumerationName"]))
             enumerationName <- ""
        else
             enumerationName <- paste("enum=\"", fileInput[currentDmRow, "enumerationName"], "\" ", sep = "")

        if (is.na(fileInput[currentDmRow, "compareOptions"]))
             compareOptions <- ""
        else
             compareOptions <- paste("compareoptions=\"", fileInput[currentDmRow, "compareOptions"], "\" ", sep = "")

        if (fieldType == "enum") {
            outputLine <- paste("\t\t\t<field name=\"", fieldName, "\" ", 
                                "caption=\"", fieldCaption, "\" ", 
                                "type=\"enum\" ", 
                                enumerationName,
                                "nullable=\"", isNullable, "\" />", "\n", sep = "")
        }
        else if (fieldType == "date") {
             outputLine <- paste("\t\t\t<field name=\"", fieldName, "\" ", 
                                 "caption=\"", fieldCaption, "\" ", 
                                 "type=\"date\" ",
                                 "nullable=\"", isNullable, "\" />", "\n", sep = "")
        }
        else if (fieldType == "datetime") {
             outputLine <- paste("\t\t\t<field name=\"", fieldName, "\" ", 
                                 "caption=\"", fieldCaption, "\" ", 
                                 "type=\"datetime\" ",
                                 "nullable=\"", isNullable, "\" />", "\n", sep = "")
        }
        else if (fieldType == "integer") {
            outputLine <- paste("\t\t\t<field name=\"", fieldName, "\" ", 
                                "caption=\"", fieldCaption, "\" ", 
                                "type=\"integer\" ",
                                "nullable=\"", isNullable, "\" />", "\n", sep = "")
        }
        else if (fieldType == "boolean") {
            outputLine <- paste("\t\t\t<field name=\"", fieldName, "\" ", 
                                "caption=\"", fieldCaption, "\" ", 
                                "type=\"boolean\" ",
                                "nullable=\"", isNullable, "\" />", "\n", sep = "")
        }
        else if (fieldType == "string") {
            outputLine <- paste("\t\t\t<field name=\"", fieldName, "\" ", 
                                "caption=\"", fieldCaption, "\" ", 
                                enumerationName,
                                "type=\"string\" ",
                                "nullable=\"", isNullable, "\" ",
                                compareOptions,
                                "size=\"", fieldLength, "\" />", "\n", sep = "")
        }
        else {
            outputLine <- paste("INVALID SOURCE DATA !!!")
        }
        cat(outputLine)
    }
}

#last elements are processed
#cat("\t\t\t</fieldgroup>\n")
cat("\t\t</entity>\n")
cat("\t</entity>\n")
cat("</config>\n")

# Write output file
sink(file = NULL)
