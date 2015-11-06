setwd("E://R//datamodel")
library(xlsx)
library(plyr)

# open Excel file
fileInput <- read.xlsx("Data Model.xlsx", sheetName = "Entities", startRow = 2, colIndex = c(1:13), encoding = "UTF-8")

# Send output to file
fileOutput <- file("entities.config", encoding = "UTF-8")
sink(fileOutput)

#
vecColnames <- c("inScope", "attributeName", "description", "permittedValues", "providedBySource",
                 "entityName", "fieldName", "fieldCaption", "fieldType", "fieldLength", "isNullable", 
                 "enumerationName", "comment")

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
            if (currentDmRow > 1) {
                #cat("\t\t\t</fieldgroup>\n")
                cat("\t\t</entity>\n")
            }
            lastDmEntity <- currentDmEntity
            cat(paste("\t\t<entity name=\"", lastDmEntity, "\" caption=\"", lastDmEntity, "\">", "\n", sep = ""))
            #cat("\t\t\t<fieldgroup>\n")
        }
        if (tolower(fileInput[currentDmRow, "fieldType"]) == "enum") {
            outputLine <- paste("\t\t\t<field name=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                "type=\"enum\" ", 
                                "enum=\"", fileInput[currentDmRow, "enumerationName"], "\" ",
                                "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" />", "\n", sep = "")
        }
        else if (tolower(fileInput[currentDmRow, "fieldType"]) == "date") {
             outputLine <- paste("\t\t\t<field name=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                 "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                 "type=\"date\" ",
                                 "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" />", "\n", sep = "")
        }
        else if (tolower(fileInput[currentDmRow, "fieldType"]) == "datetime") {
             outputLine <- paste("\t\t\t<field name=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                 "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                 "type=\"datetime\" ",
                                 "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" />", "\n", sep = "")
        }
        else if (tolower(fileInput[currentDmRow, "fieldType"]) == "integer") {
            outputLine <- paste("\t\t\t<field name=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                "type=\"integer\" ",
                                "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" />", "\n", sep = "")
        }
        else if (tolower(fileInput[currentDmRow, "fieldType"]) == "boolean") {
            outputLine <- paste("\t\t\t<field name=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                "type=\"boolean\" ",
                                "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" />", "\n", sep = "")
        }
        else if ((tolower(fileInput[currentDmRow, "fieldType"]) == "string")
             && (!is.na(tolower(fileInput[currentDmRow, "enumerationName"])))) {
            outputLine <- paste("\t\t\t<field name=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                "enum=\"", fileInput[currentDmRow, "enumerationName"], "\" ",
                                "type=\"string\" ",
                                "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" ",
                                "size=\"", fileInput[currentDmRow, "fieldLength"], "\" />", "\n", sep = "")
        }
        else if ((tolower(fileInput[currentDmRow, "fieldType"]) == "string")
             && (is.na(tolower(fileInput[currentDmRow, "enumerationName"])))) {
            outputLine <- paste("\t\t\t<field name=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                           "caption=\"", fileInput[currentDmRow, "fieldCaption"], "\" ", 
                                              "type=\"", fileInput[currentDmRow, "fieldType"], "\" ", 
                                          "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" ", 
                                              "size=\"", fileInput[currentDmRow, "fieldLength"], "\" />", "\n", sep = "")
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
