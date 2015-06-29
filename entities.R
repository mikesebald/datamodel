setwd("E://R//datamodel")
library(xlsx)

# open Excel file
fileInput <- read.xlsx("Data Model 2.xlsx", sheetName = "Entities", startRow = 2, colIndex = c(1:11), encoding = "UTF-8")

# Send output to file
fileOutput <- "entities_fragment.config"
sink(fileOutput)

#
vecColnames <- c("inScope", "attributeName", "description", "permittedValues", 
                 "entityName", "fieldName", "fieldCaption", "fieldType", "fieldLength", "isNullable", "enumerationName")

colnames(fileInput) <- vecColnames

# replace yes/no in column "is nullable" with true/false
fileInput$isNullable <- sapply(fileInput$isNullable, function(x) gsub("yes", "true", x, ignore.case = TRUE))
fileInput$isNullable <- sapply(fileInput$isNullable, function(x) gsub("no", "false", x, ignore.case = TRUE))

lastDmEntity <- ""
currentDmRow <- 0
for (currentDmEntity in fileInput[, "entityName"]) {
    if (!is.na(currentDmEntity)) {
        currentDmRow <- currentDmRow + 1
        if (currentDmEntity != lastDmEntity) {
            if (currentDmRow > 1) {
                outputLine <- "\t\t\t</fieldgroup>\n"
                cat(outputLine)
                outputLine <- "\t\t</entity>\n"
                cat(outputLine)
            }
            lastDmEntity <- currentDmEntity
            outputLine <- paste("\t\t<entityName=\"", lastDmEntity, "\">", "\n", sep = "")
            cat(outputLine)
            outputLine <- "\t\t\t<fieldgroup>\n"
            cat(outputLine)
        }
        if (tolower(fileInput[currentDmRow, "fieldType"]) == "enum") {
            outputLine <- paste("\t\t\t\t<fieldName=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                            "caption=\"", fileInput[currentDmRow, "field caption"], "\" ", 
                                               "type=\"enum\" ", 
                                               "enum=\"", fileInput[currentDmRow, "enumerationName"], "\">", "\n", sep = "")
        }
        else if (tolower(fileInput[currentDmRow, "fieldType"]) == "bool") {
            outputLine <- paste("\t\t\t\t<fieldName=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                            "caption=\"", fileInput[currentDmRow, "field caption"], "\" ", 
                                               "type=\"string\" ", 
                                           "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" ", 
                                               "size=\"", "1", "\">", "\n", sep = "")
        }
        else {
            outputLine <- paste("\t\t\t\t<fieldName=\"", fileInput[currentDmRow, "fieldName"], "\" ", 
                                            "caption=\"", fileInput[currentDmRow, "field caption"], "\" ", 
                                               "type=\"", fileInput[currentDmRow, "fieldType"], "\" ", 
                                           "nullable=\"", fileInput[currentDmRow, "isNullable"], "\" ", 
                                               "size=\"", fileInput[currentDmRow, "fieldLength"], "\">", "\n", sep = "")
        }
        cat(outputLine)
    }
}

#last elements are processed
outputLine <- "\t\t\t</fieldgroup>\n"
cat(outputLine)
outputLine <- "\t\t</entity>\n"
cat(outputLine)

# Write output file
sink()
