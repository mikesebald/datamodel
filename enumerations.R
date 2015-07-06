setwd("E://R//datamodel")
library(xlsx)
library(plyr)

# open Excel file
fileInput <- read.xlsx("Data Model.xlsx", sheetName = "Enumerations", startRow = 1, colIndex = c(1:4), encoding = "UTF-8")

# Send output to file
fileOutput <- file("enumerations_fragment.config", encoding = "UTF-8")
sink(fileOutput)

vecColnames <- c("type", "code", "description", "reversedescription")

colnames(fileInput) <- vecColnames

fileInput <- arrange(fileInput, type)

cat("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n")
cat("<config>\n")
cat("\t<enums>\n")

lastDmEnum <- ""
currentDmRow <- 0
enumCounter <- 0

for (currentDmEnum in fileInput[, "type"]) {
    if (!is.na(currentDmEnum)) {
        currentDmRow <- currentDmRow + 1
        enumCounter <- enumCounter + 1
        if (currentDmEnum != lastDmEnum) {
            if (currentDmRow > 1) {
                cat("\t\t</enum>\n")
            }
            lastDmEnum <- currentDmEnum
            cat(paste("\t\t<enum type=\"", lastDmEnum, "\">", "\n", sep = ""))
            enumCounter = 1
        }
        if (is.na(fileInput[currentDmRow, "reversedescription"])) {
            outputLine <- paste("\t\t\t<value id=\"", enumCounter, "\" ", 
                                           "code=\"", fileInput[currentDmRow, "code"], "\" ", 
                                    "description=\"", fileInput[currentDmRow, "description"], "\">", "\n", sep = "")
        }
        else {
            outputLine <- paste("\t\t\t<value id=\"", enumCounter, "\" ", 
                                           "code=\"", fileInput[currentDmRow, "code"], "\" ", 
                                    "description=\"", fileInput[currentDmRow, "description"], "\" ", 
                             "reversedescription=\"", fileInput[currentDmRow, "reversedescription"], "\">", "\n", sep = "")
        }
        cat(outputLine)
    }
}

#last elements are processed
cat("\t\t</enum>\n")
cat("\t</enums>\n")
cat("</config>\n")

# Write output file
sink(file = NULL)
