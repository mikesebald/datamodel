setwd("E://R//datamodel")
library(xlsx)

# open Excel file
fileInput <- read.xlsx("Data Model 2.xlsx", sheetName = "Entities", encoding = "UTF-8")

# Send output to file
fileOutput <- "entities_fragment.config"
sink(fileOutput)

#
vecCols <- c("entity", "field name", "caption", "type", "nullable", "size")


lastDmEntity <- ""
currentDmRow <- 0
for (currentDmEntity in fileInput[, 3]) {
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
            outputLine <- paste("\t\t<entity name=\"", lastDmEntity, "\">", "\n", sep = "")
            cat(outputLine)
            outputLine <- "\t\t\t<fieldgroup>\n"
            cat(outputLine)
        }
        if (tolower(fileInput[currentDmRow, 5]) == "enum") {
            outputLine <- paste("\t\t\t\t<field name=\"", fileInput[currentDmRow, 4], "\" ", 
                                            "caption=\"", fileInput[currentDmRow, 2], "\" ", 
                                               "type=\"enum\" ", 
                                               "enum=\"", fileInput[currentDmRow, 9], "\">", "\n", sep = "")
            
        }
        else if (tolower(fileInput[currentDmRow, 5]) == "bool") {
            outputLine <- paste("\t\t\t\t<field name=\"", fileInput[currentDmRow, 4], "\" ", 
                                            "caption=\"", fileInput[currentDmRow, 2], "\" ", 
                                               "type=\"string\" ", 
                                           "nullable=\"", fileInput[currentDmRow, 7], "\" ", 
                                               "size=\"", "1", "\">", "\n", sep = "")
        }
        else {
            outputLine <- paste("\t\t\t\t<field name=\"", fileInput[currentDmRow, 4], "\" ", 
                                            "caption=\"", fileInput[currentDmRow, 2], "\" ", 
                                               "type=\"", fileInput[currentDmRow, 5], "\" ", 
                                           "nullable=\"", fileInput[currentDmRow, 7], "\" ", 
                                               "size=\"", fileInput[currentDmRow, 6], "\">", "\n", sep = "")
        }
            
        cat(outputLine)
    }
    else {
        outputLine <- "\t\t\t</fieldgroup>\n"
        cat(outputLine)
        outputLine <- "\t\t</entity>\n"
        cat(outputLine)
        break
    }
}

# Write output file
sink()
