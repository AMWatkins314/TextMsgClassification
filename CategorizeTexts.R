# Load the required R libraries
# NOTE: RTextTools does not include NaiveBayes model
library(openxlsx)
library(RTextTools)



logFile <- 'CategorizeTexts_Log.txt'
dataWorkbook <- 'TextData.xlsx'
e_Matrix_File <- 'EnglishTrainingMatrix.Rd'
s_Matrix_File <- 'SpanishTrainingMatrix.Rd'
e_Model_File <- 'EnglishTrainedModel.Rd'
s_Model_File <- 'SpanishTrainedModel.Rd'


# Output any warnings/error messages to logfile
if(!file.exists(logFile)){
   file.create(logFile)
}
log <- file(logFile)
sink(log)


# Load the English and Spanish best model's term-document matrices and trained models
if(file.exists(e_Matrix_File)){
   load(file=e_Matrix_File)
} else{
   msg <- cat('English training matrix file ', e_Matrix_File, ' does not exist.')
   stop(msg)
}
if(file.exists(e_Model_File)){
   load(file=e_Model_File)
} else{
   msg <- cat('English trained model file ', e_Model_File, ' does not exist.')
   stop(msg)
}
if(file.exists(s_Matrix_File)){
   load(file=s_Matrix_File)
} else{
   msg <- cat('Spanish training matrix file ', s_Matrix_File, ' does not exist.')
   stop(msg)
}
if(file.exists(s_Model_File)){
   load(file=s_Model_File)
} else{
   msg <- cat('Spanish trained model file ', s_Model_File, ' does not exist.')
   stop(msg)
}


# Load the unclassified English and Spanish text message data from the Excel workbook into data frames
if(file.exists(dataWorkbook)){
   worksheets <- getSheetNames(dataWorkbook)

   if(is.element('EnglishData', worksheets)){
      e_Data <- readWorkbook(dataWorkbook, sheet='EnglishData', startRow=1, colNames=TRUE, rowNames=FALSE)
   } else{
      msg <- cat('File ', dataWorkbook, 'with worksheet EnglishData does not exist.')
      stop(msg) 
   }
   if(is.element('EnglishData', worksheets)){
      s_Data <- readWorkbook(dataWorkbook, sheet='SpanishData', startRow=1, colNames=TRUE, rowNames=FALSE)
   } else{
      msg <- cat('File ', dataWorkbook, 'with worksheet SpanishData does not exist.')
      stop(msg) 
   }
} else {
   msg <- cat('File ', dataWorkbook, ' does not exist.')
   stop(msg)
}


# Create document-term matrices for the unclassified data adjusting them for the training models by referencing the training matrices
# Do not remove numbers since if a text states that the user would call a specific number, that number could correspond to the Nurse Hotline
# Do remove terms that do not appear in at least (1-0.99)% of the texts
# NOTE: In R console run `trace("create_matrix",edit=T)` and on line 42 change `Acronym` to `acronym` to fix bug in RTextTools (should be fixed in future updates)
e_Matrix <- create_matrix(e_Data$TEXTCONTENT, language="english", removeNumbers=FALSE, removeSparseTerms=0.99, stemWords=TRUE, weighting=tm::weightTf, originalMatrix=e_TrainingMatrix)
s_Matrix <- create_matrix(s_Data$TEXTCONTENT, language="spanish", removeNumbers=FALSE, removeSparseTerms=0.99, stemWords=TRUE, weighting=tm::weightTf, originalMatrix=s_TrainingMatrix)


# A vector the same length as the number of texts to classify must be specified when creating the containers
# so create dummy vectors consisting of empty strings since we are using a saved model to classify 'virgin' data 
e_dummylabels <- vector(mode="character", length=nrow(e_Data)) 
s_dummylabels <- vector(mode="character", length=nrow(s_Data)) 


# Create lists of objects to feed to the saved best English and Spanish models
e_Container <- create_container(e_Matrix, factor(e_dummylabels), testSize=1:nrow(e_Data), virgin=TRUE)
s_Container <- create_container(s_Matrix, factor(s_dummylabels), testSize=1:nrow(s_Data), virgin=TRUE)


# Classify the English and Spanish text messages
# The result is a data frame with a column of predicted labels and a column of their corresponding confidence that they are correct
e_Classification <- classify_model(e_Container, e_TrainedModel) 
s_Classification <- classify_model(s_Container, s_TrainedModel) 


# Write the classification/categorization labels to the TextData.xlsx EnglishData and SpanishData sheets next to the texts 
if(file.exists(dataWorkbook)){
   workbook <- loadWorkbook(dataWorkbook)
} else{
   msg <- cat('File ', dataWorkbook, ' does not exist.')
   stop(msg)
}

# Assumes that the text to classify/categorize is in the first column of the worksheet
# So output predicted labels and confidence of predicted labels starting at column 2
writeData(workbook, sheet='EnglishData', e_Classification, startRow=1, startCol=2, colNames=TRUE, rowNames=FALSE)
writeData(workbook, sheet='SpanishData', s_Classification, startRow=1, startCol=2, colNames=TRUE, rowNames=FALSE)
saveWorkbook(workbook, dataWorkbook, overwrite=TRUE)

# Delete all data since the relevant data has been saved
rm(list=ls())
