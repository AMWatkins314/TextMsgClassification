# Load the required R libraries
# NOTE: RTextTools does not include NaiveBayes model
library(openxlsx)
library(RTextTools)



logFile <- 'CreateTextModels_Log.txt'
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


# Load the English and Spanish classified text message training data from the Excel workbook into data frames
if(file.exists(dataWorkbook)){
   worksheets <- getSheetNames(dataWorkbook)

   if(is.element('EnglishTrainingData', worksheets)){
      e_TrainingData <- readWorkbook(dataWorkbook, sheet='EnglishTrainingData', startRow=1, colNames=TRUE, rowNames=FALSE)
   } else{
      msg <- cat('File ', dataWorkbook, 'with worksheet EnglishTrainingData does not exist.')
      stop(msg)
   }
   if(is.element('SpanishTrainingData', worksheets)){
      s_TrainingData <- readWorkbook(dataWorkbook, sheet='SpanishTrainingData', startRow=1, colNames=TRUE, rowNames=FALSE)
   } else{
      msg <- cat('File ', dataWorkbook, 'with worksheet SpanishTrainingData does not exist.')
      stop(msg)
   }
} else{
   msg <- cat('File ', dataWorkbook, ' does not exist.')
   stop(msg)
}


# Create document-term matrices using the classified training data
# Do not remove numbers since if a text states that the user would call a specific number, that number could correspond to the Nurse Hotline
# Do remove terms that do not appear in at least (1-0.99)% of the texts
# NOTE: In R console run `trace("create_matrix",edit=T)` and on line 42 change `Acronym` to `acronym` to fix bug in RTextTools (should be fixed in future updates)
e_TrainingMatrix <- create_matrix(e_TrainingData$TEXTCONTENT, language="english", removeNumbers=FALSE, removeSparseTerms=0.99, stemWords=TRUE, weighting=tm::weightTf)
s_TrainingMatrix <- create_matrix(s_TrainingData$TEXTCONTENT, language="spanish", removeNumbers=FALSE, removeSparseTerms=0.99, stemWords=TRUE, weighting=tm::weightTf)


# Partition the matrices into containers
# Containers are lists of objects to feed to the machine learning algorithms for training, classifying, and analyzing text
# Set training data to be all of data since k-fold validation will take care of partitioning data into training and test sets
# virgin=FALSE since we are evaluating and not classifying virgin texts yet
# If editing code to use create_analytics then the text category factors must be converted to numbers using as.numeric 
e_TrainingContainer <- create_container(e_TrainingMatrix, factor(e_TrainingData$TEXTCATEGORY), trainSize=1:nrow(e_TrainingData), virgin=FALSE)
s_TrainingContainer <- create_container(s_TrainingMatrix, factor(s_TrainingData$TEXTCATEGORY), trainSize=1:nrow(s_TrainingData), virgin=FALSE)


# Set the number of rounds and number of folds for k-cross validation
cvRounds <- 3
cvFolds <- 4 

# Perform cross-validation for English models to get average Accuracy measurement
e_AvgAcc <- data.frame("SVM"=0, "GLMNET"=0, "MAXENT"=0, "SLDA"=0, "BAGGING"=0, "BOOSTING"=0, "RF"=0, "NNET"=0, "TREE"=0) 

for (i in 1:cvRounds){
   e_AvgAcc[1,] <- lapply(seq_along(e_AvgAcc),
                          function(x){
                             y <- cross_validate(e_TrainingContainer, cvFolds, algorithm=names(e_AvgAcc[x]));
                             e_AvgAcc[x] <- as.numeric(e_AvgAcc[x]) + ((y$meanAccuracy)/cvRounds)})
}
e_Idx <- which.max(e_AvgAcc)

# The most accurate model for classifying English texts
e_BestModelName <- names(e_AvgAcc[e_Idx])

# Perform cross-validation for Spanish models
# NOTE: There is not enough data for models GLMNET or SLDA 
s_AvgAcc <- data.frame("SVM"=0, "MAXENT"=0, "BAGGING"=0, "BOOSTING"=0, "RF"=0, "NNET"=0, "TREE"=0) 

for (i in 1:cvRounds){
   s_AvgAcc[1,] <- lapply(seq_along(s_AvgAcc),
                          function(x){
                             y <- cross_validate(s_TrainingContainer, cvFolds, algorithm=names(s_AvgAcc[x]));
                             s_AvgAcc[x] <- as.numeric(s_AvgAcc[x]) + ((y$meanAccuracy)/cvRounds)})
}
s_Idx <- which.max(s_AvgAcc)

# The most accurate model for classifying Spanish texts
s_BestModelName <- names(s_AvgAcc[s_Idx])


# Train the best English and Spanish models
e_TrainedModel <- train_model(e_TrainingContainer, e_BestModelName) 
s_TrainedModel <- train_model(s_TrainingContainer, s_BestModelName) 


# Write the accuracy data to the TextData Excel workbook
if(file.exists(dataWorkbook)){
   workbook <- loadWorkbook(dataWorkbook)
} else{
   msg <- cat('File ', dataWorkbook, ' does not exist.')
   stop(msg)
}
worksheets <- names(workbook)

# Create model accuracy worksheets if they do not already exist
if(!is.element('EnglishModelsAccuracy', worksheets)){
   addWorksheet(workbook, 'EnglishModelsAccuracy')
}
if(!is.element('SpanishModelsAccuracy', worksheets)){
   addWorksheet(workbook, 'SpanishModelsAccuracy')
}

writeData(workbook, sheet='EnglishModelsAccuracy', e_AvgAcc, startRow=1, startCol=1, colNames=TRUE, rowNames=FALSE)
writeData(workbook, sheet='SpanishModelsAccuracy', s_AvgAcc, startRow=1, startCol=1, colNames=TRUE, rowNames=FALSE)
saveWorkbook(workbook, dataWorkbook, overwrite=TRUE)


# Save the document-term matrices and best trained English and Spanish models so they can be used by CategorizeTexts.R script for classifying texts 
save(e_TrainingMatrix, file=e_Matrix_File)
save(s_TrainingMatrix, file=s_Matrix_File)
save(e_TrainedModel, file=e_Model_File)
save(s_TrainedModel, file=s_Model_File)


# Delete all data since the relevant data has been saved 
rm(list=ls())
