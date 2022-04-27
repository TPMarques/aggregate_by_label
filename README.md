# aggregate_by_label
The aggregate_by_label is an excel add-in that includes functions to calculate sum, the sum of squares, mean, variance, and count observations of data aggregating according to a specific label. The repository consists of .xla and .xlam Excel add-in files, .bas basic source code file, and an example macro-enable .xlsm excel sheet.

The functions available are:

1- SUMByLabel (calculates the sum of values for a specific label);
2- SUMSQByLabel (calculates the sum of squares);
3- VARByLabel (calculates the sample variance);
4- MEANByLabel (calculates the mean);
5- COUNTByLabel (calculate the number of observations).

The first four functions have the same inputs: Categories, Label and SumRange. Categories must be a range with label data for a given observation, the Label is the source label to be matched, and SumRange is the list of values used to calculate the desired statistics. COUNTByLabel doesn't have SumRange as input.

