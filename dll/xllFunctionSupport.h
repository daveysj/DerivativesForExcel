#ifndef derivativeXLLSupportInterface_INCLUDED
#define derivativeXLLSupportInterface_INCLUDED

#include "excelIntegration\xloper.h"
#include "excelIntegration\xl_array.h"
#include "excelIntegration\cpp_xloper.h"
#include "excelIntegration\xllAddIn.h"

#include <vector>

using namespace std;

/*======================================================================================
returnXloper

Convert a vector of values into an *xloper so it can be sent to excel
=======================================================================================*/
template <typename T>
xloper* returnXloper(vector<T> outputVector)
{
    WORD output_rows = (WORD)outputVector.size();
    WORD output_columns = 1;
    cpp_xloper outputMatrix(output_rows, output_columns);
    for (size_t i = 0; i < output_rows; ++i)
    {
        double rate = outputVector[i];
        outputMatrix.SetArrayElement(i, 0, rate);
    }
    return outputMatrix.ExtractXloper(false);
}

/*======================================================================================
returnXloperOnError

Converts a string (normally containing an error message) and turns it into an *xloper
so it can be returned to excel
=======================================================================================*/
xloper* returnXloperOnError(string errorMessage);

/*======================================================================================
constructVector

Take and *xl_array and turn it into a vector and detail about the success (or not) of
this opperation
=======================================================================================*/
bool constructVector(xl_array* column, vector<double> &outputVector, string &errorMessage);

/*======================================================================================
extractDataFromSurface

Take and *xl_array and turn it into a vector or vectors, creating detail about the success 
(or not) of this opperation. 
=======================================================================================*/
bool extractDataFromSurface(
    xl_array* surfaceInput,
    vector<vector<double>> &data,
    string &errorMessage);

/*======================================================================================
extractDataFromSurface

Same as previous function but make allowance for column and row headings. The assumption
is that the row and column headings are delta / strike or time and will be used to 
build a surface interpolation object. 

This method assumes that the first row contains the row column heading and the first
column contains the row headins. The first entry i.e. data[0][0] is ignored
=======================================================================================*/
bool extractDataFromSurface(
    xl_array* surfaceInput, 
    vector<double> &columnHeadings, 
    vector<double> &rowHeadings, 
    vector<vector<double>> &data, 
    string &errorMessage);


#endif
