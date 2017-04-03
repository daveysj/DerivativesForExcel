#include "xllFunctions.h"

using namespace XLLBasicLibrary;
#include <memory> // shared_ptr
#include <boost/algorithm/string.hpp> // to_lower

xloper* __stdcall xllInterpolate(
    double xValue,
    xl_array *xArray,
    xl_array *yArray,
    int arrayInputSize,
    char* interpolatorType,
    bool extrapolate)
{
    shared_ptr<ArrayInterpolator> interpolator;

    vector<double> xVector, yVector;
    string errorMessage;
    if (!constructVector(xArray, xVector, errorMessage))
    {
        return returnXloperOnError(errorMessage);
    }
    if (!constructVector(yArray, yVector, errorMessage))
    {
        return returnXloperOnError(errorMessage);
    }
    if (xVector.size() != yVector.size())
    {
        return returnXloperOnError("input vectors do not have the same size");
    }
    if ((int) xVector.size() < arrayInputSize)
    {
        return returnXloperOnError("input vectors are smaller than their size input");
    }
    string type = string(interpolatorType);
    boost::to_lower(type);
    if (type.compare("") == 0 || type.compare("linear") == 0)
    {
        interpolator = shared_ptr<ArrayInterpolator>(
            new  LinearArrayInterpolator(xVector, yVector, extrapolate));
    }
    else if (type.compare("cubic") == 0)
    {
        interpolator = shared_ptr<ArrayInterpolator>(
            new  CubicSplineInterpolator(xVector, yVector, extrapolate));

    }
    else
    {
        return returnXloperOnError("Type must be either linear or cubic");
    }

    cpp_xloper outputMatrix(1, 1);
    outputMatrix.SetArrayElement(0, 0, interpolator->getRate(xValue));
    return outputMatrix.ExtractXloper(false);
}

xloper* __stdcall xllBlackVolOffSurface(
	char* optionType,
	double forward,
	double strike,
	double time,
	xl_array *timeArray,
	xl_array *putDeltaArray,
	xl_array *surface,
	double convergenceThreshold,
	char* type,
	bool extrapolate)
{
	try
	{
		string errorMessage = "";
		if ((forward < 1e-14) || (strike < 1e-14) || (time < 1e-14))
		{
			return returnXloperOnError("Numeric inputs must be strictly positive");
		}
		vector<double> timeVector;
		if (!constructVector(timeArray, timeVector, errorMessage))
		{
			return returnXloperOnError(errorMessage);
		}
		double yearFraction = 1 / 365.0;
		transform(timeVector.begin(), timeVector.end(), timeVector.begin(), bind1st(multiplies<double>(), yearFraction));

		vector<double> deltaVector;
		if (!constructVector(putDeltaArray, deltaVector, errorMessage))
		{
			return returnXloperOnError(errorMessage);
		}
		// Check if the input array needs to be transposed or not
		WORD timeArray_rows, timeArray_columns;
		cpp_xloper timeXloper(timeArray);
		timeXloper.GetArraySize(timeArray_rows, timeArray_columns);
		bool transpose = false;
		if (timeArray_columns == 1 && timeArray_rows > 1)
		{
			transpose = true; 
		}
		vector<vector<double>> surfaceData;
		if (!extractDataFromSurface(surface, transpose, surfaceData, errorMessage))
		{
			return returnXloperOnError(errorMessage);
		}

		if (string(type).compare("") == 0)
		{
			type = "bilinear";
		}
		// Set extrapolate to true to ensure we can solve for vol 
		SimpleDeltaSurface deltaSurface(timeVector, deltaVector, surfaceData, true, type);

		double moneyness = (strike - forward) / forward;
		if (!deltaSurface.isInMoneynessRange(time * yearFraction, moneyness))
		{
			return returnXloperOnError("Point to interpolate is outside of the surface range and extrapolation is set to false");
		}

		double vol = deltaSurface.getVolatilityForMoneyness(time * yearFraction, moneyness);
		cpp_xloper outputObject(1, 1);
		outputObject.SetArrayElement(0, 0, vol);
		return outputObject.ExtractXloper(false);
	}
	catch (exception &e)
	{
		return returnXloperOnError(e.what());
	}
}

xloper* __stdcall prOptionValue(
    char* putOrCall,
    double forward,
    double strike,
    double standardDeviation,
    double discountFactor)
{
    string errorMessage = "Not yet implemented";
    return returnXloperOnError(errorMessage);


}

xloper* __stdcall prOptionDelta(
    char* putOrCall,
    double forward,
    double strike,
    double standardDeviation,
    double discountFactor)
{
    string errorMessage = "Not yet implemented";
    return returnXloperOnError(errorMessage);
}
