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
	char* type,
	bool extrapolate)
{
    string errorMessage = "Not yet implemented";
    return returnXloperOnError(errorMessage);
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
