#include "xllFunctions.h"

using namespace XLLBasicLibrary;
#include <memory> // shared_ptr
#include <boost/algorithm/string.hpp> // to_lower

xloper* __stdcall Interpolate(
    double xValue,
    xl_array *xArray,
    xl_array *yArray,
    int arrayInputSize,
    char* interpolatorType,
    bool extrapolate)
{
	try
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
			return returnXloperOnError("X and Y input arrays have inconsistent dimension");
		}
		if ((int) xVector.size() < arrayInputSize)
		{
			return returnXloperOnError("\"Size\" input is greater than the lenght of the X and Y arrays");
		}
		xVector.resize(arrayInputSize);
		yVector.resize(arrayInputSize);
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
			return returnXloperOnError("\"Type\" must be either \"Linear\" or \"Cubic\"");
		}
		if (!interpolator->isOk())
		{
			return returnXloperOnError(interpolator->getErrorMessage());
		}
		return returnXloper(interpolator->getRate(xValue));
	}
	catch (exception &e)
	{
		return returnXloperOnError(e.what());
	}
}

xloper* __stdcall BlackVolOffSurface(
	char* optionType,
	double forward,
	double strike,
	double day,
	xl_array *dayArray,
	xl_array *putDeltaArray,
	xl_array *surface,
	double convergenceThreshold,
	char* type,
	bool extrapolate)
{
	try
	{
		string errorMessage = "";
		if ((forward < 1e-14) || (strike < 1e-14) || (day < 1e-14))
		{
			return returnXloperOnError("Numeric inputs must be strictly positive");
		}
		PutCall putCallType;
		if (!getPutCall(optionType, putCallType, errorMessage))
		{
			return returnXloperOnError(errorMessage);
		}

		vector<double> timeVector;
		if (!constructVector(dayArray, timeVector, errorMessage))
		{
			return returnXloperOnError(errorMessage);
		}
		// Hard coded explicit assumption that the time input uses days but everything in
		// the code uses year fractions. The following lines convert days into year fractions
		double yearFraction = 1 / 365.0;
		transform(
			timeVector.begin(),
			timeVector.end(),
			timeVector.begin(),
			bind1st(multiplies<double>(), yearFraction));
		// The SimpleDeltaSurface class assumes the surface data is input with the time
		// in the x-dimenstion and delta in the y-dimension. If the inputs do not conform
		// then the surface needs to be transposed before it can be used
		WORD timeArray_rows, timeArray_columns;
		cpp_xloper timeXloper(dayArray);
		timeXloper.GetArraySize(timeArray_rows, timeArray_columns);
		bool transpose = false;
		if (timeArray_columns == 1 && timeArray_rows > 1)
		{
			transpose = true;
		}

		vector<double> deltaVector;
		if (!constructVector(putDeltaArray, deltaVector, errorMessage))
		{
			return returnXloperOnError(errorMessage);
		}

		vector<vector<double>> surfaceData;
		if (!extractDataFromSurface(surface, transpose, surfaceData, errorMessage))
		{
			return returnXloperOnError(errorMessage);
		}

		if (std::string(type).compare("") == 0)
		{
			type = "bilinear";
		}
		// Set extrapolate to true to ensure we can solve for vol 
		SimpleDeltaSurface deltaSurface(timeVector, deltaVector, surfaceData, true, type);

		double moneyness = (strike - forward) / forward;
		if (!deltaSurface.isInMoneynessRange(day * yearFraction, moneyness))
		{
			return returnXloperOnError("Point to interpolate is outside of the surface range and extrapolation is set to false");
		}

		double vol = deltaSurface.getVolatilityForMoneyness(day * yearFraction, moneyness);
		return returnXloper(vol);
	}
	catch (exception &e)
	{
		return returnXloperOnError(e.what());
	}
}

xloper* __stdcall Black(
    char* putOrCall,
    double forward,
    double strike,
	double dtm,
    double standardDeviation,
    double discountFactor)
{
	try
	{
		if ((forward < 1e-14) || (strike < 1e-14) || (standardDeviation < 1e-14) || (discountFactor < 1e-14))
		{
			return returnXloperOnError("All numeric inputs to this function must be strictly positive");
		}
		PutCall putCallType;
		string errorMessage;
		if (!getPutCall(putOrCall, putCallType, errorMessage))
		{
			return returnXloperOnError(errorMessage);
		}

		shared_ptr<Black76Option> option;
		string putOrCallString = string(putOrCall);
		boost::to_lower(putOrCallString);
		if (putCallType == CALL)
		{
			option = shared_ptr<Black76Call>(new
				Black76Call(forward, strike, standardDeviation, discountFactor));
		}
		else // (putCallType == PUT)
		{
			option = shared_ptr<Black76Put>(new
				Black76Put(forward, strike, standardDeviation, discountFactor));
		}

		double optionPremium;
		if (standardDeviation < 1e-14)
		{
			optionPremium = option->getPremiumAfterMaturity(forward, discountFactor);
		}
		else
		{
			optionPremium = option->getPremium();
		}
		return returnXloper(optionPremium);
	}
	catch (exception &e)
	{
		return returnXloperOnError(e.what());
	}

}

xloper* __stdcall BlackDelta(
    char* putOrCall,
    double forward,
    double strike,
	double dtm,
    double standardDeviation,
    double discountFactor)
{
	try
	{
		if ((forward < 1e-14) || (strike < 1e-14) || (standardDeviation < 1e-14) || (discountFactor < 1e-14))
		{
			return returnXloperOnError("All numeric inputs to this function must be strictly positive");
		}

		shared_ptr<Black76Option> option;
		string putOrCallString = string(putOrCall);
		boost::to_lower(putOrCallString);
		if (putOrCallString.compare("c") == 0 || putOrCallString.compare("call") == 0)
		{
			option = shared_ptr<Black76Call>(new
				Black76Call(forward, strike, standardDeviation, discountFactor));
		}
		else if (putOrCallString.compare("p") == 0 || putOrCallString.compare("put") == 0)
		{
			option = shared_ptr<Black76Put>(new
				Black76Put(forward, strike, standardDeviation, discountFactor));
		}
		else
		{
			return returnXloperOnError("Option type must be either (P)ut or (C)all");
		}
		double optionDelta;
		if (standardDeviation < 1e-14)
		{
			return returnXloperOnError("Option past maturity");
		}
		else
		{
			optionDelta = option->getDelta();
		}
		return returnXloper(optionDelta);
	}
	catch (exception &e)
	{
		return returnXloperOnError(e.what());
	}
}
