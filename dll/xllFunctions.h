#ifndef derivativeXLLInterface_INCLUDED
#define derivativeXLLInterface_INCLUDED

#include "xllFunctionSupport.h"

#include "..\Maths\maths.h"
#include "..\Maths\TwoDimensionalInterpolation.h"
#include "..\Derivatives\VolatilitySurfaceDelta.h"

/*======================================================================================
Excel Pricing functions

Converts a string (normally containing an error message) to an *xloper so it can be
returned to excel
=======================================================================================*/

xloper* __stdcall Interpolate(
    double xValue,
    xl_array *xArray,
    xl_array *yArray,
    int arrayInputSize,
    char* interpolatorType = "Linear", // Linear or Cubic                                
    bool extrapolate = false);

xloper* __stdcall BlackVolOffSurface(
	char* optionType,
	double forward,
	double strike,
	double day,
	xl_array *dayArray,
	xl_array *putDeltaArray,
	xl_array *surface,
	double convergenceThreshold, // hard coded. only here to keep the function signature unchanged
	char* type,
	bool extrapolate);

xloper* __stdcall Black(
    char* putOrCall, 
    double forward, 
    double strike, 
	double dtm,
    double standardDeviation, 
    double discountFactor);

xloper* __stdcall BlackDelta(
    char* putOrCall,
    double forward,
    double strike,
    double dtm,
	double standardDeviation,
    double discountFactor);



#endif
