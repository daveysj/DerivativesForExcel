#ifndef derivativeXLLInterface_INCLUDED
#define derivativeXLLInterface_INCLUDED

#include "xllFunctionSupport.h"

#include "..\Maths\maths.h"
#include "..\Maths\TwoDimensionalInterpolation.h"
#include "..\Derivatives\VolatilitySurfaceDelta.h"


xloper* __stdcall xllInterpolate(
    double xValue,
    xl_array *xArray,
    xl_array *yArray,
    int arrayInputSize,
    char* interpolatorType = "Linear", // Linear or Cubic                                
    bool extrapolate = false);

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
	bool extrapolate);

xloper* __stdcall prOptionValue(
    char* putOrCall, 
    double forward, 
    double strike, 
    double standardDeviation, 
    double discountFactor);

xloper* __stdcall prOptionDelta(
    char* putOrCall,
    double forward,
    double strike,
    double standardDeviation,
    double discountFactor);



#endif
