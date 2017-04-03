#ifndef XLLBASIX_VOLATILITYSURFACEDELTA_INCLUDED
#define XLLBASIX_VOLATILITYSURFACEDELTA_INCLUDED
#pragma once

#include <vector>
#include <boost\algorithm\string.hpp>
#include "..\Maths\TwoDimensionalInterpolation.h"
#include "Black76Formula.h"

using namespace std;

namespace XLLBasicLibrary
{

	/*======================================================================================
	SimpleDeltaSurface

	Note:
	- Delta is assumed to be an integer i.e. 25-delta must be input as 25 and not 0.25
	- Volatility is assumed to be a percentage to 20-vol must be input as 0.20 and not 20
	- Time is assumed to be a year fraction.

	Moneyness = (strike - forward) / forward
	=======================================================================================*/
	class SimpleDeltaSurface
	{
	public:
		SimpleDeltaSurface(vector<double> times,
			vector<double> delta,
			vector<vector<double>> volatility,
			bool extrapolate,
			string interpolationType);

		// try to change inputs so they are consistent with the class requirements
		// return false if unable to do this
		bool checkAndTransformInputs(string &reasonForFailure);

		bool isInDeltaRange(double time, double delta);
		bool isInMoneynessRange(double time, double moneyness);

		double getVolatility(double time) const;
		double getVolatilityForDelta(double time, double delta) const;
		double getVolatilityForMoneyness(double time, double moneyness) const;

	private:
		double calculateDeltaFromStrike(double forward, double strike, double time) const;

		bool extrapolate;
		vector<double> times, delta;
		vector<vector<double>> volatility;
		shared_ptr<TwoDimensionalInterpolator> interpolator;
		string interpolationType, className;
	};
}

#endif