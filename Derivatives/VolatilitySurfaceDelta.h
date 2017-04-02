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

	Used for excel *formula* rather than excel *objects*

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