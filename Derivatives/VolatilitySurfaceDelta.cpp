#include "VolatilitySurfaceDelta.h"

namespace XLLBasicLibrary
{
	/*======================================================================================
	SimpleDeltaSurface

	=======================================================================================*/
	SimpleDeltaSurface::SimpleDeltaSurface(
		vector<double> timesInput,
		vector<double> deltaInput,
		vector<vector<double>> volatilityInput,
		bool extrapolate,
		string interpolationType)
		: extrapolate(extrapolate), times(timesInput), delta(deltaInput), volatility(volatilityInput)
	{
		className = "SimpleDeltaSurface";
		string errorMessage;
		if (!checkAndTransformInputs(errorMessage))
		{
			throw runtime_error(errorMessage);
		}

		boost::algorithm::to_lower(interpolationType);
		boost::algorithm::trim(interpolationType);
		
		if (interpolationType.compare("bilinear") == 0)
		{
			interpolator = shared_ptr<TwoDimensionalInterpolator>(
				new BilinearInterpolator(
					times,
					delta,
					volatility,
					extrapolate));
		}
		else if (interpolationType.compare("bicubic") == 0)
		{
			interpolator = shared_ptr<TwoDimensionalInterpolator>(
				new BicubicInterpolator(
					times,
					delta,
					volatility,
					extrapolate));
		}
		else
		{
			throw runtime_error(className + "Interpolation Type must be either Bilinear or Bicubic");
		}
	}

	bool SimpleDeltaSurface::checkAndTransformInputs(string &reasonForFailure)
	{
		if (delta[0] < 1.0)
		{
			for (size_t i = 0; i < delta.size(); ++i)
			{
				delta[i] *= 100;
			}
		}
		// insert data at time 0 to ensure we can find sort dated volatility
		if (times[0] > 0)
		{
			times.insert(times.begin(), 0);
			for (size_t i = 0; i < volatility.size(); ++i)
			{
				double vol = volatility[i][0];
				volatility[i].insert(volatility[i].begin(), vol);
			}
		}
		// Vol is probably an integer not a decimal so change it
		if (volatility[0][0] > 2.0)
		{
			for (size_t i = 0; i < volatility.size(); ++i)
			{
				for (size_t j = 0; j < volatility[i].size(); ++j)
				{
					volatility[i][j] /= 100;
				}
			}
		}
		return true;
	}

	bool SimpleDeltaSurface::isInDeltaRange(double time, double delta)
	{
		if ((extrapolate) || interpolator->isInRange(time, delta))
		{
			return true;
		}
		return false;
	}

	bool SimpleDeltaSurface::isInMoneynessRange(double time, double moneyness)
	{
		if (extrapolate)
		{
			return true;
		}
		double fwd = 1;
		double strike = moneyness + fwd;
		double vol = calculateDeltaFromStrike(fwd, strike, time);
		if (abs(vol)<1e-8)
		{
			return false;
		}
		return true;
	}

	double SimpleDeltaSurface::getVolatility(double time) const
	{
		double delta = calculateDeltaFromStrike(1, 1, time);
		return getVolatilityForDelta(time, delta);
	}

	double SimpleDeltaSurface::getVolatilityForDelta(double time, double delta) const
	{
		if ((extrapolate) || (interpolator->isInRange(time, delta)))
		{
			return (*interpolator).getRate(time, delta);
		}
		return numeric_limits<double>::quiet_NaN();
	}

	double SimpleDeltaSurface::getVolatilityForMoneyness(double time, double moneyness) const
	{
		if (time <= 0)
		{
			return 0;
		}
		double fwd = 1;
		double strike = moneyness + fwd;
		double delta = calculateDeltaFromStrike(fwd, strike, time);
		return getVolatilityForDelta(time, delta);
	}

	// will return 0 if outside the interpolation range
	double SimpleDeltaSurface::calculateDeltaFromStrike(double forward, double strike, double time) const
	{
		double accuracy = 1.0e-8;
		size_t maxItterates = 20;
		double sqrtttm = sqrt(time);
		double guess1 = 50, guess2 = 50;
		double vol1 = 0, sd1 = 0.2, diff = accuracy + 1;
		Black76Put put(forward, strike, sd1, 1);
		size_t counter = 0;
		while ((counter < maxItterates) && (diff > accuracy))
		{
			guess1 = guess2;
			if (interpolator->isInRange(time, guess1))
			{
				vol1 = (*interpolator).getRate(time, guess1);
			}
			sd1 = vol1 * sqrtttm;
			put.setStandardDeviation(sd1);
			guess2 = -put.getDelta() * 100;
			diff = abs(guess1 - guess2);
			++counter;
		}
		return guess2;
	}
}