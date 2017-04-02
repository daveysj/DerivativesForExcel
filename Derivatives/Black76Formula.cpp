#include "Black76Formula.h"


namespace XLLBasicLibrary 
{

   /*======================================================================================
   Black76Option
   =======================================================================================*/
    void Black76Option::setParameters(double forward, double strike, double standardDeviation, double discountFactor)
    {
        setForward(forward);
        setStrike(strike);
        setStandardDeviation(standardDeviation);
        setDiscountFactor(discountFactor);
    }

   void Black76Option::calculateInternalOptionParameters()
    {
       d1 = (log(F/X) / sd + sd / 2.0);
       d2 = d1 - (sd);     
       Nd1 = cdf(n_0_1, d1);
       Nd2 = cdf(n_0_1, d2);
    }

    void Black76Option::setForward(double forwardInput)    
    {
		if (forwardInput <= 0)
		{
			throw runtime_error("Black76Option->Forward is <= 0");
		}
        F = forwardInput;
    }

    void Black76Option::setStrike(double strikeInput)      
    {
		if (strikeInput <= 0)
		{
			throw runtime_error("Black76Option->Strike is <= 0");
		}
        X = strikeInput;
    }

    void Black76Option::setStandardDeviation(double standardDeviation)    
    {
		if (standardDeviation <= 0)
		{
			throw runtime_error("Black76Option->Standard Deviation is <= 0");
		}
        sd = standardDeviation;
    }

    void Black76Option::setDiscountFactor(double discountFactor)       
    {
		if (discountFactor <= 0)
		{
			throw runtime_error("Black76Option->Discount Factor is <= 0");
		}
		df = discountFactor;
    }


    /*======================================================================================
    EuropeanCallBlack76
    =======================================================================================*/
	Black76Call::Black76Call(double f, double x, double sd, double df)
	{
		setParameters(f, x, sd, df);
	}

	double Black76Call::getPremium() 
	{
		calculateInternalOptionParameters();
		return df * (F *  Nd1 - X * Nd2);
	}

	double Black76Call::getDelta()
	{
		calculateInternalOptionParameters();
		return Nd1 * df;
	}


	/*======================================================================================
	EuropeanPutBlack76
	=======================================================================================*/
	Black76Put::Black76Put(double F, double X, double sd, double df)
	{
		setParameters(F, X, sd, df);
	}

	double Black76Put::getPremium()
	{
		calculateInternalOptionParameters();
			return df * (- F * (1-Nd1) + X * (1-Nd2));
	}

	double Black76Put::getDelta()
	{
		calculateInternalOptionParameters();
		return (Nd1 - 1.0) * df;
	}

}