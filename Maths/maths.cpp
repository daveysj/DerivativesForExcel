#include "maths.h"

using namespace boost::algorithm;

namespace XLLBasicLibrary 
{


    //======================================================================================
    // TwoPointInterpolator
    //======================================================================================
    TwoPointInterpolator::TwoPointInterpolator(double x1, 
                                               double y1, 
                                               double x2, 
                                               double y2) : x1(x1), x2(x2), y1(y1), y2(y2) 
    {}

    bool TwoPointInterpolator::arePointsDistinct() const 
    {
        if (x1 == x2)
        {
            return false;
        }
        return true;
    }


    bool TwoPointInterpolator::doesExtrapolate(double x) const 
    {
        if (x1<=x2) 
        {
            if (!((x1<=x)&&(x<=x2)))
            {
                return true;
            }
        }
        else 
        {
            if (!((x2<=x)&&(x<=x1)))
            {
                return true;
            }
        }
        return false;
    }

    //======================================================================================
    // LinearInterpolator
    //======================================================================================
    LinearInterpolator::LinearInterpolator(double x1, 
                                           double y1,
                                           double x2, 
                                           double y2) : TwoPointInterpolator(x1, y1, x2, y2)
    {}

    double LinearInterpolator::getRate(double x) const 
    {
        double m = (y2-y1)/(x2-x1);
        double c = y1-m*x1; 
        return m*x+c;
    }

    //======================================================================================
    // ArrayInterpolator 
    //======================================================================================
    ArrayInterpolator::ArrayInterpolator(vector<double> xVector, 
                                         vector<double> yVector, 
                                         bool allowExtrapolation)
        : xVector(xVector), yVector(yVector), allowExtrapolation(allowExtrapolation)
    {
        isOk();
    }

    bool ArrayInterpolator::isOk()
    {
        if (!is_strictly_increasing(xVector.begin(), xVector.end()))
        {
            setOnError("ArrayInterpolator::ArrayInterpolator. X vector input not strictly monotonic");
            return false;
        }
        if (xVector.size() != yVector.size()) 
        {
           setOnError("ArrayInterpolator::ArrayInterpolator. X and Y vectors do not have the same dimension");
           return false;
        }
        if (xVector.size() < 2) 
        {
           setOnError("ArrayInterpolator::ArrayInterpolator. X vector must have a least 2 points");
           return false; 
        }
        errorMessage = "";
        hasError = false;
        return true;        
    }

    void ArrayInterpolator::setXVector(vector<double> xVectorInput) 
    {
        xVector = xVectorInput;
        if (!is_strictly_increasing(xVector.begin(), xVector.end()))
        {
            setOnError("ArrayInterpolator::setXVector. X vector input not strictly monotonic");
        }
        if (xVector.size() != yVector.size())
        {
            setOnError("ArrayInterpolator::setXVector. X and Y vectors do not have the same dimension");
        }
        else 
        {
            hasError = false;
        }
    }

   void ArrayInterpolator::setYVector(vector<double> yVectorInput)
   {
        yVector = yVectorInput;
        if (xVector.size() != yVector.size())
        {
            setOnError("ArrayInterpolator::setYVector. X and Y vectors do not have the same dimension");
        }
        else 
        {
            hasError = false;
        }
   }

   bool ArrayInterpolator::isInRange(double x) const 
    {
        if (hasError)
        {
            return false;
        }
        if (allowExtrapolation)
        {
            return true;
        }
        else if ((x >= xVector.front()) && (x <= xVector.back()))
        {
            return true;
        }
        return false;
   }
    
   bool ArrayInterpolator::isInRange(vector<double> xInput) const 
    {
      if (hasError)
        {
         return false;
        }
        if (allowExtrapolation)
        {
            return true;
        }
        bool allInRange = true;
        for (vector<double>::iterator it = xInput.begin(); it != xInput.end(); ++ it) 
        {
            double x = *it;
            if (!isInRange(x))
            {
                allInRange = false;
            }
        }
        return allInRange;
   }

    vector<double> ArrayInterpolator::getRate(vector<double> xInput) const
    {
        vector<double> yOutput;
        for (vector<double>::iterator it = xInput.begin(); it != xInput.end(); ++ it) 
        {
            double x = *it;
            yOutput.push_back(getRate(x));
        }
        return yOutput;
    }

    void ArrayInterpolator::setOnError(string errorMessageInput) 
    {
        hasError = true;
        errorMessage = errorMessageInput;
    }


    //======================================================================================
    // LinearArrayInterpolator 
    //======================================================================================
    LinearArrayInterpolator::LinearArrayInterpolator(vector<double> xVector, 
                                                     vector<double> yVector, 
                                                     bool inputAllowExtrapolation)
    : ArrayInterpolator(xVector, yVector, inputAllowExtrapolation)
    {}

    double LinearArrayInterpolator::getRate(double x) const
    {
        if (hasError)
        {
            return numeric_limits<double>::quiet_NaN();
        }

        if ((allowExtrapolation == false) && !isInRange(x))
        {
            return numeric_limits<double>::quiet_NaN();
        }
        // we are now either inside the range (if we do not allow extrapolation) or we allow extrapolation
        // using the "closest" points in the input arrays

        // We find the right place in the table by means of bisection. This will return the two "extreme" points
        // if we are extrapolating outside the curve
        int ilo = 0;
        int ihi = (int) xVector.size() - 1;
        int i;
        while (ihi - ilo > 1) 
        {
            i = (ihi + ilo) >> 1;
            if (xVector[i] > x) 
            {
                    ihi = i;
            }
            else 
            {
                ilo = i;
            }
        }
        double h = xVector[ihi] - xVector[ilo];
        if (h == 0) 
        {
            return numeric_limits<float>::quiet_NaN(); // shouldn't happen since we should check the x array is strictly monotonic but you never know
        }
        return LinearInterpolator(xVector[ilo], yVector[ilo], xVector[ihi], yVector[ihi]).getRate(x);
    }



    //======================================================================================
    // CubicSplineInterpolator 
    //======================================================================================
    CubicSplineInterpolator::CubicSplineInterpolator(vector<double> xVector, vector<double> yVector, double yp1, double ypn, bool allowExtrapolation)
        : ArrayInterpolator(xVector, yVector, allowExtrapolation), _yp1(yp1), _ypn(ypn), spline(xVector.size())
    {
        if (!hasError)
        {
            setSpline();
        }
    }

    CubicSplineInterpolator::CubicSplineInterpolator(std::vector<double> xVector, std::vector<double> yVector, bool allowExtrapolation)
        : ArrayInterpolator(xVector, yVector, allowExtrapolation), _yp1(0.0), _ypn(0.0), spline(xVector.size())
    {
        if (!hasError)
        {
            setSpline();
        }
    }

    double CubicSplineInterpolator::getRate(double x) const
    {
        if (hasError)
        {
            return numeric_limits<float>::quiet_NaN();
        }

        if (!allowExtrapolation && !isInRange(x))
        {
            return numeric_limits<float>::quiet_NaN();
        }

        // We find the right place in the table by means of bisection. This will return the two "extreme" points
        // if we are extrapolating outside the curve
        int klo = 0;
        int khi = (int) spline.size() - 1;
        int k;
        while (khi - klo > 1) 
        {
            k = (khi + klo) >> 1;
            if (xVector[k] > x) 
            {
                khi = k;
            }
            else 
            {
                klo = k;
            }
        }

        double h = xVector[khi] - xVector[klo];
        if (h == 0) 
        {
            return numeric_limits<float>::quiet_NaN(); // shouldn't happen since we should check the x array is strictly monotonic but you never know
        }
        double a = (xVector[khi] - x) / h;
        double b = (x - xVector[klo]) / h;
        double y = a * yVector[klo] + b * yVector[khi] + ((a*a*a - a) * spline[klo] + (b*b*b - b) * spline[khi]) * (h*h) / 6.0;
        return y;
    }

    void CubicSplineInterpolator::setSpline()
    {
        size_t n = spline.size();
        std::vector<double> tempCalcuation(n);
        spline[0] = -.5;
        if (_yp1 == 0)
        {
            tempCalcuation[0] = 0.0;
        }
        else
        {
            tempCalcuation[0] = (3.0 / (xVector[1] - xVector[0])) * ((yVector[1] - yVector[0]) / 
            (xVector[1] - xVector[0]) - _yp1);
        }
   
        double sig, p;
        for (int i = 1; i < n - 1; ++i) 
        {
            sig = (xVector[i] - xVector[i-1]) / (xVector[i+1] - xVector[i-1]);
            p = sig * spline[i-1] + 2.0;
            spline[i] = (sig - 1.0) / p;
            tempCalcuation[i] = (yVector[i+1] - yVector[i]) / (xVector[i+1] - xVector[i]) -
            (yVector[i] - yVector[i-1]) / (xVector[i] - xVector[i-1]);
            tempCalcuation[i] = (6.0 * tempCalcuation[i] / (xVector[i+1] - xVector[i-1]) - 
            sig * tempCalcuation[i-1]) / p;
        }
        double qn, un;
        if (_ypn == 0)
        {
            qn = un = 0;
        }
        else 
        {
            qn = 0.5;
            un = (3.0 / (xVector[n-1] - xVector[n-2])) * (_ypn - (yVector[n-1] - yVector[n-2]) / 
            (xVector[n-1] - xVector[n-2]));
        }
        spline[n-1] = (un - qn * tempCalcuation[n-2]) / (qn * spline[n-2] + 1.0);
        for (int k = n-2; k >= 0; k--)
        {
            spline[k] = spline[k] * spline[k+1] + tempCalcuation[k];
        }
    }

}