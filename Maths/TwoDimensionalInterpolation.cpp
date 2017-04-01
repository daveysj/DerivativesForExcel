#include "TwoDimensionalInterpolation.h"

using namespace boost::algorithm;

namespace XLLBasicLibrary
{

    /*======================================================================================
    TwoDimensionalInterpolator
    
    ======================================================================================*/
    TwoDimensionalInterpolator::TwoDimensionalInterpolator(
        vector<double> xVector,
        vector<double> yVector,
        vector<vector<double> > zMatrix,
        bool extrapolate) :
        x(xVector), y(yVector), z(zMatrix), allowExtrapolation(extrapolate)
    {
        className = "TwoDimensionalInterpolation";
        if (!isOk())
        {
            throw runtime_error(errorMessage);
        }
    }


    bool TwoDimensionalInterpolator::isOk() 
    {
        size_t zRows = z.size();
        size_t zColumns = z[0].size();

        if ((x.size() < 2) || (y.size() < 2) || (zRows < 2) || (zColumns < 2))
        {
            errorMessage = className + ": The input x, y or z data does not have sufficient data";
            return false;
        }
        for (size_t i = 1; i < zRows; ++i)
        {
            if (z[i].size() != zColumns)
            {
                errorMessage = className + ": Columns do not have consistent dimension";
                return false;
            }
        }
        if (!is_strictly_increasing(x.begin(), x.end()))
        {
            errorMessage = className + ": X inputs not structly increasing";
            return false;
        }
        if (!is_strictly_increasing(y.begin(), y.end()))
        {
            errorMessage = className + ": Y inputs not structly increasing";
            return false;
        }
        if (y.size() != zRows) 
        {
            errorMessage = className + ": Y vector has inconsistent dimension with z data";
            return false;
        }
        if (x.size() != zColumns) 
        {
            errorMessage = className + ": X vector has inconsistent dimension with z data";
            return false;
        }
        errorMessage = className + ": No Error";
        return true;
    }

    bool TwoDimensionalInterpolator::isInRange(double x, double y) const
    {
        if (allowExtrapolation)
        {
            return true;
        }
        else if ((x >= getXStart()) && (x <= getXEnd()) && (y >= getYStart()) && (y <= getYEnd()))
        {
            return true;
        }
        return false;
    }

    size_t TwoDimensionalInterpolator::locateX(double xInput) const
    {
        if (xInput < getXStart())
        {
            return 0;
        }
        else if (xInput > getXEnd())
        {
            return x.size() - 2;
        }
        size_t index = 0;
        while ((index < x.size()) && (x[index] < xInput))
        {
            ++index;
        }
        if (index > 0)
        {
            --index;
        }
        return index;
    }

    size_t TwoDimensionalInterpolator::locateY(double yInput) const 
    {
        if (yInput < getYStart())
        {
            return 0;
        }
        else if (yInput > getYEnd())
        {
            return y.size() - 2;
        }
        size_t index = 0;
        while ((index < y.size()) && (y[index] < yInput))
        {
            ++index;
        }
        if (index > 0)
        {
            --index;
        }
        return index;
    }


    /*======================================================================================
    BilinearInterpolator
    
    ======================================================================================*/
    BilinearInterpolator::BilinearInterpolator(
        vector<double> xVector, 
        vector<double> yVector, 
        vector<vector<double> > zMatrix,
        bool extrapolate) :
        TwoDimensionalInterpolator(xVector, yVector, zMatrix, extrapolate)
    {
        className = "BilinearInterpolator";
    };

    double BilinearInterpolator::getRate(double xInput, double yInput) const
    {
        if (!isInRange(xInput, yInput))
        {
            return numeric_limits<double>::quiet_NaN();
        }

        size_t i = locateX(xInput);
        double x1 = x[i], x2 = x[i + 1];

        size_t j = locateY(yInput);
        double y1 = y[j], y2 = y[j + 1];

        double z1 = z[j][i],
               z2 = z[j+1][i],
               z3 = z[j][i+1],
               z4 = z[j+1][i+1];

      double t = (xInput-x1)/(x2-x1);
      double u = (yInput-y1)/(y2-y1);
      return (1-t)*(1-u)*z1 + t*(1-u)*z3 + t*u*z4 + (1-t)*u*z2;
    }

   /*======================================================================================
   BicubicInterpolator
    
   ======================================================================================*/
    BicubicInterpolator::BicubicInterpolator(
        vector<double> xVector, 
        vector<double> yVector, 
        vector<vector<double> > zMatrix, 
        bool extrapolate) :
        TwoDimensionalInterpolator(xVector, yVector, zMatrix, extrapolate) 
    {
        className = "BicubicInterpolator";

        for (size_t i = 0; i < y.size(); ++i) 
        {
            vector<double> z_i = z[i];
            splines.push_back(CubicSplineInterpolator(xVector, z_i, 0, 0, false));

            //splines.push_back(CubicInterpolation(
            //    x.begin(), x.end(),
            //    this->z.row_begin(i),
            //    CubicInterpolation::Spline, false,
            //    CubicInterpolation::SecondDerivative, 0.0,
            //    CubicInterpolation::SecondDerivative, 0.0));
        }
    };

    double BicubicInterpolator::getRate(double xInput, double yInput) const
    {
        if (!isInRange(xInput, yInput))
        {
            return numeric_limits<double>::quiet_NaN();
        }

        std::vector<double> section(splines.size());
        for (size_t i = 0; i < splines.size(); i++)
        {
            section[i] = splines[i].getRate(xInput);
        }

        CubicSplineInterpolator spline = CubicSplineInterpolator(y, section, 0, 0, true);

        return spline.getRate(yInput);
    }
}