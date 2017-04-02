#ifndef XLLBasic_TWODIMINTERP_INCLUDED
#define XLLBasic_TWODIMINTERP_INCLUDED
#pragma once

//#include <ql\math\interpolations\all.hpp>
#include "maths.h"

namespace XLLBasicLibrary 
{

   /*======================================================================================
   TwoDimensionalInterpolation
    
    Abstract Base Class for 2 dimensional interpolation. Note the dimensions
    Assuming we want a volatility surface with time / delta dimensions with time being
    the "row" or first dimension and delta being the "column" or second dimension.

    Then the inputs would be built as:
        timeVector = [0, 1/12, 2/12, 3/12, 6/12, 12/12]
        deltaVector = [0.25, 50, 75]

    Then, "using namespace boost::assign;", construct the input vector of vectors as
        vector<double> v0, v1, v2;
        v0 += .17938,   .182884,    .193908,    .219688,    .248396,    .263268; // 25 Delta put
        v1 += .175,     .175,       .18,        .205,       .235,       .2475;
        v2 += .20128,   .204784,    .216708,    .250288,    .287796,    .307068; // 75 Delta put

        vector<vector<double>> volatility;
        volatility += v0, v1, v2;

   ======================================================================================*/
    class TwoDimensionalInterpolator
    {
    public:
        // Throws a runtime error if the inputs are not correct, i.e. if 
        //  - xVector and yVector are not strictly increasing
        //  - dimensions of zMatrix are not consistent
        TwoDimensionalInterpolator(
            vector<double> xVector,
            vector<double> yVector,
            vector<vector<double> > zMatrix,
            bool extrapolate = false);

        virtual ~TwoDimensionalInterpolator()   {};
   
        virtual void setX(vector<double> xVector)             {x = xVector;};
        virtual void setY(vector<double> yVector)             {y = yVector;};
        virtual void setZ(vector<vector<double> > zMatrix)    {z = zMatrix;};

        virtual bool isOk();
        string getErrorMessage() const  {return errorMessage;};

        // Returns true if the input (x, y) point is inside the interpolation range
        bool isInRange(double x, double y) const;
        // I assume yu have called isInRange(x, y) by this stage if you need to
        virtual double getRate(double x, double y) const = 0;

        // given a point (xInput, yInput) we use the following methods to find the "boundary" 
        size_t locateX(double xInput) const;
        size_t locateY(double yInput) const;

        double getXStart()   const {return x.front();};
        double getXEnd()   const {return x.back();};
        double getYStart()   const {return y.front();};
        double getYEnd()   const {return y.back();};

   protected:

        vector<double> x, y;
        vector<vector<double> > z;

        bool allowExtrapolation;
        bool hasError; // used if any of the inputs are not of the assumed type
        string className;
        string errorMessage;      
    };

   /*======================================================================================
   BilinearInterpolator
    
   ======================================================================================*/
    class BilinearInterpolator : public TwoDimensionalInterpolator 
    {
    public:
        BilinearInterpolator(
            vector<double> xVector, 
            vector<double> yVector, 
            vector<vector<double> > zMatrix, 
            bool extrapolate);

        ~BilinearInterpolator() {};

        virtual double getRate(double x, double y) const;

    };

   /*======================================================================================
   BicubicInterpolator
    
   ======================================================================================*/
    class BicubicInterpolator : public TwoDimensionalInterpolator 
    {
    public:
        BicubicInterpolator(
            vector<double> xVector, 
            vector<double> yVector, 
            vector<vector<double> > zMatrix, 
            bool extrapolate);

        ~BicubicInterpolator() {};

        virtual double getRate(double x, double y) const;

    protected:
        vector<CubicSplineInterpolator> splines;
        //vector<QuantLib::Interpolation> splines;


    };
}

#endif

