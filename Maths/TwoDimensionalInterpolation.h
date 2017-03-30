#ifndef XLLBasic_TWODIMINTERP_INCLUDED
#define XLLBasic_TWODIMINTERP_INCLUDED
#pragma once

//#include <ql\math\interpolations\all.hpp>
#include "maths.h"

namespace XLLBasicLibrary 
{

   /*======================================================================================
   TwoDimensionalInterpolation
    
    Abstract Base Class for 2 dimensional interpolation
   ======================================================================================*
    class TwoDimensionalInterpolator
    {
    public:
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
        // (locateX(xInput), locateY(yInput)+1)       (locateX(xInput)+1, locateY(yInput)+1)       
        // (locateX(xInput), locateY(yInput))         (locateX(xInput)+1, locateY(yInput))       
        size_t locateX(double xInput) const;
        size_t locateY(double yInput) const;

        double getXStart()   const {return x.front();};
        double getXEnd()   const {return x.back();};
        double getYStart()   const {return y.front();};
        double getYEnd()   const {return y.back();};

   protected:
        void setOnError(string errorMessage);

        vector<double> x, y;
        vector<vector<double> > z;

        bool allowExtrapolation;
        bool hasError; // used if any of the inputs are not of the assumed type
        string className;
        string errorMessage;      
    };

   /*======================================================================================
   BilinearInterpolator
    
   ======================================================================================*
    class BilinearInterpolator : public TwoDimensionalInterpolator 
    {
    public:
        BilinearInterpolator(vector<double> xVector, vector<double> yVector, QuantLib::Matrix zMatrix, bool extrapolate) : 
          TwoDimensionalInterpolator(xVector, yVector, zMatrix, extrapolate) {};

        ~BilinearInterpolator() {};

        virtual double getRate(double x, double y) const;

    };

   /*======================================================================================
   BicubicInterpolator
    
   ======================================================================================*
    class BicubicInterpolator : public TwoDimensionalInterpolator 
    {
    public:
        BicubicInterpolator(vector<double> xVector, vector<double> yVector, QuantLib::Matrix zMatrix, bool extrapolate);

        ~BicubicInterpolator() {};

        virtual double getRate(double x, double y) const;

    protected:
        //vector<CubicSplineInterpolator> splines;
        vector<QuantLib::Interpolation> splines;


    };
    */
}

#endif

