#ifndef XLLBASIC_test_2dinterp
#define XLLBASIC_2dinterp

#include <iostream>
#include <boost\test\unit_test.hpp>
#include <boost\math\special_functions\fpclassify.hpp> // boost::math::isnan
#include "TwoDimensionalInterpolation.h"

class Maths2DInterpTest 
{
  public:

    static void testBilinearInterpolator();
    static void testBicubicInterpolator();

    static boost::unit_test_framework::test_suite* suite();
};


#endif
