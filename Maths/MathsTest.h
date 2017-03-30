#ifndef sjd_test_maths
#define sjd_test_maths

#include <iostream>
#include <boost\test\unit_test.hpp>
#include <boost\math\special_functions\fpclassify.hpp> // boost::math::isnan
#include "maths.h"

class MathsFunctionsTest 
{
public:
    static void testLinearInterpolator();
    // Test some of the functionality of the abstract base class using one instance
    static void testArrayInterpolator();
    static void testLinearArrayInterpolator();
    static void testCubicSplineInterpolator();   

    static boost::unit_test_framework::test_suite* suite();
};

#endif
