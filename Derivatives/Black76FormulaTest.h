#ifndef XLLBASIC_black76_test_maths
#define XLLBASIC_black76_test_maths
#pragma once

#include <iostream>
#include <boost\test\unit_test.hpp>
#include <boost\math\special_functions\fpclassify.hpp> // boost::math::isnan
#include "Black76Formula.h"

class Black76Test 
{
  public:
    static void testPutCallParity();

    static boost::unit_test_framework::test_suite* suite();
};

#endif


