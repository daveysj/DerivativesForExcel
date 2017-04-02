#ifndef sjdvolatilitysurfaces_test_maths
#define sjdvolatilitysurfaces_test_maths

#include <iostream>
#include <boost\test\unit_test.hpp>
#include <boost\math\special_functions\fpclassify.hpp> // boost::math::isnan
#include "VolatilitySurfaceDelta.h"


class VolatilitySurfacesDeltaTest 
{
  public:      
    static void testSimpleDeltaSurfaceConstruction();

    static boost::unit_test_framework::test_suite* suite();

};

#endif
