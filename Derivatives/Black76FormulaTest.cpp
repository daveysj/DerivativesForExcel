
#include "Black76FormulaTest.h"

using namespace std;
using namespace boost::unit_test_framework;
using namespace boost::math;
using namespace XLLBasicLibrary;

void Black76Test::testPutCallParity() 
{
    BOOST_TEST_MESSAGE("Testing Black Scholes Put / Call parity ...");

    double F = 100, X = 110, sd = 0.2, df = 0.97;
    Black76Call call(F,X,sd,df);
    Black76Put put(F,X,sd,df);

    // Put call parity
    BOOST_CHECK(abs(put.getPremium() - call.getPremium() - (X-F) * df) < 1e-12);
}


test_suite* Black76Test::suite() 
{
    test_suite* suite = BOOST_TEST_SUITE("Black 76 Option Pricing Suite");
    suite->add(BOOST_TEST_CASE(&Black76Test::testPutCallParity));

    return suite;
}

