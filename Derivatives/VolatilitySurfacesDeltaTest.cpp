#include "VolatilitySurfacesDeltaTest.h"

#include <boost/assign/std/vector.hpp>
using namespace boost::assign; // used to initialize vector

using namespace std;
using namespace boost::unit_test_framework;
using namespace XLLBasicLibrary;


void VolatilitySurfacesDeltaTest::testSimpleDeltaSurfaceConstruction()
{
    BOOST_TEST_MESSAGE("Testing SimpleDeltaSurface Construction ...");

    vector<double> observationTimes;
	observationTimes += 1.0 / 12.0, 2.0 / 12.0, 0.25, 0.5, 1.0, 2.0;

    vector<double> delta;
    delta += 10, 25, 50, 75, 90;

    vector<double> v1, v2, v3, v4, v5;
    v1 += .17938,   .182884,    .193908,    .219688,    .248396,    .263268; // 10 Delta put
    v2 += .17575,   .17575,     .18247,     .206225,    .234775,    .2475;
    v3 += .175,     .175,       .18,        .205,       .235,       .2475;
    v4 += .18825,   .18825,     .19547,     .223725,    .223725,    .2725;
    v5 += .20128,   .204784,    .216708,    .250288,    .287796,    .307068; // 90 Delta put

    vector<vector<double>> volatility;
    volatility += v1, v2, v3, v4, v5;

	shared_ptr<SimpleDeltaSurface> vs;
	try 
	{
		bool extrapolate = false;
		string interpolationType = "bilinear";
		vs = shared_ptr<SimpleDeltaSurface>(
			new SimpleDeltaSurface(observationTimes, delta, volatility, extrapolate, interpolationType));
	}
	catch (runtime_error &e)
	{
		BOOST_FAIL(e.what());
	}

	double time = 1.0;
	double forward = 100;
	double strike = 90; 
	double moneyness = (strike - forward) / forward;
	BOOST_REQUIRE(vs->isInMoneynessRange(time, moneyness));
	cout << "Vol: " << vs->getVolatilityForMoneyness(time, moneyness);
	BOOST_CHECK(abs(vs->getVolatilityForMoneyness(time, moneyness) - 0.234807) < 1e-6);
}

test_suite* VolatilitySurfacesDeltaTest::suite() 
{
    test_suite* suite = BOOST_TEST_SUITE("Volatility Surfaces");
        
    suite->add(BOOST_TEST_CASE(&VolatilitySurfacesDeltaTest::testSimpleDeltaSurfaceConstruction));
    
    return suite;
}

