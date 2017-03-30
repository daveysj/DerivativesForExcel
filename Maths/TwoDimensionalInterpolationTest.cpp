#include "TwoDimensionalInterpolationTest.h"

//using namespace std;
using namespace boost::unit_test_framework;
using namespace boost::math;
using namespace XLLBasicLibrary;

#include <boost/assign/std/vector.hpp>
using namespace boost::assign; // used to initialize vector


void Maths2DInterpTest::testBilinearInterpolator()
{
    BOOST_TEST_MESSAGE("Testing BilinearInterpolation class ...");

    vector<double> time;
    time += 1, 2, 3, 6, 12, 24;

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

    BilinearInterpolator interpolator(time, delta, volatility, false);
    BOOST_REQUIRE(interpolator.isOk());
    BOOST_CHECK(interpolator.isInRange(1.5, 75));
    BOOST_CHECK(!interpolator.isInRange(0.5, 75));
    BOOST_CHECK(!interpolator.isInRange(36.5, 75));
    BOOST_CHECK(!interpolator.isInRange(1.5, 5));
    BOOST_CHECK(!interpolator.isInRange(1.5, 95));

    BOOST_CHECK(interpolator.locateX(0) == 0);
    BOOST_CHECK(interpolator.locateX(1.1) == 0);
    BOOST_CHECK(interpolator.locateX(2.1) == 1);
    BOOST_CHECK(interpolator.locateX(36) == 4);

    BOOST_CHECK(interpolator.locateY(0) == 0);
    BOOST_CHECK(interpolator.locateY(12) == 0);
    BOOST_CHECK(interpolator.locateY(26) == 1);
    BOOST_CHECK(interpolator.locateY(100) == 3);

    vector<LinearArrayInterpolator> lVector;
    lVector.push_back(LinearArrayInterpolator(time, v1, true));
    lVector.push_back(LinearArrayInterpolator(time, v2, true));
    lVector.push_back(LinearArrayInterpolator(time, v3, true));
    lVector.push_back(LinearArrayInterpolator(time, v4, true));
    lVector.push_back(LinearArrayInterpolator(time, v5, true));

    double xPoint = 9;
    double yPoint = 33;
    vector<double> lPoints = vector<double>(lVector.size());
    for (size_t i = 0; i < lVector.size(); ++i)
    {
        lPoints[i] = lVector[i].getRate(yPoint);
    }

    LinearArrayInterpolator finalInterpolator = LinearArrayInterpolator(delta, lPoints, true);

    BOOST_CHECK(abs(interpolator.getRate(xPoint, yPoint) - finalInterpolator.getRate(xPoint)) < 1e-12);
    BOOST_CHECK(boost::math::isnan<double>(interpolator.getRate(9, 95)));


}

void Maths2DInterpTest::testBicubicInterpolator()
{
    BOOST_TEST_MESSAGE("Testing BilinearCubic class ...");
    /*
    vector<double> time;
    time += 1, 2, 3, 6, 12, 24;

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

    QuantLib::Matrix M = QuantLib::Matrix(time.size(), delta.size());
    for (size_t i = 0; i < delta.size(); ++i) {
        for (size_t j = 0; j < time.size(); ++j) {
            M[i][j] = volatility[i][j];
        }
    }

    BicubicInterpolator interpolator(time, delta, M, false);
    BOOST_REQUIRE(interpolator.isOk());
   BOOST_CHECK(interpolator.isInRange(1.5, 75));
   BOOST_CHECK(!interpolator.isInRange(0.5, 75));
   BOOST_CHECK(!interpolator.isInRange(36.5, 75));
   BOOST_CHECK(!interpolator.isInRange(1.5, 5));
   BOOST_CHECK(!interpolator.isInRange(1.5, 95));

   BOOST_CHECK(interpolator.locateX(0) == 0);
   BOOST_CHECK(interpolator.locateX(1.1) == 0);
   BOOST_CHECK(interpolator.locateX(2.1) == 1);
   BOOST_CHECK(interpolator.locateX(36) == 4);

   BOOST_CHECK(interpolator.locateY(0) == 0);
   BOOST_CHECK(interpolator.locateY(12) == 0);
   BOOST_CHECK(interpolator.locateY(26) == 1);
   BOOST_CHECK(interpolator.locateY(100) == 3);

    QuantLib::BicubicSpline qlInterp = QuantLib::BicubicSpline(time.begin(), time.end(), delta.begin(), delta.end(), M);
    double me = interpolator.getRate(6, 50);
    double ql = qlInterp(6, 50);
   BOOST_CHECK(interpolator.getRate(6, 50) == qlInterp(6, 50));
   BOOST_CHECK(interpolator.getRate(1, 10) == qlInterp(1, 10));
   BOOST_CHECK(interpolator.getRate(24, 90) == qlInterp(24, 90));
   BOOST_CHECK(interpolator.getRate(9, 33) == qlInterp(9, 33));
   BOOST_CHECK(boost::math::isnan<double>(interpolator.getRate(9, 95)));
   */

}



test_suite* Maths2DInterpTest::suite() 
{
    test_suite* suite = BOOST_TEST_SUITE("Maths TwoDimnsionalInterpolation Tests");

    suite->add(BOOST_TEST_CASE(&Maths2DInterpTest::testBilinearInterpolator));
    suite->add(BOOST_TEST_CASE(&Maths2DInterpTest::testBicubicInterpolator));

    return suite;
}

