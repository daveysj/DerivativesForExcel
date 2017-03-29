#include "MathsTest.h"

//using namespace std;
using namespace boost::unit_test_framework;
using namespace boost::math;
using namespace XLLBasicLibrary;

#include <boost/assign/std/vector.hpp>
using namespace boost::assign; // used to initialize vector


void MathsFunctionsTest::testFunctions() 
{
    BOOST_TEST_MESSAGE("Testing Maths Functions...");
    vector<double> increasingVector, strictlyIncreasingVector, mixedVector, emptyVector;
    vector<size_t> decreasingVector, strictlyDecreasingVector, constantVector;
    decreasingVector.push_back(10);
    for (size_t z = 0; z < 10; ++ z) {
        increasingVector.push_back((double) z);
        strictlyIncreasingVector.push_back((double) z);
        decreasingVector.push_back(10-z);
        strictlyDecreasingVector.push_back(10-z);
    }
    increasingVector.push_back(9);
    mixedVector = increasingVector;
    for (size_t i = 0; i < decreasingVector.size(); ++ i) {
        mixedVector.push_back(decreasingVector[i]);
        constantVector.push_back(10);
    }

    BOOST_CHECK(isStrictlyIncreasing(strictlyIncreasingVector) == true);
    BOOST_CHECK(isStrictlyIncreasing(increasingVector) == false);
    BOOST_CHECK(isStrictlyIncreasing(constantVector) == false);
    BOOST_CHECK(isStrictlyIncreasing(emptyVector) == false);

    BOOST_CHECK(isStrictlyDecreasing(strictlyDecreasingVector) == true);
    BOOST_CHECK(isStrictlyDecreasing(decreasingVector) == false);
    BOOST_CHECK(isStrictlyDecreasing(constantVector) == false);
    BOOST_CHECK(isStrictlyDecreasing(emptyVector) == false);

    BOOST_CHECK(isStrictlyMonotonic(strictlyIncreasingVector) == true);
    BOOST_CHECK(isStrictlyMonotonic(strictlyDecreasingVector) == true);
    BOOST_CHECK(isStrictlyMonotonic(decreasingVector) == false);
    BOOST_CHECK(isStrictlyMonotonic(constantVector) == false);
    BOOST_CHECK(isStrictlyMonotonic(emptyVector) == false);
}

void MathsFunctionsTest::testLinearInterpolator() 
{
    BOOST_TEST_MESSAGE("Testing LinearInterpolator class ...");

    double x1 = 1, x2 = 10, y1 = 1, y2 = 10;
    LinearInterpolator linearInterpolator(x1, y1, x2, y2);
    BOOST_CHECK(linearInterpolator.getRate(2) == 2); // interpolate
    BOOST_CHECK(linearInterpolator.getRate(15.5) == 15.5) ; // extrapolate
    BOOST_CHECK(linearInterpolator.doesExtrapolate(5) == false);
    BOOST_CHECK(linearInterpolator.doesExtrapolate(10) == false); // end point
    BOOST_CHECK(linearInterpolator.doesExtrapolate(-1) == true);

    x2 = 1; // x1 == x2, should result in NaN
    linearInterpolator = LinearInterpolator(x1, y1, x2, y2);
    BOOST_CHECK(linearInterpolator.arePointsDistinct() == false);
}

void MathsFunctionsTest::testArrayInterpolator() 
{
    BOOST_TEST_MESSAGE("Testing ArrayInterpolator class ...");
    vector<double> xVector, yVector;

    // non monotonic inputs
    double m = 2.4, c = -10;
    for (int i = 0; i < 5; ++i) 
    {
        double x = (double) i;
        xVector.push_back(x);
        yVector.push_back(m * x + c);
    }
    double x = 4;
    xVector.push_back(x);
    yVector.push_back(m * x + c);
    LinearArrayInterpolator interp(xVector, yVector, true);    
    BOOST_CHECK(interp.isOk() == false);
    BOOST_CHECK(boost::math::isnan<double>(interp.getRate(3.5)) == true);

    xVector.pop_back();
    xVector.push_back(5);
    xVector.push_back(6);
    interp = LinearArrayInterpolator(xVector, yVector, true); // inconsistent dimensions
    BOOST_CHECK(interp.isOk() == false);
    BOOST_CHECK(boost::math::isnan<double>(interp.getRate(3.5)) == true);

    xVector.clear();
    yVector.clear();
    interp = LinearArrayInterpolator(xVector, yVector, true); // no data
    BOOST_CHECK(interp.isOk() == false);
    BOOST_CHECK(boost::math::isnan<double>(interp.getRate(3.5)) == true);

    for (size_t i = 0; i < 5; ++i) 
    {
        double x = (double) i;
        xVector.push_back(x);
        yVector.push_back(m * x + c);
    }
    LinearArrayInterpolator interpWithExtrap(xVector, yVector, true);    
    BOOST_CHECK(interpWithExtrap.isInRange(1.75) == true);
    BOOST_CHECK(interpWithExtrap.isInRange(10.75) == true);
    BOOST_CHECK(interpWithExtrap.isInRange(-.28) == true);

    vector<double> xInputs;
    xInputs.push_back(0.5);
    xInputs.push_back(1.5);
    BOOST_CHECK(interpWithExtrap.isInRange(xInputs) == true);
    xInputs.push_back(5.5);
    BOOST_CHECK(interpWithExtrap.isInRange(xInputs) == true);

    vector<double> yVector2, xVector2;
    for (size_t i = 0; i < 6; ++i) 
    {
        double x = (double) i;
        xVector2.push_back(x);
        yVector2.push_back(m * x + c);
    }
    interpWithExtrap.setYVector(yVector2);
    BOOST_CHECK(interpWithExtrap.isOk() == false);
    interpWithExtrap.setXVector(xVector2);
    BOOST_CHECK(interpWithExtrap.isOk() == true);

    // Check polymorphic behaviour for getRange(vector)
    xInputs.clear();
    vector<double> yOutputs1, yOutputs2;
    x = 0.5;
    xInputs.push_back(x);
    double y = interpWithExtrap.getRate(x);
    yOutputs1.push_back(y);
    x = 1.5;
    xInputs.push_back(x);
    yOutputs1.push_back(interpWithExtrap.getRate(x));
    yOutputs2 = interpWithExtrap.getRate(xInputs);  
    BOOST_CHECK(yOutputs1.size() == yOutputs2.size());
    for (size_t i = 0; i < yOutputs1.size(); ++i)
    {
        BOOST_CHECK(yOutputs1[i] == yOutputs2[i]);
    }
}

void MathsFunctionsTest::testLinearArrayInterpolator()
{
    BOOST_TEST_MESSAGE("Testing LinearArrayInterpolator class ...");

    vector<double> xVector, yVector;
    double m = 2.4, c = -10;
    for (size_t i = 0; i < 6; ++i) 
    {
        double x = (double) i;
        xVector.push_back(x);
        yVector.push_back(m * x + c);
    }

    double epsilon = 1e-10;
    LinearArrayInterpolator interpWithExtrap(xVector, yVector, true);    
    LinearArrayInterpolator interpWithoutExtrap(xVector, yVector, false);    
    BOOST_CHECK(interpWithExtrap.getRate(3.5) == (m * 3.5 + c));
    BOOST_CHECK(interpWithExtrap.getRate(4.75) == interpWithoutExtrap.getRate(4.75));
    BOOST_CHECK(abs(interpWithExtrap.getRate(10.75) - (m * 10.75 + c)) < epsilon);
    BOOST_CHECK(boost::math::isnan<double>(interpWithoutExtrap.getRate(10.75)) == true);
}

void MathsFunctionsTest::testCubicSplineInterpolator() 
{
    vector<double> xVector, yVector;
    double a = 2.4, b = -3.1, c = -10;
    for (size_t i = 0; i < 6; ++i) 
    {
        double x = (double) i;
        xVector.push_back(x);
        yVector.push_back(a * x * x + b * x + c);
    }
    CubicSplineInterpolator cSplineInpterp1(xVector, yVector, 0, 0, true);
    CubicSplineInterpolator cSplineInpterp2(xVector, yVector, true);

    BOOST_CHECK(cSplineInpterp1.getRate(0.73) == cSplineInpterp2.getRate(0.73)); 
    BOOST_CHECK(cSplineInpterp1.getRate(1.73) == -8.24323396243094);
}


test_suite* MathsFunctionsTest::suite() 
{
    test_suite* suite = BOOST_TEST_SUITE("Maths Functions Tests");
    suite->add(BOOST_TEST_CASE(&MathsFunctionsTest::testFunctions));
    suite->add(BOOST_TEST_CASE(&MathsFunctionsTest::testLinearInterpolator));
    suite->add(BOOST_TEST_CASE(&MathsFunctionsTest::testArrayInterpolator));   
    suite->add(BOOST_TEST_CASE(&MathsFunctionsTest::testLinearArrayInterpolator));
    suite->add(BOOST_TEST_CASE(&MathsFunctionsTest::testCubicSplineInterpolator));

    return suite;
}

