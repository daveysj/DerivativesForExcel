#include "XLLBasicLibraryTest.h"

using namespace boost::unit_test_framework;


namespace 
{
    boost::timer t;

    void startTimer() { t.restart(); }
    void stopTimer() {
        double seconds = t.elapsed();
        int hours = int(seconds/3600);
        seconds -= hours * 3600;
        int minutes = int(seconds/60);
        seconds -= minutes * 60;
        std::cout << " \nTests completed in ";
        if (hours > 0)
            std::cout << hours << " h ";
        if (hours > 0 || minutes > 0)
            std::cout << minutes << " m ";
        std::cout << std::fixed << std::setprecision(0)
                  << seconds << " s\n" << std::endl;
    }
}

test_suite* init_unit_test_suite(int, char* []) {

    std::string header = "Testing XLLBasic Library";
    std::string rule = std::string(header.length(),'=');

    BOOST_TEST_MESSAGE(rule);
    BOOST_TEST_MESSAGE(header);
    BOOST_TEST_MESSAGE(rule);
    test_suite* test = BOOST_TEST_SUITE("XLLBasic Library test suite");

    test->add(BOOST_TEST_CASE(startTimer));

    test->add(MathsFunctionsTest::suite());
    test->add(Maths2DInterpTest::suite());    


    test->add(BOOST_TEST_CASE(stopTimer));

    return test;
}