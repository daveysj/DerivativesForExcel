Basic tools to price options in excel

Contents
=============================================================
 0: To Do
 1: Introduction
 2: Development Guidelines
 3: Requirements
 4: Windows: Installation
 5: Windows: Compiling & Linking
 
0: To Do 
=============================================================

1: Introduction
=============================================================
A sample, stand-alone set of functions for excel that can be used to price derivatives. Pricing formula are implemented using a methodology similar to that used by Microsoft to write the set of core functions in excel. 

There was a requirement to build a library that had a pre-defined excel signature (i.e. the order and type of the excel functions was pre-specified) so some of the code may seem a little artificial or overly specific. Examples include 
    - explicit assumptions that the time input is in days, 
    - a convergence input in the BlackVolOffSurface which I chose to hard code rather than leave as a parameter
    - a time parameter input in the Black76 and Delta formula which is unused in the code    

2: Development Guidelines
=============================================================
Stand-alone C++ functions are 
- Created in the main "DerivatievesForExcel" project and compiled as a static library
- Tested in "LibraryTest" using the Boost Testing framework
- Exposed to excel in the BasicExcelFormula library. To expose a new function to excel, always
    - Increment the NUM_FUNCTIONS variable in \excelIntegration\xllAddin.h to cater for the new number of functions
    - Add the function name to the xllDefinitiona.def file
    - Register the function in the registerXLLFunctions.cpp file
  This project creates an ".xll" addin for Microsoft excel in the appropriate debug or release library


3: Requirements
=============================================================
This code has been built and tested using
1) Microsoft Visual Studio 2015 Community Edition;
2) Boost 1.62.0 (using an environmental variable “boost” set to point to the root boost directory and "boostlib" pointing to the folder holding the compiled libraries - usually something equivalent to “$(boost)lib32-msvc-14.0\”


4: Windows: Installation 
=============================================================


5: Windows: Compiling & Linking 
=============================================================
Other than the links to Boost, this project is intended to be stand alone and should have no external dependencies.
