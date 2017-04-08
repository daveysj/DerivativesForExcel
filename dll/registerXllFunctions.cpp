
#include "excelIntegration\xllAddIn.h"


//========================================================================
char *AddinVersionStr = "Derivative pricing";
char *AddinName = "Pricing";
char *DevAddinName = "Pricing - DEV";

//---------------------------------------------------------
// Registered function argument and return types
// Data type					Pass by value		Pass by ref
// boolean						A					L
// double						B					E
// char*											C, F
// unsigned char*									D, G
// unsigned short int (DWORD)	H
// signed short int				I					M
// signed long int				J					N
// struct FP (xl_array)								K
// struct oper										P
// struct xloper									R
//---------------------------------------------------------
char *FunctionExports[NUM_FUNCTIONS][MAX_EXCEL4_ARGS - 1] =
{
	{
		"Interpolate",
		"PBKKICA",
		"Interpolate",
		"x, x array, y array, size, Type, Extrap",
		"1",
		AddinName,
		"",
		"",
		"Find the y-value corresponding to the input x-value for the points (x_1, y_1),...(x_n,y_n)",
		// Help text line (optional)
		"x Value at which to interpolate",
		"Array containing all the x inputs (must be strictly increasing or decreasing)",
		"Array containing all the y inputs",
		"Size of the input arrays to use in the interpolator",
		"\"Linear\" or \"Cubic\" interpolation (default = \"Linear\")",
		"Allow extrapolation (default = false)",
		"",
	},
	{
		"BlackVolOffSurface",
		"PCBBBKKKBCA",
		"BlackVolOffSurface",
		"OptionType,Forward,Strike,Day,Days Array,Put Delta,Volatility,Convergence,InterpType,Extrap",
		"1",
		AddinName,
		"",
		"",
		"The volatility of a Black-Scholes option from a delta based surface. Explict assumptions "
		"include that *time* is measured in *days*; delta is that of a *put* option; "
		"volatility is a number between 0 and 100 i.e. not a percentage",
        // Help text line (optional)
        "Option Type = (P)ut or (C)all",
		"Market forward",
		"Option strike",
		"Day to interpolate to",
		"Days array (NB Time is explicitly assumed to be *DAYS*)",
		"Delta array (NB Delta is explicitly assumed to be for a *PUT* option)",
		"Volatility Surface",
		"Convergence (Hard coded - input preserved to keep the function signature constant)",
		"Linear or Cubic (default = linear)",
        "Allow extrapolation (default = false)",
        "",
    },
    {
        "Black",
        "RCBBBBB",
        "Black",
        "P/C,forward,strike,dtm,sd,df",
        "1",
        AddinName,
        "",
        "",
        "Returns the PV premium of a Black 76 option on a future / forward",
        // Help text line (optional)
        "(P)ut or (C)all",
        "Forward",
        "Strike",
		"Days to maturity"
        "Standard Deviation (=vol*sqrt(time))",
        "Discount Factor",
        "",
    },
    {
        "BlackDelta",
        "RCBBBBB",
        "BlackDelta",
        "P/C,forward,strike,dtm,sd,df",
        "1",
        AddinName,
        "",
        "",
        "Returns the analytic delta of a Black 76 option on a future / forward",
        // Help text line (optional)
        "(P)ut or (C)all",
        "Forward",
        "Strike",
		"Days to maturity"
		"Standard Deviation (=vol*sqrt(time))",
        "Discount Factor",
        "",
    },
};
