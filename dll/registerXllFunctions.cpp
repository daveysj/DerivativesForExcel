
#include "excelIntegration\xllAddIn.h"


//========================================================================
char *AddinVersionStr = "Simple pricing add ins for Excel";
char *AddinName = "SBK";
char *DevAddinName = "SBK - DEV";

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
        "xllInterpolate",
        "RBKKICA",
        "xllInterpolate",
        "x, x array, y array, size, Type, Extrap",
        "1",
        AddinName,
        "",
        "",
        "Interpolation algorithm",
        // Help text line (optional)
        "Value to interpolate"
        "Input x array",
        "Input y array",
        "Value(s) at which to interpolate",
        "Linear or Cubic (default = linear)",
        "Allow extrapolation (default = false)",
        "",
    },
    {
        "prVolFromDeltaSurface",
        "RCBKBBBA",
        "prVolFromDeltaSurface",
        "Type,anchorDate,surface,toDate,strike,forward,Extrap",
        "1",
        AddinName,
        "",
        "",
        "Returns the volatility to use for a Black-Scholes option with a strike from a delta based suface",
        // Help text line (optional)
        "BiLinear or BiCubic interpolation",
        "The anchor date for the volatility surface",
        "The entire volatility surface including the time / dates in the top row and the PUT deltas in the first column. Note that the first entry A1 is ignored",
        "The value to interpolate to in the time dimension",
        "The strike of the option",
        "The market forward for the option",
        "Allow extrapolation (default = false)",
        "",
    },
    {
        "prOptionValue",
        "RCBBBB",
        "prOptionValue",
        "P/C,forward,strike,sd,df",
        "1",
        AddinName,
        "",
        "",
        "Returns the PV premium of a Black 76 option on a future / forward",
        // Help text line (optional)
        "(P)ut or (C)all",
        "Forward",
        "Strike",
        "Standard Deviation (=vol*sqrt(time))",
        "Discount Factor",
        "",
    },
    {
        "prOptionDelta",
        "RCBBBB",
        "prOptionDelta",
        "P/C,forward,strike,sd,df",
        "1",
        AddinName,
        "",
        "",
        "Returns the PV premium of a Black 76 option on a future / forward",
        // Help text line (optional)
        "(P)ut or (C)all",
        "Forward",
        "Strike",
        "Standard Deviation (=vol*sqrt(time))",
        "Discount Factor",
        "",
    },
};
