#ifndef XLLADDIN
#define XLLADDIN


#include <iostream>

// #define NUM_COMMANDS      0
#define NUM_FUNCTIONS        4
#define MAX_EXCEL4_ARGS      30

// Used to register DLL functions
extern char *FunctionExports[NUM_FUNCTIONS][MAX_EXCEL4_ARGS - 1];
// extern char *CommandExports[NUM_COMMANDS][2];

// These are displayed by the Excel add-in manager
extern char *AddinVersionStr;
extern char *AddinName;
extern char *DevAddinName;


bool called_from_paste_fn_dlg(void);

#else 

#pragma message (__FILE__ " has already been compiled")

#endif



