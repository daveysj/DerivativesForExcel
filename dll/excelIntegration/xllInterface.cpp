#include "xllAddIn.h"

#ifndef CPP_XLOPER_H
#include "cpp_xloper.h"
#endif

// -- here we are

// PricingXLL.cpp : Defines the entry point for the DLL application.
//

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
                )
{
    return TRUE;
}


//========================================================================
// Given the order in which the following xll functions are called by the
// add in manager, care is required to ensure that activities are syncronised.
// This is best achieved by placing all the initialisation code in a single
// function and checking in all required places that this initialisation
// has occoured using this global variable.
bool xll_initialised = false;
//========================================================================

#define CLASS_NAME_BUFFER_SIZE    50

typedef struct
{
   BOOL is_paste_fn;
   short low_hwnd;
}
   fnwiz_enum_struct;

//------------------------------------------------------------------------
// The callback function called by Windows for every top-level window
BOOL __stdcall fnwiz_enum_proc(HWND hwnd, fnwiz_enum_struct *p_enum)
{
// Check if the parent window is Excel
   if(LOWORD((DWORD)GetParent(hwnd)) != p_enum->low_hwnd)
      return TRUE; // keep iterating

   char class_name[CLASS_NAME_BUFFER_SIZE + 1];
//   Ensure that class_name is always null terminated
   class_name[CLASS_NAME_BUFFER_SIZE] = 0;

   GetClassNameA(hwnd, class_name, CLASS_NAME_BUFFER_SIZE);

//   Do a case-insensitve comparison for the Paste Function window
//   class name with the Excel version number truncated
   if(_strnicmp(class_name, "bosa_sdm_xl", 11) == 0)
   {
      p_enum->is_paste_fn = TRUE;
      return FALSE;  // Tells Windows to stop iterating
   }
   return TRUE; // Tells Windows to continue iterating
}
//------------------------------------------------------------------------
bool called_from_paste_fn_dlg(void)
{
   xloper hwnd = {0.0, xltypeNil}; // super-safe

   if(Excel4(xlGetHwnd, &hwnd, 0))
//   Can't get Excel's main window handle, so assume not
      return false;

   fnwiz_enum_struct es = {FALSE, hwnd.val.w};
   EnumWindows((WNDENUMPROC)fnwiz_enum_proc, (LPARAM)&es);

   return es.is_paste_fn == TRUE;
}
//========================================================================



void display_register_error(char *fn_name, int XL4_err_num, int err_num)
{
   char temp_buffer[256];
   sprintf(temp_buffer, "Could not register function %s (XL4:%d,Err:%d)", fn_name, XL4_err_num, err_num);

   cpp_xloper xStr(temp_buffer);
   cpp_xloper xInt(2); // Dialog box type.

   Excel4(xlcAlert, NULL, 2, &xStr, &xInt);
}


xloper *register_function(int fn_index)
{
   xloper *ptr_array[MAX_EXCEL4_ARGS];

//-------------------------------------------------------
// default to this value in case of a problem
//-------------------------------------------------------
   cpp_xloper RetVal((WORD)xlerrValue);

//-------------------------------------------------------
// Get the full path and name of the DLL.
// Passed as the first argument to xlfRegister, so need
// to set first pointer in array to point to this.
//-------------------------------------------------------
   cpp_xloper DllName;

   if(Excel4(xlGetName, &DllName, 0) != xlretSuccess)
      return NULL;

// Tell destructor to use Excel to free memory
   DllName.SetExceltoFree();
   ptr_array[0] = &DllName;

//-------------------------------------------------------
// Set up the rest of the array of pointers.
//-------------------------------------------------------
   cpp_xloper *fn_args = new cpp_xloper[MAX_EXCEL4_ARGS];

   char *p_arg;
   int i = 0, num_args = 1;

   do
   {
      // get the next string from the char * array
      if((p_arg = FunctionExports[fn_index][i]) == NULL)
         break;

      // Set the corresponding xlfRegister argument
      fn_args[i] = p_arg; // create 
      ptr_array[num_args++] = &(fn_args[i++]); // & operater overloaded to return &xloper
   }
   while(num_args < MAX_EXCEL4_ARGS);

   int xl4_retval = Excel4v(xlfRegister, &RetVal, num_args, ptr_array);

   if(xl4_retval != xlretSuccess || RetVal.GetType() == xltypeErr)
   {
      display_register_error(FunctionExports[fn_index][0], xl4_retval, (int)RetVal);
   }

   delete[] fn_args;
   return RetVal.ExtractXloper(false);
}



//========================================================================
// Excel calls this function whenever it starts up or the add in is loaded.
//========================================================================
int __stdcall xlAutoOpen(void)
{
   if(xll_initialised)
      return 1;

// Register the functions
   for(int i = 0 ; i < NUM_FUNCTIONS; i++)
      register_function(i);
//      register_ID[i] = *(register_function(i));

//   for(i = 0 ; i < NUM_COMMANDS; i++)
//      register_command(CommandExports[i][0], CommandExports[i][1]);

// Setup any custom menus, event traps, etc.
//   customise_Excel_UI(true);

// Initialise the COM interface and the Excel IDispatch object
//   InitExcelOLE();

// Initialise the background thread in the TaskList class
//   ExampleTaskList.CreateTaskThread();

   xll_initialised = true;
   return 1;
}


//========================================================================
bool unregister_function(int fn_index)
{
// Decrement the usage count for the function
//   int xl4 = Excel4(xlfUnregister, 0, 1, register_ID + fn_index);

// Get the name that Excel associates with the function
   cpp_xloper xStr(FunctionExports[fn_index][2]);

// Undo the association of the name with the resource
   int xl4 = Excel4(xlfSetName, 0 , 1, &xStr);

   if(xl4 != xlretSuccess)
      return false;

   return true;
}
//========================================================================

//========================================================================
int __stdcall xlAutoClose(void)
{
   if(!xll_initialised)
      return 1;

// Unregister the functions and commands
   for(int i = 0 ; i < NUM_FUNCTIONS; i++)
      unregister_function(i);

//   for(i = 0 ; i < NUM_COMMANDS; i++)
//      unregister_command(i);

// Remove the custom menus, event traps, etc.
//   customise_Excel_UI(false);

// Un-initialse the Com interface and free the associated objects
//   UninitExcelOLE();

// Shut down the background thread in the TaskList class
//   ExampleTaskList.DeleteTaskThread();

// Delete all DLL-internal names including those created
// via instances of the xlName class
//   delete_all_xll_names();

   xll_initialised = false;
   return 1;
}

//========================================================================
// Excel calls this function the first time the add-in manager is invoked
// It should return an xloper string with the full names of the add-in which
// is then displayed in the add-in manager
xloper * __stdcall xlAddInManagerInfo(xloper *p_arg)
{
   if(!xll_initialised)
      xlAutoOpen();

   cpp_xloper Arg(p_arg);
   cpp_xloper RetVal;

   if(Arg == 1)
      RetVal = AddinName;
   else
      RetVal = (WORD)xlerrValue;

   return RetVal.ExtractXloper();
}



//========================================================================
// Whenever Excel has been returned a pointer to an xloper by the dll with
// the xlbitDLLFree bit set it calls this function passing back the same pointer. 
// This allows the dll to release any dynamically allocated memory that
// was associated with the xloper
void __stdcall xlAutoFree(xloper *p_oper)
{
   if(p_oper->xltype & xltypeMulti)
   {
// Check if the elements need to be freed then check if the array
// itself needs to be freed.
      int size = p_oper->val.array.rows * p_oper->val.array.columns;
      xloper *p = p_oper->val.array.lparray;

      for(; size-- > 0; p++)
         if(p->xltype & (xlbitDLLFree | xlbitXLFree))
            xlAutoFree(p);

      if(p_oper->xltype & xlbitDLLFree)
         free(p_oper->val.array.lparray);
   }
   else if(p_oper->xltype == (xltypeStr | xlbitDLLFree))
   {
      free(p_oper->val.str);
   }
   else if(p_oper->xltype == (xltypeRef | xlbitDLLFree))
   {
      free(p_oper->val.mref.lpmref);
   }
   else if(p_oper->xltype | xlbitXLFree)
   {
      Excel4(xlFree, 0, 1, p_oper);
   }
}


