//============================================================================
//============================================================================
//
// Excel Add-in Development in C/C++, Applications in Finance
// 
// Author: Steve Dalton
// 
// Published by John Wiley & Sons Ltd, The Atrium, Southern Gate, Chichester,
// West Sussex, PO19 8SQ, UK.
// 
// All intellectual property rights and copyright in the code listed in this
// module are reserved.  You may reproduce and use this code only as permitted
// in the Rules of Use that you agreed to on installation.  To use the material
// in this source module, or material derived from it, in any other way you
// must first obtain written permission.  Requests for permission can be sent
// by email to permissions@eigensys.com.
// 
// No warranty, explicit or implied, is made by either the author or publisher
// as to the quality, fitness for a particular purpose, accuracy or
// appropriateness of material in this module.  The code and methods contained
// are intended soley for example and clarification.  You should not rely on
// any part of this code without having completely satisfied yourself that it
// is correct and appropriate for your needs.
//
//============================================================================
//============================================================================
#ifndef _CPP_XLOPER_H
#define _CPP_XLOPER_H
//#pragma message ("_CPP_XLOPER_H defined")

#ifndef _XLCALL_H
#include <windows.h>
#include "xlcall.h"
#endif

#ifndef _XLOPER_H
#include "xloper.h"
#endif

class cpp_xloper
{
public:
//--------------------------------------------------------------------
// constructors
//--------------------------------------------------------------------
   cpp_xloper();            // created as xltypeMissing
   cpp_xloper(xloper *p_oper);   // contains copy of given xloper
   cpp_xloper(char *text);      // xltypeStr
   cpp_xloper(int w);         // xltypeInt
   cpp_xloper(int w, int min, int max); // xltypeInt (or xltypeMissing)
   cpp_xloper(double d);      // xltypeNum
   cpp_xloper(bool b);         // xltypeBool
   cpp_xloper(WORD e);         // xltypeErr
   cpp_xloper(WORD, WORD, BYTE, BYTE);   // xltypeSRef
   cpp_xloper(char *, WORD, WORD, BYTE, BYTE); // xltypeRef from sheet name
   cpp_xloper(DWORD, WORD, WORD, BYTE, BYTE);  // xltypeRef from sheet ID
   cpp_xloper(VARIANT *pv);   // Takes its type from the VARTYPE

   // xltypeMulti constructors
   cpp_xloper(WORD rows, WORD cols); // array of undetermined type
   cpp_xloper(WORD rows, WORD cols, double *d_array); // array of xltypeNum
   cpp_xloper(WORD rows, WORD cols, char **str_array); // array of xltypeStr
   cpp_xloper(WORD &rows, WORD &cols, xloper *input_oper); // from SRef/Ref
   cpp_xloper(WORD rows, WORD cols, cpp_xloper *init_array);
   cpp_xloper(xl_array *array);

   cpp_xloper(cpp_xloper &source); // Copy constructor

//--------------------------------------------------------------------
// destructor
//--------------------------------------------------------------------
   ~cpp_xloper();

//--------------------------------------------------------------------
// Overloaded operators
//--------------------------------------------------------------------
   cpp_xloper &operator=(const cpp_xloper &source);
   void operator=(int);      // xltypeInt
   void operator=(bool b);      // xltypeBool
   void operator=(double);      // xltypeNum
   void operator=(WORD e);      // xltypeErr
   void operator=(char *);      // xltypeStr
   void operator=(xloper *);   // same type as passed-in xloper
   void operator=(VARIANT *);   // same type as passed-in Variant
   void operator=(xl_array *array);
   bool operator==(cpp_xloper &cpp_op2);
   bool operator==(int w);
   bool operator==(bool b);
   bool operator==(double d);
   bool operator==(WORD e);
   bool operator==(char *text);
   bool operator==(xloper *);
   bool operator!=(cpp_xloper &cpp_op2);
   bool operator!=(int w);
   bool operator!=(bool b);
   bool operator!=(double d);
   bool operator!=(WORD e);
   bool operator!=(char *text);
   bool operator!=(xloper *);
   void operator++(void);
   void operator--(void);
   operator int(void);
   operator bool(void);
   operator double(void);
   operator char *(void);
   xloper *operator&() {return &m_Op;}   // return xloper address

//--------------------------------------------------------------------
// property get and set functions
//--------------------------------------------------------------------
   int  GetType(void);
   void SetType(int new_type);
   bool SetTypeMulti(WORD array_rows, WORD array_cols);
   bool SetCell(WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast);
   bool GetVal(WORD &e);
   bool IsType(int);
   bool IsStr(void)   {return IsType(xltypeStr);}
   bool IsNum(void)   {return IsType(xltypeNum);}
   bool IsBool(void)   {return IsType(xltypeBool);}
   bool IsInt(void)   {return IsType(xltypeInt);}
   bool IsErr(void)   {return IsType(xltypeErr);}
   bool IsMulti(void)   {return IsType(xltypeMulti);}
   bool IsNil(void)   {return IsType(xltypeNil);}
   bool IsMissing(void){return IsType(xltypeMissing);}
   bool IsRef(void)   {return IsType(xltypeRef | xltypeSRef);}
   bool IsBigData(void);

//--------------------------------------------------------------------
// property get and set functions for xltypeMulti
//--------------------------------------------------------------------
   int  GetArrayElementType(WORD row, WORD column);
   bool GetArraySize(WORD &rows, WORD &cols);
   xloper *GetArrayElement(WORD row, WORD column);

   bool GetArrayElement(WORD row, WORD column, int &w);
   bool GetArrayElement(WORD row, WORD column, bool &b);
   bool GetArrayElement(WORD row, WORD column, double &d);
   bool GetArrayElement(WORD row, WORD column, WORD &e);
   bool GetArrayElement(WORD row, WORD column, char *&text); // makes new string

   bool SetArrayElementType(WORD row, WORD column, int new_type);
   bool SetArrayElement(WORD row, WORD column, int w);
   bool SetArrayElement(WORD row, WORD column, bool b);
   bool SetArrayElement(WORD row, WORD column, double d);
   bool SetArrayElement(WORD row, WORD column, WORD e);
   bool SetArrayElement(WORD row, WORD column, char *text);
   bool SetArrayElement(WORD row, WORD column, xloper *p_source);

   int  GetArrayElementType(DWORD offset);
   bool GetArraySize(DWORD &size);
   xloper *GetArrayElement(DWORD offset);

   bool GetArrayElement(DWORD offset, char *&text); // makes new string
   bool GetArrayElement(DWORD offset, double &d);
   bool GetArrayElement(DWORD offset, int &w);
   bool GetArrayElement(DWORD offset, bool &b);
   bool GetArrayElement(DWORD offset, WORD &e);

   bool SetArrayElementType(DWORD offset, int new_type);
   bool SetArrayElement(DWORD offset, int w);
   bool SetArrayElement(DWORD offset, bool b);
   bool SetArrayElement(DWORD offset, double d);
   bool SetArrayElement(DWORD offset, WORD e);
   bool SetArrayElement(DWORD offset, char *text);
   bool SetArrayElement(DWORD offset, xloper *p_source);

   void InitialiseArray(WORD rows, WORD cols, double *init_data);
   void InitialiseArray(WORD rows, WORD cols, cpp_xloper *init_array);
   bool Transpose(void);

   double *ConvertMultiToDouble(void);

//--------------------------------------------------------------------
// other public functions
//--------------------------------------------------------------------
   void Clear(void);   // Clears the xloper without freeing memory
   void SetExceltoFree(void);   // Tell the destructor to use xlFree
   void SetDlltoFree(void);   // Tell the destructor to use free()
   cpp_xloper *Addr(void) {return this;} // Returns address of cpp_xloper
   xloper *ExtractXloper(bool ExceltoFree = false); // extract xloper, clear cpp_xloper
   void Free(bool ExceltoFree = false); // free memory 
   bool ConvertToString(bool ExceltoFree);
   bool AsVariant(VARIANT &var); // Return an equivalent Variant
   xl_array *AsDblArray(void); // Return an xl_array

private:
   xloper m_Op;
   bool m_RowByRowArray;
   bool m_DLLtoFree;
   bool m_XLtoFree;
};


#else
#pragma message ("_CPP_XLOPER_H skipped")
#endif // _CPP_XLOPER_H
