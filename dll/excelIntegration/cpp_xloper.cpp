//#include "stdafx.h"

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
// This source file (and header file) contain an example class that wraps the
// xloper to simplify the tasks of creating, reading, modifying and managing
// the memory associated with them.  The class code relies on a number of
// routines in the xloper.cpp source file.
//============================================================================
//============================================================================
#include <windows.h>
#include <stdio.h>
#include "cpp_xloper.h"

//===================================================================
//===================================================================
//
// Public functions for class cpp_xloper
//
//===================================================================
//===================================================================

//-------------------------------------------------------------------
// The default constructor.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper()
{
   Clear();
}
//-------------------------------------------------------------------
// The general xloper type constructor.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(xloper *p_op)
{
   Clear();
// Need to ensure that the memory does not get freed twice so
// reset flags in the passed in xloper.  (Note: This does not
// work if the xloper in question is itself contained in a
// cpp_xloper)

// Use the boolean properties to record whether Excel or the DLL
// should be freeing memory.

   if(p_op->xltype & xlbitXLFree)
   {
      p_op->xltype &= ~(xlbitXLFree | xlbitDLLFree);
      m_XLtoFree = true;
   }
   else if(p_op->xltype & xlbitDLLFree)
   {
      p_op->xltype &= ~(xlbitXLFree | xlbitDLLFree);
      m_DLLtoFree = true;
   }
// Copy the xloper and the pointers it contains (i.e. do not copy
// the memory the xloper might refer to, just make a shallow copy).
   m_Op = *p_op;
}
//-------------------------------------------------------------------
// The Variant constructor.  Takes type from the VARTYPE.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(VARIANT *pv)
{
   Clear();

   if(vt_to_xloper(m_Op, pv, true) && (m_Op.xltype == xltypeStr || m_Op.xltype == xltypeMulti))
      m_DLLtoFree = true;
}
//-------------------------------------------------------------------
// The string constructor.  Assumes null-terminated input.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(char *text)
{
   Clear();
   set_to_text(&m_Op, text);
   m_DLLtoFree = true;
}
//-------------------------------------------------------------------
// xltypeInt constructors.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(int w)
{
   Clear();
   set_to_int(&m_Op, w);
}
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(int w, int min, int max)
{
   Clear();
   if(w >= min && w <= max)
   {
      set_to_int(&m_Op, w);
   }
}
//-------------------------------------------------------------------
// xltypeNum (double) constructor.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(double d)
{
   Clear();
   set_to_double(&m_Op, d);
}
//-------------------------------------------------------------------
// xltypeBool constructor.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(bool b)
{
   Clear();
   set_to_bool(&m_Op, b);
}
//-------------------------------------------------------------------
// xltypeErr constructor.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(WORD e)
{
   Clear();
   set_to_err(&m_Op, e);
}
//-------------------------------------------------------------------
// The xltypeSRef constructor.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast)
{
   Clear();
   set_to_xltypeSRef(&m_Op, rwFirst, rwLast, colFirst, colLast);
}
//-------------------------------------------------------------------
// The xltypeRef constructors.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(char *sheet_name, WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast)
{
   Clear();

// Check the inputs.  No need to check sheet_name, as
// creation of cpp_xloper Name will set to xltypeNil (was xltypeMissing)
// if sheet_name is not a valid name.
   if(rwFirst > rwLast || colFirst > colLast)
      return;

// Get the sheetID corresponding to the sheet name provided. If
// sheet_name was NULL, a reference on the active sheet is created.
   cpp_xloper Name(sheet_name);
   cpp_xloper RetOper;

   int xl4 = Excel4(xlSheetId, &RetOper, 1, &Name);

   RetOper.SetExceltoFree();

   if(xl4 == xlretSuccess
   && set_to_xltypeRef(&m_Op, RetOper.m_Op.val.mref.idSheet, rwFirst, rwLast, colFirst, colLast))
   {
   // created successfully
      m_DLLtoFree = true;
   }
   return;
}
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(DWORD SheetID, WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast)
{
   Clear();

   if(rwFirst <= rwLast && colFirst <= colLast
   && set_to_xltypeRef(&m_Op, SheetID, rwFirst, rwLast, colFirst, colLast))
   {
   // created successfully
      m_DLLtoFree = true;
   }
   return;
}
//-------------------------------------------------------------------
// The xltypeMulti constructors.
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(WORD rows, WORD cols)
{
   Clear();
   xloper *p_oper;

   if(!set_to_xltypeMulti(&m_Op, rows, cols)
   || !(p_oper = m_Op.val.array.lparray))
      return;

   m_DLLtoFree = true;

   for(int i = rows * cols; i--; p_oper++)
      p_oper->xltype = xltypeNil; // a safe default
//      p_oper->xltype = xltypeMissing; // a safe default
}
//-------------------------------------------------------------------
void cpp_xloper::InitialiseArray(WORD rows, WORD cols, double *init_data)
{
   Free();
   
   xloper *p_oper;

   if(!init_data || !set_to_xltypeMulti(&m_Op, rows, cols)
   || !(p_oper = m_Op.val.array.lparray))
      return;

   m_DLLtoFree = true;

   for(int i = rows * cols; i--; p_oper++)
   {
      p_oper->xltype = xltypeNum;
      p_oper->val.num = *init_data++;
   }
}
//-------------------------------------------------------------------
void cpp_xloper::InitialiseArray(WORD rows, WORD cols, cpp_xloper *init_array)
{
   Free();
   
   xloper *p_oper;
   cpp_xloper *p_cpp_oper = init_array;

   if(!init_array || !set_to_xltypeMulti(&m_Op, rows, cols)
   || !(p_oper = m_Op.val.array.lparray))
      return;

   m_DLLtoFree = true;
   WORD type;

   for(int i = rows * cols; i--; p_oper++, p_cpp_oper++)
   {
      type = p_cpp_oper->m_Op.xltype;

      if(type != xltypeMulti) // don't permit arrays of arrays
      {
         *p_oper = p_cpp_oper->m_Op; // copy the byte contents
         clone_xloper(p_oper); // make copies of strings and xltypeRef memory
      }
   }
}
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(WORD rows, WORD cols, double *d_array)
{
   Clear();
   InitialiseArray(rows, cols, d_array);
}
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(WORD rows, WORD cols, cpp_xloper *cpp_xloper_array)
{
   Clear();
   InitialiseArray(rows, cols, cpp_xloper_array);
}
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(WORD rows, WORD cols, char **str_array)
{
   Clear();
   xloper *p_oper;

   if(!str_array || !set_to_xltypeMulti(&m_Op, rows, cols)
   || !(p_oper = m_Op.val.array.lparray))
      return;

   m_DLLtoFree = true; // Class will free the strings too
   char *p;

   for(int i = rows * cols; i--; p_oper++)
   {
      p = new_xlstring(*str_array++); // make copies

      if(p)
      {
         p_oper->xltype = xltypeStr;
         p_oper->val.str = p;
      }
      else
      {
         p_oper->xltype = xltypeNil;
//         p_oper->xltype = xltypeMissing;
      }
   }
}
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(WORD &rows, WORD &cols, xloper *input_oper)
{
   Clear();

// Ask Excel to convert the reference to an array (xltypeMulti)
   if(!coerce_xloper(input_oper, m_Op, xltypeMulti))
   {
      rows = cols = 0;
   }
   else
   {
      rows = m_Op.val.array.rows;
      cols = m_Op.val.array.columns;
// Ensure destructor will tell Excel to free memory
      m_XLtoFree = true;
   }
}
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(xl_array *array)
{
   Clear();
   if(array)
      InitialiseArray(array->rows, array->columns, array->array);
}
//-------------------------------------------------------------------
// the copy constructor
//-------------------------------------------------------------------
cpp_xloper::cpp_xloper(cpp_xloper &source)
  :   m_Op(source.m_Op), // make a shallow copy
   m_RowByRowArray(source.m_RowByRowArray)
{
   m_XLtoFree = false;
   m_DLLtoFree = clone_xloper(&m_Op); // convert into a deep copy
}
//-------------------------------------------------------------------
// The destructor.
//-------------------------------------------------------------------
cpp_xloper::~cpp_xloper()
{
   Free();
}
//-------------------------------------------------------------------
cpp_xloper & cpp_xloper::operator=(const cpp_xloper &source)
{
   if(&source == this)
      return *this;

   m_RowByRowArray = true;
   m_XLtoFree = false;
   m_Op = source.m_Op; // make a shallow copy
   m_DLLtoFree = clone_xloper(&m_Op); // convert it into a deep copy
   return *this;
}
//-------------------------------------------------------------------
void cpp_xloper::operator=(int w)
{
   Free();
   set_to_int(&m_Op, w);
}
//-------------------------------------------------------------------
void cpp_xloper::operator=(bool b)
{
   Free();
   set_to_bool(&m_Op, b);
}
//-------------------------------------------------------------------
void cpp_xloper::operator=(double d)
{
   Free();
   set_to_double(&m_Op, d);
}
//-------------------------------------------------------------------
void cpp_xloper::operator=(WORD e)
{
   Free();
   set_to_err(&m_Op, e);
}
//-------------------------------------------------------------------
void cpp_xloper::operator=(char *text)
{
   Free();
   set_to_text(&m_Op, text);
   m_DLLtoFree = true;
}
//-------------------------------------------------------------------
void cpp_xloper::operator=(xloper *p_op)
{
   Free();

   if(!p_op)
      return; // Safe to assign a NULL pointer

   m_Op = *p_op; // Make a shallow copy

   if(p_op->xltype & xlbitDLLFree)
   {
      m_DLLtoFree = true;
      p_op->xltype &= ~xlbitDLLFree;
   }
   else if(p_op->xltype & xlbitXLFree)
   {
      m_XLtoFree = true;
      p_op->xltype &= ~xlbitXLFree;
   }
}
//-------------------------------------------------------------------
void cpp_xloper::operator=(VARIANT *pv)
{
   Free();

   if(vt_to_xloper(m_Op, pv, false) && m_Op.xltype == xltypeStr)
      m_DLLtoFree = true;
}
//-------------------------------------------------------------------
void cpp_xloper::operator=(xl_array *array)
{
   Free();
   if(array)
      InitialiseArray(array->rows, array->columns, array->array);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator==(cpp_xloper &cpp_op2)
{
   return xlopers_equal(m_Op, cpp_op2.m_Op);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator==(xloper *p_op)
{
   return xlopers_equal(m_Op, *p_op);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator==(int w)
{
   int i;

   if(!coerce_to_int(&m_Op, i) || i != w)
      return false;

   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::operator==(bool b)
{
   bool a;

   if(!coerce_to_bool(&m_Op, a) || a != b)
      return false;

   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::operator==(double d)
{
   double dd;

   if(!coerce_to_double(&m_Op, dd) || dd != d)
      return false;

   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::operator==(WORD e)
{
   int i;

   if(!coerce_to_int(&m_Op, i) || (WORD)i != e)
      return false;

   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::operator==(char *text)
{
   if(m_Op.xltype != xltypeStr)
      return false;

   int i;
   char *p = m_Op.val.str + 1;

   for(i = m_Op.val.str[0]; i--;)
      if(*p++ != *text++)
         return false;

   if(*text)
      return false; // text is longer than m_Op.val.str

   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::operator!=(int w)
{
   return !operator==(w);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator!=(bool b)
{
   return !operator==(b);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator!=(double d)
{
   return !operator==(d);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator!=(WORD e)
{
   return !operator==(e);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator!=(char *text)
{
   return !operator==(text);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator!=(cpp_xloper &cpp_op2)
{
   return !xlopers_equal(m_Op, cpp_op2.m_Op);
}
//-------------------------------------------------------------------
bool cpp_xloper::operator!=(xloper *p_op)
{
   return !xlopers_equal(m_Op, *p_op);
}
//-------------------------------------------------------------------
void cpp_xloper::operator++(void)
{
   switch(m_Op.xltype)
   {
   case xltypeNum:      m_Op.val.num += 1.0;   break;
   case xltypeInt:      m_Op.val.w++;         break;

   case xltypeSRef:
   case xltypeRef:
      xloper ValueType;
      ValueType.xltype = xltypeInt;
      ValueType.val.w = xltypeNum;
      xloper Value;
      if(Excel4(xlCoerce, &Value, 2, &m_Op, &ValueType) || Value.xltype != xltypeNum)
         break;

      Value.val.num += 1.0;
      Excel4(xlSet, 0, 2, &m_Op, &Value);
      break;
   }
}
//-------------------------------------------------------------------
void cpp_xloper::operator--(void)
{
   switch(m_Op.xltype)
   {
   case xltypeNum:      m_Op.val.num -= 1.0;   break;
   case xltypeInt:      m_Op.val.w--;         break;

   case xltypeSRef:
   case xltypeRef:
      xloper ValueType;
      ValueType.xltype = xltypeInt;
      ValueType.val.w = xltypeNum;
      xloper Value;
      if(Excel4(xlCoerce, &Value, 2, &m_Op, &ValueType) || Value.xltype != xltypeNum)
         break;

      Value.val.num -= 1.0;
      Excel4(xlSet, 0, 2, &m_Op, &Value);
      break;
   }
}
//-------------------------------------------------------------------
// Returns the type of the xloper
//-------------------------------------------------------------------
int cpp_xloper::GetType(void)
{
   return m_Op.xltype;
}
//-------------------------------------------------------------------
void cpp_xloper::SetType(int new_type)
{
   Free();
   set_to_type(&m_Op, new_type);
}
//-------------------------------------------------------------------
bool cpp_xloper::SetTypeMulti(WORD array_rows, WORD array_cols)
{
   DWORD array_size = array_rows * array_cols;

   if(!array_size || array_cols > 256)
      return false;

   Free();
   set_to_type(&m_Op, xltypeMulti);
   m_Op.val.array.lparray = (xloper *)calloc(sizeof(xloper), array_size);
   m_Op.val.array.columns = array_cols;
   m_Op.val.array.rows = array_rows;

   xloper *p = m_Op.val.array.lparray;

   for(;array_size--; p++)
      p->xltype = xltypeNil; // equivalent to an empty cell

   return m_DLLtoFree = true;
}
//-------------------------------------------------------------------
// Change the type to an xltypeSref referencing the given cells
//-------------------------------------------------------------------
bool cpp_xloper::SetCell(WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast)
{
   Free();

   return set_to_xltypeSRef(&m_Op, rwFirst, rwLast, colFirst, colLast);
}
//-------------------------------------------------------------------
// Change the type and assign the given value
//-------------------------------------------------------------------
cpp_xloper::operator int(void)
{
   int i;

   if(coerce_to_int(&m_Op, i))
      return i;

   return 0;
}
//-------------------------------------------------------------------
cpp_xloper::operator bool(void)
{
   bool b;

   if(coerce_to_bool(&m_Op, b))
      return b;

   return false;
}
//-------------------------------------------------------------------
cpp_xloper::operator double(void)
{
   double d;

   if(coerce_to_double(&m_Op, d))
      return d;

   return 0.0;
}
//-------------------------------------------------------------------
cpp_xloper::operator char *(void)
{
   char *p;

   if(coerce_to_string(&m_Op, p))
      return p;

   return NULL;
}
//-------------------------------------------------------------------
bool cpp_xloper::AsVariant(VARIANT &var)
{
   return xloper_to_vt(&m_Op, var, true);
}
//-------------------------------------------------------------------
xl_array *cpp_xloper::AsDblArray(void)
{
   xl_array *p_ret_array;
   double *p;
   int size;
   xloper *p_op;

   switch(m_Op.xltype)
   {
   case xltypeMulti:
      size = m_Op.val.array.rows * m_Op.val.array.columns;
      p_ret_array = (xl_array *)malloc(sizeof(xl_array) + sizeof(double) * (size - 1));
      p_ret_array->columns = m_Op.val.array.columns;
      p_ret_array->rows = m_Op.val.array.rows;

      if(!p_ret_array)
         return NULL;

// Get the cell values one-by-one as doubles and place in the array.
// Store the array row-by-row in memory.
      if(!(p_op = m_Op.val.array.lparray))
      {
         free(p_ret_array);
         return NULL;
      }

      for(p = p_ret_array->array; size--; p++)
         if(!coerce_to_double(p_op++, *p))
            *p = 0.0;

      return p_ret_array;

   case xltypeNum:
   case xltypeStr:
   case xltypeBool:
      p_ret_array = (xl_array *)malloc(sizeof(xl_array));
      p_ret_array->columns = p_ret_array->rows = 1;
      if(!coerce_to_double(&m_Op, p_ret_array->array[0]))
         p_ret_array->array[0] = 0.0;

      return p_ret_array;
   }
   return NULL;
}
//-------------------------------------------------------------------
bool cpp_xloper::GetVal(WORD &e)
{
   if(m_Op.xltype == xltypeErr)
   {
      e = m_Op.val.err;
      return true;
   }
   e = 0;
   return false;
}
//-------------------------------------------------------------------
// Array access functions with row and column parameters
//-------------------------------------------------------------------
int cpp_xloper::GetArrayElementType(WORD row, WORD column)
{
   xloper *p_op = GetArrayElement(row, column);

   if(p_op)
      return p_op->xltype;

   return 0;
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArraySize(WORD &rows, WORD &cols)
{
   if(m_Op.xltype != xltypeMulti)
   {
      rows = cols = 0;
      return false;
   }
   rows = m_Op.val.array.rows;
   cols = m_Op.val.array.columns;
   return true;
}
//-------------------------------------------------------------------
xloper *cpp_xloper::GetArrayElement(WORD row, WORD column)
{
   if(m_Op.xltype != xltypeMulti)
      return NULL;

   DWORD offset;
   WORD rows = m_Op.val.array.rows;
   WORD cols = m_Op.val.array.columns;

   if(row >= rows || column >= cols)
      return NULL;

   if(m_RowByRowArray)
      offset = cols * row + column;
   else
      offset = column * rows + row;

   return m_Op.val.array.lparray + offset;
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(WORD row, WORD column, int &w)
{
   return coerce_to_int(GetArrayElement(row, column), w);
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(WORD row, WORD column, bool &b)
{
   return coerce_to_bool(GetArrayElement(row, column), b);
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(WORD row, WORD column, double &d)
{
   return coerce_to_double(GetArrayElement(row, column), d);
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(WORD row, WORD column, WORD &e)
{
   xloper *p_op = GetArrayElement(row, column);

   if(p_op && p_op->xltype == xltypeErr)
   {
      e = p_op->val.err;
      return true;
   }
   e = 0;
   return false;
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(WORD row, WORD column, char *&text)
{
   text = 0;

   xloper *p_op = GetArrayElement(row, column);

   if(!p_op)
      return false;

   return coerce_to_string(p_op, text);
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(WORD row, WORD column, int w)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   set_to_int(GetArrayElement(row, column), w);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(WORD row, WORD column, bool b)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   set_to_bool(GetArrayElement(row, column), b);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(WORD row, WORD column, double d)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(row, column);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   set_to_double(p_op, d);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(WORD row, WORD column, WORD e)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(row, column);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   set_to_err(p_op, e);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(WORD row, WORD column, char *text)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(row, column);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   set_to_text(p_op, text);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElementType(WORD row, WORD column, int new_type)
{
   if(m_XLtoFree)
      return false;  // Don't change an Excel-allocated array

   switch(new_type)
   {
   case xltypeNum:
   case xltypeStr:
   case xltypeBool:
   case xltypeErr:
   case xltypeMissing:
   case xltypeNil:
   case xltypeInt:
      break;

   default:
      return false; // nothing else permitted for arrays
   }

   xloper *p_op = GetArrayElement(row, column);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   return set_to_type(p_op, new_type);
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(WORD row, WORD column, xloper *p_source)
{
   if(m_XLtoFree || !p_source)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(row, column);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }

   if(p_source->xltype == xltypeMulti)
   {
// can't assign an array to array element so
// just assign the top left value if array
// has memory assigned
      if(p_source->val.array.lparray)
         *p_op = p_source->val.array.lparray[0];
      else
         p_op->xltype = xltypeNil;
//         p_op->xltype = xltypeMissing;
   }
   else
   {
      *p_op = *p_source;
      clone_xloper(p_op);
   }
   return true;
}
//-------------------------------------------------------------------
// Array access functions with offset parameter
//-------------------------------------------------------------------
int cpp_xloper::GetArrayElementType(DWORD offset)
{
   xloper *p_op = GetArrayElement(offset);

   if(p_op)
      return p_op->xltype;

   return 0;
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArraySize(DWORD &size)
{
   if(m_Op.xltype != xltypeMulti)
   {
      size = 0;
      return false;
   }
   size = m_Op.val.array.rows * m_Op.val.array.columns;
   return true;
}
//-------------------------------------------------------------------
xloper *cpp_xloper::GetArrayElement(DWORD offset)
{
   if(m_Op.xltype != xltypeMulti)
      return NULL;

   DWORD size = m_Op.val.array.rows * m_Op.val.array.columns;

   if(offset >= size)
      return NULL;

   return m_Op.val.array.lparray + offset;
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(DWORD offset, char *&text)
{
   return coerce_to_string(GetArrayElement(offset), text);
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(DWORD offset, double &d)
{
   return coerce_to_double(GetArrayElement(offset), d);
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(DWORD offset, int &w)
{
   return coerce_to_int(GetArrayElement(offset), w);
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(DWORD offset, bool &b)
{
   return coerce_to_bool(GetArrayElement(offset), b);
}
//-------------------------------------------------------------------
bool cpp_xloper::GetArrayElement(DWORD offset, WORD &e)
{
   xloper *p_op = GetArrayElement(offset);

   if(!p_op || p_op->xltype != xltypeErr)
      return false;

   e = p_op->val.err;
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(DWORD offset, char *text)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(offset);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   set_to_text(p_op, text);
   return true;
}

//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(DWORD offset, double d)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(offset);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   set_to_double(p_op, d);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(DWORD offset, int w)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(offset);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   set_to_int(p_op, w);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(DWORD offset, bool b)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(offset);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   set_to_bool(p_op, b);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(DWORD offset, WORD e)
{
   if(m_XLtoFree)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(offset);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   set_to_err(p_op, e);
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElement(DWORD offset, xloper *p_source)
{
   if(m_XLtoFree || !p_source)
      return false;  // Should not assign to an Excel-allocated array

   xloper *p_op = GetArrayElement(offset);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }

   if(p_source->xltype == xltypeMulti)
   {
// can't assign an array to array element so
// just assign the top left value if array
// has memory assigned
      if(p_source->val.array.lparray)
         *p_op = p_source->val.array.lparray[0];
      else
         p_op->xltype = xltypeNil;
//         p_op->xltype = xltypeMissing;
   }
   else
   {
      *p_op = *p_source;
      clone_xloper(p_op);
   }
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::SetArrayElementType(DWORD offset, int new_type)
{
   if(m_XLtoFree)
      return false;  // Don't change an Excel-allocated array

   switch(new_type)
   {
   case xltypeNum:
   case xltypeStr:
   case xltypeBool:
   case xltypeErr:
   case xltypeMissing:
   case xltypeNil:
   case xltypeInt:
      break;

   default:
      return false; // nothing else permitted for arrays
   }

   xloper *p_op = GetArrayElement(offset);

   if(!p_op)
      return false;
      
   if(m_DLLtoFree)
   {
      free_xloper(p_op, false);
   }
   return set_to_type(p_op, new_type);
}
//-------------------------------------------------------------------
bool cpp_xloper::Transpose(void)
{
   if(!IsType(xltypeMulti))
      return false;

   WORD r, c, rows, columns;

   GetArraySize(rows, columns);

   xloper *new_array = (xloper *)malloc(sizeof(xloper) * rows * columns);
   xloper *p_source = m_Op.val.array.lparray;
   int new_index;

   for(r = 0; r < rows; r++)
   {
      new_index = r;

      for(c = 0; c < columns; c++, new_index += rows)
      {
         new_array[new_index] = *p_source++;
      }
   }

   r = m_Op.val.array.columns;
   m_Op.val.array.columns = m_Op.val.array.rows;
   m_Op.val.array.rows = r;

   memcpy(m_Op.val.array.lparray, new_array, sizeof(xloper) * rows * columns);
   free(new_array);
   return true;
}
//-------------------------------------------------------------------
double *cpp_xloper::ConvertMultiToDouble(void)
{
   if(m_Op.xltype != xltypeMulti)
      return NULL;

// Allocate the space for the array of doubles
   int size = m_Op.val.array.rows * m_Op.val.array.columns;
   double *ret_array = (double *)malloc(size * sizeof(double));

   if(!ret_array)
      return NULL;

// Get the cell values one-by-one as doubles and place in the array.
// Store the array row-by-row in memory.
   xloper *p_op = m_Op.val.array.lparray;

   if(!p_op)
   {
      free(ret_array);
      return NULL;
   }

   double *p = ret_array;

   for(; size--; p++)
      if(!coerce_to_double(p_op++, *p))
         *p = 0.0;

   return ret_array; // caller must free the memory!
}
//-------------------------------------------------------------------
//-------------------------------------------------------------------
xloper *cpp_xloper::ExtractXloper(bool ExceltoFree)
{
   static xloper ret_val;

   ret_val = m_Op;

   if(ExceltoFree || m_XLtoFree)
   {
      ret_val.xltype |= xlbitXLFree;
   }
   else if(m_DLLtoFree)
   {
      ret_val.xltype |= xlbitDLLFree;

      if(m_Op.xltype & xltypeMulti)
      {
         int limit = m_Op.val.array.rows * m_Op.val.array.columns;
         xloper *p = m_Op.val.array.lparray;

         for(int i = limit; i--; p++)
            if(p->xltype == xltypeStr || p->xltype == xltypeRef)
               p->xltype |= xlbitDLLFree;
      }
   }
// Prevent the destructor from freeing memory by resetting properties
   Clear();
   return &ret_val;
}
//-------------------------------------------------------------------
void cpp_xloper::SetExceltoFree(void)
{
   m_DLLtoFree = false;
   m_XLtoFree = true;
}
//-------------------------------------------------------------------
void cpp_xloper::SetDlltoFree(void)
{
   m_DLLtoFree = true;
   m_XLtoFree = false;
}
//-------------------------------------------------------------------
bool cpp_xloper::ConvertToString(bool ExceltoFree)
{
   if(m_Op.xltype == xltypeStr)
   {
      m_DLLtoFree = !(m_XLtoFree = ExceltoFree);
      return true;
   }

   char *text;

   if(coerce_to_string(&m_Op, text) == false)
      return false;

   if(ExceltoFree)
      m_XLtoFree = true;

   Free();
   set_to_text(&m_Op, text);
   m_DLLtoFree = true;
   return true;
}
//-------------------------------------------------------------------
bool cpp_xloper::IsType(int type)
{
   int op_type = m_Op.xltype & ~(xlbitDLLFree | xlbitXLFree);

   if((op_type & xltypeBigData) == xltypeBigData)
      return false;  // cannot check in this way for this type

   return (op_type & type) != 0;
}
//-------------------------------------------------------------------
bool cpp_xloper::IsBigData(void)
{
   int op_type = m_Op.xltype & ~(xlbitDLLFree | xlbitXLFree);

   if((op_type & xltypeBigData) == xltypeBigData)
      return true;

   return false;
}
//-------------------------------------------------------------------
// Resets cpp_oper properties without freeing up memory 
//-------------------------------------------------------------------
void cpp_xloper::Clear(void)
{
   m_XLtoFree = m_DLLtoFree = false;
   m_RowByRowArray = true;
   m_Op.xltype = xltypeNil;
}
//===================================================================
//===================================================================
// End of public functions
//===================================================================
//===================================================================
// Private functions 
//===================================================================
//===================================================================
// Free up memory and reset cpp_oper properties
//===================================================================
//===================================================================
void cpp_xloper::Free(bool ExceltoFree)  // free memory and initialise
{
   if(ExceltoFree)
      m_XLtoFree = true;

   if(m_XLtoFree)
   {
      Excel4(xlFree, 0, 1, &m_Op);
   }
   else if(m_DLLtoFree)
   {
      free_xloper(&m_Op, false);
   }
// Reset the properties ready for destruction or reuse
   Clear();
}
