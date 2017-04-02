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
// This source file contains definitions of constant xlopers, functions that
// convert between xlopers and other data types, compare xlopers, clone
// xlopers, convert to and from Variant data types.
//============================================================================
//============================================================================
#include <windows.h>
#include "xloper.h"

// Use this structure to initialise constant xloopers
typedef struct
{
   WORD word1;
   WORD word2;
   WORD word3;
   WORD word4;
   WORD xltype;
}
   const_xloper;

const_xloper xloperBooleanTrue = {1, 0, 0, 0, xltypeBool};
const_xloper xloperBooleanFalse = {0, 0, 0, 0, xltypeBool};
const_xloper xloperMissing = {0, 0, 0, 0, xltypeMissing};
const_xloper xloperNil = {0, 0, 0, 0, xltypeNil};
const_xloper xloperErrNull = {0, 0, 0, 0, xltypeErr};
const_xloper xloperErrDiv0 = {7, 0, 0, 0, xltypeErr};
const_xloper xloperErrValue = {15, 0, 0, 0, xltypeErr};
const_xloper xloperErrRef = {23, 0, 0, 0, xltypeErr};
const_xloper xloperErrName = {29, 0, 0, 0, xltypeErr};
const_xloper xloperErrNum = {36, 0, 0, 0, xltypeErr};
const_xloper xloperErrNa = {42, 0, 0, 0, xltypeErr};

xloper *p_xlTrue = ((xloper *)&xloperBooleanTrue);
xloper *p_xlFalse = ((xloper *)&xloperBooleanFalse);
xloper *p_xlMissing = ((xloper *)&xloperMissing);
xloper *p_xlNil = ((xloper *)&xloperNil);
xloper *p_xlErrNull = ((xloper *)&xloperErrNull);
xloper *p_xlErrDiv0 = ((xloper *)&xloperErrDiv0);
xloper *p_xlErrValue = ((xloper *)&xloperErrValue);
xloper *p_xlErrRef = ((xloper *)&xloperErrRef);
xloper *p_xlErrName = ((xloper *)&xloperErrName);
xloper *p_xlErrNum = ((xloper *)&xloperErrNum);
xloper *p_xlErrNa = ((xloper *)&xloperErrNa);


//--------------------------------------------
// Excel worksheet cell error codes are passed
// via VB OLE Variant arguments in 'ulVal'.
// These are equivalent to the offset below
// plus the value defined in "xlcall.h"
//
// This is easier than using the VT_SCODE variant
// property 'scode'.
//--------------------------------------------
#define VT_XL_ERR_OFFSET      2148141008ul

//===================================================================
//
// Functions to manage xlopers
//
//===================================================================
bool is_xloper_missing(xloper *p_op)
{
   return (!p_op || (p_op->xltype & (xltypeMissing | xltypeNil)) != 0);
}
//===================================================================
// Frees xloper memory using either Excel or free().  If use_xlbits
// is true, then only does this where the appropriate bit is set. If
// use_xlbits is false, then uses free() and assumes that all types
// that have memory were allocated in a way that is compatible with
// freeing by a call to free().  Calling this with use_xlbits false
// on an xloper pointing to static memory will cause an exception.
//===================================================================
void free_xloper(xloper *p_op, bool use_xlbits)
{
// If created by Excel and the DLL has set this bit, then use Excel
// to free the memory.
   if(use_xlbits && (p_op->xltype & xlbitXLFree))
   {
      p_op->xltype &= ~xlbitXLFree;
      Excel4(xlFree, 0, 1, p_op);
      return;
   }

   WORD dll_free = use_xlbits ? xlbitDLLFree : 0;
   WORD xl_free = use_xlbits ? xlbitXLFree : 0;

   if(p_op->xltype & xltypeMulti)
   {
// First check if elements need to be freed then check if the array
// itself needs to be freed.
      int limit = p_op->val.array.rows * p_op->val.array.columns;
      xloper *p = p_op->val.array.lparray;

      for(int i = limit; i--; p++)
         if(p->xltype & (xl_free | dll_free))
            free_xloper(p, use_xlbits);

      if(p_op->xltype & dll_free)
         free(p_op->val.array.lparray);
   }
   else if(p_op->xltype == (xltypeStr | dll_free))
   {
      free(p_op->val.str);
   }
   else if(p_op->xltype == (xltypeRef | dll_free))
   {
      free(p_op->val.mref.lpmref);
   }
}
//-------------------------------------------------------------------
char *copy_xlstring(char *xlstring)
{
   if(!xlstring)
      return NULL;

   int len = xlstring[0] + 2;
   char *p = (char *)malloc(len);
   memcpy(p, xlstring, len);
   return p;
}
//-------------------------------------------------------------------
char *new_xlstring(char *text)
{
   int len;

   if(!text || !(len = strlen(text)))
      return NULL;

   if(len > 255)
      len = 255; // truncate

   char *p = (char *)malloc(len + 2);
   memcpy(p + 1, text, len + 1);
   p[0] = (char)len;
   return p;
}
//-------------------------------------------------------------------
bool coerce_xloper(xloper *p_op, xloper &ret_val, int target_type)
{
// Target will contain the information that tells Excel what type to
// convert to.
   xloper target;

   target.xltype = xltypeInt;
   target.val.w = target_type; // can be more than one type

   if(Excel4(xlCoerce, &ret_val, 2, p_op, &target) != xlretSuccess
   || (ret_val.xltype & target_type) == 0)
      return false;

   return true;
}
//-------------------------------------------------------------------
bool coerce_to_string(xloper *p_op, char *&text)
{
   char *str;
   xloper ret_val;

   text = 0;

   if(!p_op || (p_op->xltype & (xltypeMissing | xltypeNil)) != 0)
      return false;

   if(p_op->xltype != xltypeStr)
   {
// xloper is not a string type, so try to convert it.
      if(!coerce_xloper(p_op, ret_val, xltypeStr))
         return false;

      str = ret_val.val.str;
   }
   else if(!(str = p_op->val.str)) // make a working copy of the ptr
      return false;

   int len = str[0];

   if((text = (char *)malloc(len + 1)) == NULL) // caller must free this
   {
      if(p_op->xltype != xltypeStr)
         Excel4(xlFree, 0, 1, &ret_val);

      return false;
   }

   if(len)
      memcpy(text, str + 1, len);

   text[len] = 0;

// If the string from which the copy was made was created in a call
// to coerce_xloper above, then need to free it with a call to xlFree
   if(p_op->xltype != xltypeStr)
      Excel4(xlFree, 0, 1, &ret_val);
   
   return true;
}
//-------------------------------------------------------------------
bool coerce_to_double(xloper *p_op, double &d)
{
   if(!p_op || (p_op->xltype & (xltypeMissing | xltypeNil)) != 0)
      return false;

   if(p_op->xltype == xltypeNum)
   {
      d = p_op->val.num;
      return true;
   }

// xloper is not a floating point number type, so try to convert it.
   xloper ret_val;

   if(!coerce_xloper(p_op, ret_val, xltypeNum))
      return false;

   d = ret_val.val.num;
   return true;
}
//-------------------------------------------------------------------
// Allocate and populate an array of doubles based on input xloper
//-------------------------------------------------------------------
double *coerce_to_double_array(xloper *p_op, double invalid_value, WORD &cols, WORD &rows)
{
   if(!p_op || (p_op->xltype & (xltypeMissing | xltypeNil)))
      return NULL;

// xloper is not an xloper array type, so try to convert it.
   xloper ret_val;

   if(!coerce_xloper(p_op, ret_val, xltypeMulti))
      return NULL;

// Allocate the space for the array of doubles
   cols = ret_val.val.array.columns;
   rows = ret_val.val.array.rows;
   int size = rows * cols;
   double *d_array = (double *)malloc(size * sizeof(double));

   if(!d_array)
   {
   // Must free array memory allocated by xlCoerce
      Excel4(xlFree, 0, 1, &ret_val);
      return NULL;
   }

// Get the cell values one-by-one as doubles and place in the array.
// Store the array row-by-row in memory.
   xloper *p_elt = ret_val.val.array.lparray;

   if(!p_elt) // array could not be created
   {
   // Must free array memory allocated by xlCoerce
      Excel4(xlFree, 0, 1, &ret_val);
      free(d_array);
      return NULL;
   }

   double *p = d_array;

   for(; size--; p++)
      if(!coerce_to_double(p_elt++, *p))
         *p = invalid_value;

   Excel4(xlFree, 0, 1, &ret_val);
   return d_array; // caller must free this
}
//-------------------------------------------------------------------
bool coerce_to_int(xloper *p_op, int &w)
{
   if(!p_op || (p_op->xltype & (xltypeMissing | xltypeNil)) != 0)
      return false;

   if(p_op->xltype == xltypeInt)
   {
      w = p_op->val.w;
      return true;
   }

   if(p_op->xltype == xltypeErr)
   {
      w = p_op->val.err;
      return true;
   }

// xloper is not an integer type, so try to convert it.
   xloper ret_val;

   if(!coerce_xloper(p_op, ret_val, xltypeInt))
      return false;

   w = ret_val.val.w;
   return true;
}
//-------------------------------------------------------------------
bool coerce_to_short(xloper *p_op, short &s)
{
   int i;
   if(!coerce_to_int(p_op, i))
      return false;

   s = (short)i;
   return true;
}
//-------------------------------------------------------------------
bool coerce_to_bool(xloper *p_op, bool &b)
{
   if(!p_op || (p_op->xltype & (xltypeMissing | xltypeNil)) != 0)
      return false;

   if(p_op->xltype == xltypeBool)
   {
      b = (p_op->val._bool != 0);
      return true;
   }

//   if(p_op->xltype == xltypeErr)
//      return (p_op->val.err != 0);

// xloper is not a Boolean number type, so try to convert it.
   xloper ret_val;

   if(!coerce_xloper(p_op, ret_val, xltypeBool))
      return false;

   b = (ret_val.val._bool != 0);
   return true;
}
//-------------------------------------------------------------------
void set_to_double(xloper *p_op, double d)
{
   if(!p_op) return;
   p_op->xltype = xltypeNum;
   p_op->val.num = d;
}
//-------------------------------------------------------------------
void set_to_int(xloper *p_op, int w)
{
   if(!p_op) return;
   p_op->xltype = xltypeInt;
   p_op->val.w = w;
}
//-------------------------------------------------------------------
void set_to_bool(xloper *p_op, bool b)
{
   if(!p_op) return;
   p_op->xltype = xltypeBool;
   p_op->val._bool = (b ? 1 : 0);
}
//-------------------------------------------------------------------
void set_to_text(xloper *p_op, char *text)
{
   if(!p_op) return;

   if(!(p_op->val.str = new_xlstring(text)))
      p_op->xltype = xltypeNil;
   else
      p_op->xltype = xltypeStr;
}
//-------------------------------------------------------------------
void set_to_err(xloper *p_op, WORD e)
{
   if(!p_op) return;

   switch(e)
   {
   case xlerrNull:
   case xlerrDiv0:
   case xlerrValue:
   case xlerrRef:
   case xlerrName:
   case xlerrNum:
   case xlerrNA:
      p_op->xltype = xltypeErr;
      p_op->val.err = e;
      break;

   default:
      p_op->xltype = xltypeMissing; // not a valid error code
   }
}
//-------------------------------------------------------------------
bool set_to_xltypeMulti(xloper *p_op, WORD rows, WORD cols)
{
   int size = rows * cols;

   if(!p_op || !size || rows == 0xffff || cols > 0x00ff)
      return false;

   p_op->xltype = xltypeMulti;
   p_op->val.array.lparray = (xloper *)malloc(sizeof(xloper) * size);
   p_op->val.array.rows = rows;
   p_op->val.array.columns = cols;
   return true;
}
//-------------------------------------------------------------------
bool set_to_xltypeSRef(xloper *p_op, WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast)
{
   if(!p_op || rwFirst > rwLast || colFirst > colLast)
      return false;

// Create a simple single-cell reference to a cell on the current sheet
   p_op->xltype = xltypeSRef;
   p_op->val.sref.count = 1;

   xlref &ref = p_op->val.sref.ref; // to simplify code
   ref.rwFirst = rwFirst;
   ref.rwLast = rwLast;
   ref.colFirst = colFirst;
   ref.colLast = colLast;
   return true;
}
//-------------------------------------------------------------------
bool set_to_xltypeRef(xloper *p_op, DWORD idSheet, WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast)
{
   if(!p_op || rwFirst > rwLast || colFirst > colLast)
      return false;

// Allocate memory for the xlmref and set pointer within the xloper
   xlmref *p = (xlmref *)malloc(sizeof(xlmref));

   if(!p)
   {
      p_op->xltype = xltypeMissing;
      return false;
   }

   p_op->xltype = xltypeRef;
   p_op->val.mref.lpmref = p;
   p_op->val.mref.idSheet = idSheet;
   p_op->val.mref.lpmref->count = 1;

   xlref &ref = p->reftbl[0]; // to simplify code
   ref.rwFirst = rwFirst;
   ref.rwLast = rwLast;
   ref.colFirst = colFirst;
   ref.colLast = colLast;
   return true;
}
//-------------------------------------------------------------------
bool set_to_type(xloper *p_op, int new_type)
{
   if(!p_op)
      return false;

   switch(p_op->xltype = new_type)
   {
   case xltypeStr:
      p_op->val.str = NULL;
      break;

   case xltypeRef:
      p_op->val.mref.lpmref = NULL;
      p_op->val.mref.idSheet = 0;
      break;

   case xltypeMulti:
      p_op->val.array.lparray = NULL;
      p_op->val.array.rows = p_op->val.array.columns = 0;
      break;

   case xltypeInt:
   case xltypeNum:
   case xltypeBool:
   case xltypeErr:
   case xltypeSRef:
   case xltypeMissing:
   case xltypeNil:
   case xltypeBigData:
      break;

   default: // not valid or not supported
      p_op->xltype = 0;
      return false;
   }
   return true;
}
//-------------------------------------------------------------------
bool xlopers_equal(xloper &op1, xloper &op2)
{
   if(op1.xltype != op2.xltype)
      return false;

   int i;

   switch(op1.xltype)
   {
   case xltypeNum:
      if(op1.val.num != op2.val.num)
         return false;
      break;

   case xltypeStr:
      if(memcmp(op1.val.str + 1, op2.val.str + 1, *op1.val.str))
         return false;
      break;

   case xltypeBool:
      if(op1.val._bool != op2.val._bool)
         return false;
      break;

   case xltypeRef:
      if(op1.val.mref.idSheet != op2.val.mref.idSheet
      || op1.val.mref.lpmref->count != op2.val.mref.lpmref->count)
         return false;

      if(memcmp(op1.val.mref.lpmref->reftbl, op2.val.mref.lpmref->reftbl, sizeof(xlmref) * op1.val.mref.lpmref->count))
         return false;
      break;

   case xltypeSRef:
      if(memcmp(&(op1.val.sref.ref), &(op2.val.sref.ref), sizeof(xlref)))
         return false;
      break;

   case xltypeMulti:
      if(op1.val.array.columns != op2.val.array.columns
      || op1.val.array.rows != op2.val.array.rows)
         return false;

      for(i = op1.val.array.columns * op1.val.array.rows; i--; )
         if(!xlopers_equal(op1.val.array.lparray[i], op2.val.array.lparray[i]))
            return false;
      break;

   case xltypeMissing:
   case xltypeNil:
      return true;

   case xltypeInt:
      if(op1.val.w != op2.val.w)
         return false;
      break;

   case xltypeErr:
      if(op1.val.err != op2.val.err)
         return false;
      break;

   default:
      return false; // not suported by this function
   }
   return true;
}
//-------------------------------------------------------------------
// For those types that have memory associated, allocate new memory
// and copy contents.
//-------------------------------------------------------------------
bool clone_xloper(xloper *p_op)
{
   p_op->xltype &= ~(xlbitXLFree | xlbitDLLFree);

// These have no memory associated with them, so just return.
// The test for the xltypeInt intentionally catches the
// xltypeBigOper type.
   if(p_op->xltype & (xltypeInt | xltypeNum | xltypeBool | xltypeErr | xltypeMissing | xltypeNil | xltypeSRef))
      return false;

   switch(p_op->xltype)
   {
// These types have memory associated with them, so need to
// allocate new memory and then copy the contents from source
// memory pointers.
   case xltypeStr:
      p_op->val.str = copy_xlstring(p_op->val.str);
      return true;

   case xltypeRef: // assume lpmref->count == 1
      {
         xlmref *p = (xlmref *)malloc(sizeof(xlmref));
         memcpy(p, p_op->val.mref.lpmref, sizeof(xlmref));
         p_op->val.mref.lpmref = p;
         return true;
      }
      break;

   case xltypeMulti:
      {
         int limit = p_op->val.array.rows * p_op->val.array.columns;
         xloper *p = (xloper *)malloc(limit * sizeof(xloper));
         memcpy(p, p_op->val.array.lparray, limit * sizeof(xloper));
         p_op->val.array.lparray = p;

         for(int i = limit; i--; p++)
         {
            p->xltype &= ~(xlbitXLFree | xlbitDLLFree);

            switch(p->xltype)
            {
            case xltypeStr:
               p->val.str = copy_xlstring(p->val.str);
               break;

            case xltypeRef:
            case xltypeMulti:
               p->xltype = xltypeMissing; // not copied
               break;

            default: // other types all fine
               break;
            }
         }
         return true;
      }
      break;

   default: // other types not handled
      p_op->xltype = xltypeMissing;
      return false;
   }
}
//-------------------------------------------------------------------
//-------------------------------------------------------------------
// Variant conversion routines
//-------------------------------------------------------------------
//-------------------------------------------------------------------
// Converts a VT_BSTR wide-char string to a newly allocated C API
// byte-counted string.  Memory returned must be freed by caller.
//-------------------------------------------------------------------
char *vt_bstr_to_xlstring(BSTR bstr)
{
   if(!bstr)
      return NULL;

   int len = SysStringLen(bstr);

   if(len > 255)
      len = 255; // truncate

   char *p = (char *)malloc(len + 2);

// VT_BSTR is a wchar_t string, so need to convert to a byte-string
   if(!p || wcstombs(p + 1, bstr, len + 1) < 0)
   {
      free(p);
      return false;
   }

   p[0] = (char)len;
   return p;
}
//-------------------------------------------------------------------
// Converts a C API byte-counted string to a VT_BSTR wide-char string
// Does not rely on (or assume) that input string is null-terminated.
//-------------------------------------------------------------------
BSTR xlstring_to_vt_bstr(char *str)
{
   if(!str)
      return NULL;

   wchar_t *p = (wchar_t *)malloc(str[0] * sizeof(wchar_t));

   if(!p || mbstowcs(p, str + 1, str[0]) < 0)
   {
      free(p);
      return NULL;
   }

   BSTR bstr = SysAllocStringLen(p, str[0]);
   free(p);
   return bstr;
}
//-------------------------------------------------------------------
// Returns true and a Variant equivalent to the passed-in xloper, or
// false if xloper's type could not be converted.
//-------------------------------------------------------------------
bool xloper_to_vt(xloper *p_op, VARIANT &var, bool convert_array)
{
   VariantInit(&var); // type is set to VT_EMPTY

   switch(p_op->xltype)
   {
   case xltypeNum:
      var.vt = VT_R8;
      var.dblVal = p_op->val.num;
      break;

   case xltypeInt:
      var.vt = VT_I2;
      var.iVal = p_op->val.w;
      break;

   case xltypeBool:
      var.vt = VT_BOOL;
      var.boolVal = p_op->val._bool;
      break;

   case xltypeStr:
      var.vt = VT_BSTR;
      var.bstrVal = xlstring_to_vt_bstr(p_op->val.str);
      break;

   case xltypeErr:
      var.vt = VT_ERROR;
      var.ulVal = VT_XL_ERR_OFFSET + p_op->val.err;
      break;

   case xltypeMulti:
      if(convert_array)
      {
         VARIANT temp_vt;
         SAFEARRAYBOUND bound[2];
         long elt_index[2];

         bound[0].lLbound = bound[1].lLbound = 0;
         bound[0].cElements = p_op->val.array.rows;
         bound[1].cElements = p_op->val.array.columns;

         var.vt = VT_ARRAY | VT_VARIANT; // array of Variants
         var.parray = SafeArrayCreate(VT_VARIANT, 2, bound);

         if(!var.parray)
            return false;

         xloper *p_op_temp = p_op->val.array.lparray;

         for(WORD r = 0; r < p_op->val.array.rows; r++)
         {
            for(WORD c = 0; c < p_op->val.array.columns;)
            {
               xloper_to_vt(p_op_temp++, temp_vt, false); // don't convert array within array

               elt_index[0] = r;
               elt_index[1] = c++;
               SafeArrayPutElement(var.parray, elt_index, &temp_vt);
            }
         }
         break;
      }
      // else, fall through to default option

   default: // type not converted
      return false;
   }
   return true;
}
//-------------------------------------------------------------------
// Converts the passed-in array Variant to an xltypeMulti xloper.
//-------------------------------------------------------------------
bool array_vt_to_xloper(xloper &op, VARIANT *pv)
{
   if(!pv || !pv->parray)
      return false;

   int ndims = pv->parray->cDims;

   if(ndims == 0 || ndims > 2)
      return false;

   int dims[2] = {1,1};
   
   long ubound[2], lbound[2];

   for(int d = 0; d < ndims; d++)
   {
// Dimension argument counts from 1, so need to pass d + 1
      if(FAILED(SafeArrayGetUBound(pv->parray, d + 1, ubound + d))
      || FAILED(SafeArrayGetLBound(pv->parray, d + 1, lbound + d)))
         return false;

      dims[d] = (short)(ubound[d] - lbound[d] + 1);
   }

   op.val.array.lparray = (xloper *)malloc(dims[0] * dims[1] * sizeof(xloper));

   if(!op.val.array.lparray)
      return false;

   op.val.array.rows = dims[0];
   op.val.array.columns = dims[1];
   op.xltype = xltypeMulti;

   xloper *p_op = op.val.array.lparray;

// Use this union structure to retrieve elements of the SafeArray
   union
   {
      VARIANT var;
      short I2;
      double R8;
      ULONG ulVal;
      CY cyVal;
      bool boolVal;
      BSTR bstrVal;
   }
      temp_union;

   xloper temp_op;
   VARTYPE vt_type = pv->vt & ~VT_ARRAY;
   long element_index[2];

   void *vp;
   SafeArrayAccessData(pv->parray, &vp);
   SafeArrayUnaccessData(pv->parray);

   for(WORD r = 0; r < dims[0]; r++)
   {
      for(WORD c = 0; c < dims[1];)
      {
         element_index[0] = r + lbound[0];
         element_index[1] = c++ + lbound[1];

         if(FAILED(SafeArrayGetElement(pv->parray, element_index, &temp_union)))
         {
            temp_op.xltype = xltypeNil;
         }
         else
         {
            switch(vt_type)
            {
            case VT_I2:
               temp_op.xltype = xltypeInt;
               temp_op.val.w = temp_union.I2;
               break;

            case VT_R8:
               temp_op.xltype = xltypeNum;
               temp_op.val.num = temp_union.R8;
               break;

            case VT_BOOL:
               temp_op.xltype = xltypeBool;
               temp_op.val._bool = temp_union.boolVal;
               break;

            case VT_VARIANT:
// Don't allow Variant to contain an array Variant to prevent recursion
               if(vt_to_xloper(temp_op, &temp_union.var, false))
                  break;

            case VT_ERROR:
               temp_op.xltype = xltypeErr;
               temp_op.val.err = (unsigned short)(temp_union.ulVal - VT_XL_ERR_OFFSET);
               break;

            case VT_CY:
               temp_op.xltype = xltypeNum;
               temp_op.val.num = (double)(temp_union.cyVal.int64 / 1e4);
               break;

            case VT_BSTR:
               op.xltype = xltypeStr;
               op.val.str = vt_bstr_to_xlstring(temp_union.bstrVal);
               break;

            default:
               temp_op.xltype = xltypeNil;
            }
         }
         *p_op++ = temp_op;
      }
   }
   return true;
}
//-------------------------------------------------------------------
// Converts the passed-in Variant to an xloper if one of the supported types.
//-------------------------------------------------------------------
bool vt_to_xloper(xloper &op, VARIANT *pv, bool convert_array)
{
   if(pv->vt & (VT_VECTOR | VT_BYREF))
      return false;

   if(pv->vt & VT_ARRAY)
   {
      if(!convert_array)
         return false;

      return array_vt_to_xloper(op, pv);
   }

   switch(pv->vt)
   {
   case VT_R8:
      op.xltype = xltypeNum;
      op.val.num = pv->dblVal;
      break;

   case VT_I2:
      op.xltype = xltypeInt;
      op.val.w = pv->iVal;
      break;

   case VT_BOOL:
      op.xltype = xltypeBool;
      op.val._bool = pv->boolVal;
      break;

   case VT_ERROR:
      op.xltype = xltypeErr;
      op.val.err = (unsigned short)(pv->ulVal - VT_XL_ERR_OFFSET);
      break;

   case VT_BSTR:
      op.xltype = xltypeStr;
      op.val.str = vt_bstr_to_xlstring(pv->bstrVal);
      break;

   case VT_CY:
      op.xltype = xltypeNum;
      op.val.num = (double)(pv->cyVal.int64 / 1e4);
      break;

   default: // type not converted
      return false;
   }
   return true;
}

//=====================================================
char * __stdcall oper_type_str(xloper *pxl)
{
   if(pxl == NULL)
      return NULL;

   switch(pxl->xltype)
   {
   case xltypeNum:      return "0x0001 xltypeNum";
   case xltypeStr:      return "0x0002 xltypeStr";
   case xltypeBool:   return "0x0004 xltypeBool";
   case xltypeErr:      return "0x0010 xltypeErr";
   case xltypeMulti:   return "0x0040 xltypeMulti";
   case xltypeMissing:   return "0x0080 xltypeMissing";
   case xltypeNil:      return "0x0100 xltypeNil";
   default:         return "Unexpected type";
   }
}
//=====================================================
char * __stdcall xloper_type_str(xloper *pxl)
{
   if(pxl == NULL)
      return NULL;

   switch(pxl->xltype)
   {
   case xltypeNum:      return "0x0001 xltypeNum";
   case xltypeStr:      return "0x0002 xltypeStr";
   case xltypeBool:   return "0x0004 xltypeBool";
   case xltypeRef:      return "0x0008 xltypeRef";
   case xltypeErr:      return "0x0010 xltypeErr";
   case xltypeMulti:   return "0x0040 xltypeMulti";
   case xltypeMissing:   return "0x0080 xltypeMissing";
   case xltypeNil:      return "0x0100 xltypeNil";
   case xltypeSRef:   return "0x0400 xltypeSRef";
   default:         return "Unexpected type";
   }
}

//=============================================
//=============================================
// WARNING: DO NOT COMPILE THE FOLLOWING CODE
//
// Provided for example only
//=============================================
//=============================================

#if 0
//=============================================
typedef struct _oper
{
   union
   {
      double num;
      BYTE *str;
      WORD _bool;
      WORD err;

      struct
      {
         struct _oper *lparray;
         WORD rows;
         WORD columns;
      }
         array;
   }
      val;

   WORD type;
}
   oper;


//==========================================================
int __stdcall xloper_type(xloper *p_op)
{
// Unsafe. Might be xltypeBigData
   if(p_op->xltype & xltypeStr)
      return xltypeStr;

// Unsafe. Might be xltypeBigData
   if(p_op->xltype & xltypeInt)
      return xltypeInt;

// Unsafe. Might be xltypeStr or xltypeInt
   if(p_op->xltype & xltypeBigData)
      return xltypeBigData;

// Unsafe. Might have xlbitXLFree or xlbitDLLFree set
   if(p_op->xltype == xltypeStr)
      return xltypeStr;

// Unsafe. Might have xlbitXLFree or xlbitDLLFree set
   if(p_op->xltype == xltypeMulti)
      return xltypeMulti;

// Unsafe. Might have xlbitXLFree or xlbitDLLFree set
   if(p_op->xltype == xltypeRef)
      return xltypeRef;

// Safe.
   if((p_op->xltype & xltypeBigData) == xltypeStr)
      return xltypeStr;

// Safe.
   if((p_op->xltype & ~(xlbitXLFree | xlbitDLLFree)) == xltypeRef)
      return xltypeRef;

   return 0; // not a valid xltype
}

//==========================================================
xloper * __stdcall get_a1_ref(xloper *sheet_name)
{
   static xloper ret_val;

   Excel4(xlSheetId, &ret_val, 1, sheet_name);

   if(ret_val.xltype == xltypeErr)
      return &ret_val;

// Sheet ID is contained in ret_val.val.mref.idSheet
// Now fill in the other fields to refer to the cell A1
   ret_val.val.mref.lpmref = (xlmref *)malloc(sizeof(xlmref));
   ret_val.val.mref.lpmref->count = 1;
   ret_val.val.mref.lpmref->reftbl[0].rwFirst = 0;
   ret_val.val.mref.lpmref->reftbl[0].rwLast = 0;
   ret_val.val.mref.lpmref->reftbl[0].colFirst = 0;
   ret_val.val.mref.lpmref->reftbl[0].colLast = 0;

// Ensure Excel calls back into the DLL to free the memory
   ret_val.xltype |= xlbitDLLFree;
   return &ret_val;
}

#endif