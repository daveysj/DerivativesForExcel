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
// This source file (and header file) contain the definition of the xl_array
// data type used by Excel to pass simple arrays of doubles.  The source file
// contains examples relating to the use of this data type.
//============================================================================
//============================================================================
#include <windows.h>

#include "xloper.h"

//====================================================================
xl_array * __stdcall xl_array_example1(int rows, int columns)
{
   static xl_array *p_array = NULL;

   if(p_array) // free memory allocated on last call
   {
      free(p_array);
      p_array = NULL;
   }

   int size = rows * columns;

   if(size <= 0)
      return NULL;

   size_t mem_size = sizeof(xl_array) + (size - 1) * sizeof(double);

   if((p_array = (xl_array *)malloc(mem_size)))
   {
      p_array->rows = rows;
      p_array->columns = columns;

      for(int i = 0; i < size; i++)
         p_array->array[i] = i / 10.0;
   }
   return p_array;
}
//====================================================================
void __stdcall xl_array_example2(xl_array *p_array)
{
   if(!p_array || !p_array->rows
   || !p_array->columns || p_array->columns > 0x100)
      return;

   int size = p_array->rows * p_array->columns;

// Change the values in the array
   for(int i = 0; i < size; i++)
      p_array->array[i] /= 10.0;
}
//====================================================================
void __stdcall xl_array_example3(xl_array *p_array)
{
   if(!p_array || !p_array->rows
   || !p_array->columns || p_array->columns > 0x100)
      return;

   int size = p_array->rows * p_array->columns;

// Change the shape of the array but not the size
   int temp = p_array->rows;
   p_array->rows = p_array->columns;
   p_array->columns = temp;

// Change the values in the array
   for(int i = 0; i < size; i++)
      p_array->array[i] /= 10.0;
}
//====================================================================
void __stdcall xl_array_example4(xl_array *p_array)
{
   if(!p_array || !p_array->rows
   || !p_array->columns || p_array->columns > 0x100)
      return;

// Reduce the size of the array
   if(p_array->rows > 1)
      p_array->rows--;

   if(p_array->columns > 1)
      p_array->columns--;

   int size = p_array->rows * p_array->columns;

// Change the values in the array
   for(int i = 0; i < size; i++)
      p_array->array[i] /= 10.0;
}
//====================================================================
xl_array *new_xl_array(WORD rows, WORD columns, double *array)
{
   if(!array || rows <= 0 || columns <= 0)
      return NULL;

   int size = rows * columns;

   xl_array *ret_array = (xl_array *)malloc(sizeof(xl_array) + sizeof(double) * (size - 1));
   ret_array->columns = columns;
   ret_array->rows = rows;
   memcpy(ret_array->array, array, size * sizeof(double));
   return ret_array;
}
//====================================================================
xl_array *new_xl_array(WORD rows, WORD columns)
{
   if(rows <= 0 || columns <= 0)
      return NULL;

   int size = rows * columns;

   xl_array *ret_array = (xl_array *)calloc(1, sizeof(xl_array) + sizeof(double) * (size - 1));
   ret_array->columns = columns;
   ret_array->rows = rows;
   return ret_array;
}
//====================================================================
bool transpose_xl_array(xl_array *p)
{
   if(!p || p->columns < 1 || p->rows < 1)
      return false;

   DWORD size = p->columns * p->rows;

   double *new_array = (double *)malloc(sizeof(double) * size);

   WORD r, c;
   double *p_source = p->array;
   double *p_target;

   for(r = 0; r < p->rows; r++)
   {
      p_target = new_array + r;

      for(c = 0; c < p->columns; c++, p_target += p->rows)
      {
         *p_target = *p_source++;
      }
   }

   r = p->columns;
   p->columns = p->rows;
   p->rows = r;

   memcpy(p->array, new_array, sizeof(double) * size);
   free(new_array);
   return true;
}
//=====================================================================
// This worksheet function returns the transposed array via the
// first argument modified in place.  It needs to be declared as
// returning void, so a wrapper to the function transpose_xl_array()
// is required.
//=====================================================================
void __stdcall xl_array_transpose(xl_array *input)
{
   transpose_xl_array(input); // ignore the return value
}
//=====================================================================


//====================================================================
//====================================================================
// WARNING: DO NOT COMPILE THE FOLLOWING CODE
//
// Provided for example only
//====================================================================
//====================================================================

#if 0


typedef struct
{
   WORD rows;
   WORD columns;
   double *array;
}
   xl_array;  // NO IT IS NOT!!!

xl_array * __stdcall bad_xl_array_example(int rows, int columns)
{
   static xl_array rtn_array = {0,0, NULL};

   if(rtn_array.array) // free memory allocated on last call
   {
      free(rtn_array.array);
      rtn_array.array = NULL;
   }

   int size = rows * columns;

   if(size <= 0)
      return NULL;

   if(!(rtn_array.array = (double *)malloc(size * sizeof(double))))
   {
      rtn_array.rows = rows;
      rtn_array.columns = columns;

      for(int i = 0; i < size; i++)
         rtn_array.array[i] = i / 10.0;
   }
   return &rtn_array;
}

#endif