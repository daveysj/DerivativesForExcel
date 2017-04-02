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

#ifndef XL_ARRAY_H
#define XL_ARRAY_H

typedef struct
{
   WORD rows;
   WORD columns;
   double array[1]; // Start of array[rows * columns]
}
   xl_array;

xl_array *new_xl_array(WORD rows, WORD columns, double *array);
xl_array *new_xl_array(WORD rows, WORD columns);
bool transpose_xl_array(xl_array *p);


#endif
