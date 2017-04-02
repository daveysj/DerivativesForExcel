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
#ifndef _XLOPER_H
#define _XLOPER_H

#ifndef VARIANT
#include <ole2.h>
#endif

#ifndef _XLCALL_H
#include "xlcall.h"
#endif

#ifndef _XL_ARRAY_H
#include "xl_array.h"
#endif

//===================================================================
// Frees xloper memory using either Excel or free().  If use_xlbits
// is true, then only does this where the appropriate bit is set. If
// use_xlbits is false, then uses free() and assumes that all types
// that have memory were allocated in a way that is compatible with
// freeing by a call to free().  Calling this with use_xlbits false
// on an xloper pointing to static memory will cause an exception.
//===================================================================
void free_xloper(xloper *p_op, bool use_xlbits);

//===================================================================

bool is_xloper_missing(xloper *p_op);
char *copy_xlstring(char *xlstring);
char *new_xlstring(char *text);
bool coerce_xloper(xloper *p_op, xloper &ret_val, int target_type);
bool coerce_to_string(xloper *p_op, char *&text); // makes new string
bool coerce_to_double(xloper *p_op, double &d);
bool coerce_to_int(xloper *p_op, int &w);
bool coerce_to_short(xloper *p_op, short &s);
bool coerce_to_bool(xloper *p_op, bool &b);
double *coerce_to_double_array(xloper *p_op, double invalid_value, WORD &cols, WORD &rows);
void set_to_double(xloper *p_op, double d);
void set_to_int(xloper *p_op, int w);
void set_to_bool(xloper *p_op, bool b);
void set_to_text(xloper *p_op, char *text);
void set_to_err(xloper *p_op, WORD e);
bool set_to_xltypeMulti(xloper *p_op, WORD rows, WORD cols);
bool set_to_xltypeSRef(xloper *p_op, WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast);
bool set_to_xltypeRef(xloper *p_op, DWORD idSheet, WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast);
bool set_to_type(xloper *p_op, int new_type);
bool xlopers_equal(xloper &op1, xloper &op2);
bool clone_xloper(xloper *p_op);

// Routines for converting Variants to/from xlopers
bool xloper_to_vt(xloper *p_op, VARIANT &var, bool convert_array);
bool vt_to_xloper(xloper &op, VARIANT *pv, bool convert_array);

extern xloper *p_xlTrue;
extern xloper *p_xlFalse;
extern xloper *p_xlMissing;
extern xloper *p_xlNil;
extern xloper *p_xlErrNull;
extern xloper *p_xlErrDiv0;
extern xloper *p_xlErrValue;
extern xloper *p_xlErrRef;
extern xloper *p_xlErrName;
extern xloper *p_xlErrNum;
extern xloper *p_xlErrNa;


#endif
