

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 04:14:07 2038
 */
/* Compiler settings for MFCActiveXControl1.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0622 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        EXTERN_C __declspec(selectany) const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif // !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, LIBID_MFCActiveXControl1Lib,0x46374548,0xd2d0,0x4396,0xa0,0xcc,0xdd,0x75,0x21,0x8a,0x57,0xf9);


MIDL_DEFINE_GUID(IID, DIID__DMFCActiveXControl1,0xc2abec10,0x0924,0x4d4c,0xbf,0xb4,0x7c,0xb0,0xc8,0xcb,0xc5,0x76);


MIDL_DEFINE_GUID(IID, DIID__DMFCActiveXControl1Events,0xbd05f7c0,0x1f73,0x4fbc,0xbc,0xbc,0x3e,0xc2,0x2e,0x24,0xa7,0x73);


MIDL_DEFINE_GUID(CLSID, CLSID_MFCActiveXControl1,0x23646dca,0xd367,0x481d,0xbf,0x77,0xd9,0x4b,0x46,0x4a,0x1e,0xa5);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



