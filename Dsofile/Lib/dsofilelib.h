

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0500 */
/* at Thu Jul 26 14:26:54 2012
 */
/* Compiler settings for .\lib\dsofile.odl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __dsofilelib_h__
#define __dsofilelib_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __CustomProperty_FWD_DEFINED__
#define __CustomProperty_FWD_DEFINED__
typedef interface CustomProperty CustomProperty;
#endif 	/* __CustomProperty_FWD_DEFINED__ */


#ifndef __CustomProperties_FWD_DEFINED__
#define __CustomProperties_FWD_DEFINED__
typedef interface CustomProperties CustomProperties;
#endif 	/* __CustomProperties_FWD_DEFINED__ */


#ifndef __SummaryProperties_FWD_DEFINED__
#define __SummaryProperties_FWD_DEFINED__
typedef interface SummaryProperties SummaryProperties;
#endif 	/* __SummaryProperties_FWD_DEFINED__ */


#ifndef ___OleDocumentProperties_FWD_DEFINED__
#define ___OleDocumentProperties_FWD_DEFINED__
typedef interface _OleDocumentProperties _OleDocumentProperties;
#endif 	/* ___OleDocumentProperties_FWD_DEFINED__ */


#ifndef __OleDocumentProperties_FWD_DEFINED__
#define __OleDocumentProperties_FWD_DEFINED__

#ifdef __cplusplus
typedef class OleDocumentProperties OleDocumentProperties;
#else
typedef struct OleDocumentProperties OleDocumentProperties;
#endif /* __cplusplus */

#endif 	/* __OleDocumentProperties_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __DSOFile_LIBRARY_DEFINED__
#define __DSOFile_LIBRARY_DEFINED__

/* library DSOFile */
/* [lcid][version][helpstring][uuid] */ 

typedef 
enum dsoFilePropertyType
    {	dsoPropertyTypeUnknown	= 0,
	dsoPropertyTypeString	= 1,
	dsoPropertyTypeLong	= ( dsoPropertyTypeString + 1 ) ,
	dsoPropertyTypeDouble	= ( dsoPropertyTypeLong + 1 ) ,
	dsoPropertyTypeBool	= ( dsoPropertyTypeDouble + 1 ) ,
	dsoPropertyTypeDate	= ( dsoPropertyTypeBool + 1 ) 
    } 	dsoFilePropertyType;

typedef 
enum dsoFileOpenOptions
    {	dsoOptionDefault	= 0,
	dsoOptionOnlyOpenOLEFiles	= 1,
	dsoOptionOpenReadOnlyIfNoWriteAccess	= 2,
	dsoOptionDontAutoCreate	= 4,
	dsoOptionUseMBCStringsForNewSets	= 8
    } 	dsoFileOpenOptions;


DEFINE_GUID(LIBID_DSOFile,0x58968145,0xCF00,0x4341,0x99,0x5F,0x2E,0xE0,0x93,0xF6,0xAB,0xA3);

#ifndef __CustomProperty_INTERFACE_DEFINED__
#define __CustomProperty_INTERFACE_DEFINED__

/* interface CustomProperty */
/* [object][oleautomation][nonextensible][dual][uuid] */ 


DEFINE_GUID(IID_CustomProperty,0x58968145,0xCF03,0x4341,0x99,0x5F,0x2E,0xE0,0x93,0xF6,0xAB,0xA3);

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("58968145-CF03-4341-995F-2EE093F6ABA3")
    CustomProperty : public IDispatch
    {
    public:
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR *pbstrName) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Type( 
            /* [retval][out] */ dsoFilePropertyType *dsoType) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Value( 
            /* [retval][out] */ VARIANT *pvValue) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Value( 
            /* [in] */ VARIANT *pvValue) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct CustomPropertyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            CustomProperty * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            CustomProperty * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            CustomProperty * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            CustomProperty * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            CustomProperty * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            CustomProperty * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            CustomProperty * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Name )( 
            CustomProperty * This,
            /* [retval][out] */ BSTR *pbstrName);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Type )( 
            CustomProperty * This,
            /* [retval][out] */ dsoFilePropertyType *dsoType);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Value )( 
            CustomProperty * This,
            /* [retval][out] */ VARIANT *pvValue);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Value )( 
            CustomProperty * This,
            /* [in] */ VARIANT *pvValue);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Remove )( 
            CustomProperty * This);
        
        END_INTERFACE
    } CustomPropertyVtbl;

    interface CustomProperty
    {
        CONST_VTBL struct CustomPropertyVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define CustomProperty_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define CustomProperty_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define CustomProperty_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define CustomProperty_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define CustomProperty_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define CustomProperty_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define CustomProperty_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define CustomProperty_get_Name(This,pbstrName)	\
    ( (This)->lpVtbl -> get_Name(This,pbstrName) ) 

#define CustomProperty_get_Type(This,dsoType)	\
    ( (This)->lpVtbl -> get_Type(This,dsoType) ) 

#define CustomProperty_get_Value(This,pvValue)	\
    ( (This)->lpVtbl -> get_Value(This,pvValue) ) 

#define CustomProperty_put_Value(This,pvValue)	\
    ( (This)->lpVtbl -> put_Value(This,pvValue) ) 

#define CustomProperty_Remove(This)	\
    ( (This)->lpVtbl -> Remove(This) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __CustomProperty_INTERFACE_DEFINED__ */


#ifndef __CustomProperties_INTERFACE_DEFINED__
#define __CustomProperties_INTERFACE_DEFINED__

/* interface CustomProperties */
/* [object][oleautomation][nonextensible][dual][uuid] */ 


DEFINE_GUID(IID_CustomProperties,0x58968145,0xCF04,0x4341,0x99,0x5F,0x2E,0xE0,0x93,0xF6,0xAB,0xA3);

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("58968145-CF04-4341-995F-2EE093F6ABA3")
    CustomProperties : public IDispatch
    {
    public:
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long *lCount) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ BSTR sPropName,
            /* [in] */ VARIANT *Value,
            /* [retval][out] */ CustomProperty **ppDocProperty) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ CustomProperty **ppDocProperty) = 0;
        
        virtual /* [propget][restricted][id] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **ppunk) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct CustomPropertiesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            CustomProperties * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            CustomProperties * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            CustomProperties * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            CustomProperties * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            CustomProperties * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            CustomProperties * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            CustomProperties * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            CustomProperties * This,
            /* [retval][out] */ long *lCount);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Add )( 
            CustomProperties * This,
            /* [in] */ BSTR sPropName,
            /* [in] */ VARIANT *Value,
            /* [retval][out] */ CustomProperty **ppDocProperty);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            CustomProperties * This,
            /* [in] */ VARIANT Index,
            /* [retval][out] */ CustomProperty **ppDocProperty);
        
        /* [propget][restricted][id] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            CustomProperties * This,
            /* [retval][out] */ IUnknown **ppunk);
        
        END_INTERFACE
    } CustomPropertiesVtbl;

    interface CustomProperties
    {
        CONST_VTBL struct CustomPropertiesVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define CustomProperties_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define CustomProperties_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define CustomProperties_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define CustomProperties_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define CustomProperties_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define CustomProperties_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define CustomProperties_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define CustomProperties_get_Count(This,lCount)	\
    ( (This)->lpVtbl -> get_Count(This,lCount) ) 

#define CustomProperties_Add(This,sPropName,Value,ppDocProperty)	\
    ( (This)->lpVtbl -> Add(This,sPropName,Value,ppDocProperty) ) 

#define CustomProperties_get_Item(This,Index,ppDocProperty)	\
    ( (This)->lpVtbl -> get_Item(This,Index,ppDocProperty) ) 

#define CustomProperties_get__NewEnum(This,ppunk)	\
    ( (This)->lpVtbl -> get__NewEnum(This,ppunk) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __CustomProperties_INTERFACE_DEFINED__ */


#ifndef __SummaryProperties_INTERFACE_DEFINED__
#define __SummaryProperties_INTERFACE_DEFINED__

/* interface SummaryProperties */
/* [object][oleautomation][nonextensible][dual][uuid] */ 


DEFINE_GUID(IID_SummaryProperties,0x58968145,0xCF02,0x4341,0x99,0x5F,0x2E,0xE0,0x93,0xF6,0xAB,0xA3);

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("58968145-CF02-4341-995F-2EE093F6ABA3")
    SummaryProperties : public IDispatch
    {
    public:
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Title( 
            /* [retval][out] */ BSTR *pbstrTitle) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Title( 
            /* [in] */ BSTR bstrTitle) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Subject( 
            /* [retval][out] */ BSTR *pbstrSubject) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Subject( 
            /* [in] */ BSTR bstrSubject) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Author( 
            /* [retval][out] */ BSTR *pbstrAuthor) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Author( 
            /* [in] */ BSTR bstrAuthor) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Keywords( 
            /* [retval][out] */ BSTR *pbstrKeywords) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Keywords( 
            /* [in] */ BSTR bstrKeywords) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Comments( 
            /* [retval][out] */ BSTR *pbstrComments) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Comments( 
            /* [in] */ BSTR bstrComments) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Template( 
            /* [retval][out] */ BSTR *pbstrTemplate) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_LastSavedBy( 
            /* [retval][out] */ BSTR *pbstrLastSavedBy) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_LastSavedBy( 
            /* [in] */ BSTR bstrLastSavedBy) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_RevisionNumber( 
            /* [retval][out] */ BSTR *pbstrRevisionNumber) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_TotalEditTime( 
            /* [retval][out] */ long *plTotalEditTime) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_DateLastPrinted( 
            /* [retval][out] */ VARIANT *pdtDateLastPrinted) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_DateCreated( 
            /* [retval][out] */ VARIANT *pdtDateCreated) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_DateLastSaved( 
            /* [retval][out] */ VARIANT *pdtDateLastSaved) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_PageCount( 
            /* [retval][out] */ long *plPageCount) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_WordCount( 
            /* [retval][out] */ long *plWordCount) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_CharacterCount( 
            /* [retval][out] */ long *plCharacterCount) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Thumbnail( 
            /* [retval][out] */ VARIANT *pvtThumbnail) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_ApplicationName( 
            /* [retval][out] */ BSTR *pbstrAppName) = 0;
        
        virtual /* [helpstring][propget][hidden][id] */ HRESULT STDMETHODCALLTYPE get_DocumentSecurity( 
            /* [retval][out] */ long *plDocSecurity) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Category( 
            /* [retval][out] */ BSTR *pbstrCategory) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Category( 
            /* [in] */ BSTR bstrCategory) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_PresentationFormat( 
            /* [retval][out] */ BSTR *pbstrPresFormat) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_ByteCount( 
            /* [retval][out] */ long *plByteCount) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_LineCount( 
            /* [retval][out] */ long *plLineCount) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_ParagraphCount( 
            /* [retval][out] */ long *plParagraphCount) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_SlideCount( 
            /* [retval][out] */ long *plSlideCount) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_NoteCount( 
            /* [retval][out] */ long *plPresNotes) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_HiddenSlideCount( 
            /* [retval][out] */ long *plHiddenSlides) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_MultimediaClipCount( 
            /* [retval][out] */ long *plMultimediaClips) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Manager( 
            /* [retval][out] */ BSTR *pbstrManager) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Manager( 
            /* [in] */ BSTR bstrManager) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Company( 
            /* [retval][out] */ BSTR *pbstrCompany) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Company( 
            /* [in] */ BSTR bstrCompany) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_CharacterCountWithSpaces( 
            /* [retval][out] */ long *plCharCountWithSpaces) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_SharedDocument( 
            /* [retval][out] */ VARIANT_BOOL *pbSharedDocument) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Version( 
            /* [retval][out] */ BSTR *pbstrVersion) = 0;
        
        virtual /* [helpstring][propget][hidden][id] */ HRESULT STDMETHODCALLTYPE get_DigitalSignature( 
            /* [retval][out] */ VARIANT *pvtDigSig) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct SummaryPropertiesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            SummaryProperties * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            SummaryProperties * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            SummaryProperties * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            SummaryProperties * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            SummaryProperties * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            SummaryProperties * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            SummaryProperties * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Title )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrTitle);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Title )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrTitle);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Subject )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrSubject);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Subject )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrSubject);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Author )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrAuthor);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Author )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrAuthor);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Keywords )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrKeywords);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Keywords )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrKeywords);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Comments )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrComments);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Comments )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrComments);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Template )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrTemplate);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_LastSavedBy )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrLastSavedBy);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_LastSavedBy )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrLastSavedBy);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_RevisionNumber )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrRevisionNumber);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_TotalEditTime )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plTotalEditTime);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_DateLastPrinted )( 
            SummaryProperties * This,
            /* [retval][out] */ VARIANT *pdtDateLastPrinted);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_DateCreated )( 
            SummaryProperties * This,
            /* [retval][out] */ VARIANT *pdtDateCreated);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_DateLastSaved )( 
            SummaryProperties * This,
            /* [retval][out] */ VARIANT *pdtDateLastSaved);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_PageCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plPageCount);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_WordCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plWordCount);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_CharacterCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plCharacterCount);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Thumbnail )( 
            SummaryProperties * This,
            /* [retval][out] */ VARIANT *pvtThumbnail);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_ApplicationName )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrAppName);
        
        /* [helpstring][propget][hidden][id] */ HRESULT ( STDMETHODCALLTYPE *get_DocumentSecurity )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plDocSecurity);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Category )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrCategory);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Category )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrCategory);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_PresentationFormat )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrPresFormat);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_ByteCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plByteCount);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_LineCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plLineCount);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_ParagraphCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plParagraphCount);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_SlideCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plSlideCount);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_NoteCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plPresNotes);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_HiddenSlideCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plHiddenSlides);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_MultimediaClipCount )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plMultimediaClips);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Manager )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrManager);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Manager )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrManager);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Company )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrCompany);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Company )( 
            SummaryProperties * This,
            /* [in] */ BSTR bstrCompany);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_CharacterCountWithSpaces )( 
            SummaryProperties * This,
            /* [retval][out] */ long *plCharCountWithSpaces);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_SharedDocument )( 
            SummaryProperties * This,
            /* [retval][out] */ VARIANT_BOOL *pbSharedDocument);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Version )( 
            SummaryProperties * This,
            /* [retval][out] */ BSTR *pbstrVersion);
        
        /* [helpstring][propget][hidden][id] */ HRESULT ( STDMETHODCALLTYPE *get_DigitalSignature )( 
            SummaryProperties * This,
            /* [retval][out] */ VARIANT *pvtDigSig);
        
        END_INTERFACE
    } SummaryPropertiesVtbl;

    interface SummaryProperties
    {
        CONST_VTBL struct SummaryPropertiesVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define SummaryProperties_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define SummaryProperties_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define SummaryProperties_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define SummaryProperties_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define SummaryProperties_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define SummaryProperties_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define SummaryProperties_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define SummaryProperties_get_Title(This,pbstrTitle)	\
    ( (This)->lpVtbl -> get_Title(This,pbstrTitle) ) 

#define SummaryProperties_put_Title(This,bstrTitle)	\
    ( (This)->lpVtbl -> put_Title(This,bstrTitle) ) 

#define SummaryProperties_get_Subject(This,pbstrSubject)	\
    ( (This)->lpVtbl -> get_Subject(This,pbstrSubject) ) 

#define SummaryProperties_put_Subject(This,bstrSubject)	\
    ( (This)->lpVtbl -> put_Subject(This,bstrSubject) ) 

#define SummaryProperties_get_Author(This,pbstrAuthor)	\
    ( (This)->lpVtbl -> get_Author(This,pbstrAuthor) ) 

#define SummaryProperties_put_Author(This,bstrAuthor)	\
    ( (This)->lpVtbl -> put_Author(This,bstrAuthor) ) 

#define SummaryProperties_get_Keywords(This,pbstrKeywords)	\
    ( (This)->lpVtbl -> get_Keywords(This,pbstrKeywords) ) 

#define SummaryProperties_put_Keywords(This,bstrKeywords)	\
    ( (This)->lpVtbl -> put_Keywords(This,bstrKeywords) ) 

#define SummaryProperties_get_Comments(This,pbstrComments)	\
    ( (This)->lpVtbl -> get_Comments(This,pbstrComments) ) 

#define SummaryProperties_put_Comments(This,bstrComments)	\
    ( (This)->lpVtbl -> put_Comments(This,bstrComments) ) 

#define SummaryProperties_get_Template(This,pbstrTemplate)	\
    ( (This)->lpVtbl -> get_Template(This,pbstrTemplate) ) 

#define SummaryProperties_get_LastSavedBy(This,pbstrLastSavedBy)	\
    ( (This)->lpVtbl -> get_LastSavedBy(This,pbstrLastSavedBy) ) 

#define SummaryProperties_put_LastSavedBy(This,bstrLastSavedBy)	\
    ( (This)->lpVtbl -> put_LastSavedBy(This,bstrLastSavedBy) ) 

#define SummaryProperties_get_RevisionNumber(This,pbstrRevisionNumber)	\
    ( (This)->lpVtbl -> get_RevisionNumber(This,pbstrRevisionNumber) ) 

#define SummaryProperties_get_TotalEditTime(This,plTotalEditTime)	\
    ( (This)->lpVtbl -> get_TotalEditTime(This,plTotalEditTime) ) 

#define SummaryProperties_get_DateLastPrinted(This,pdtDateLastPrinted)	\
    ( (This)->lpVtbl -> get_DateLastPrinted(This,pdtDateLastPrinted) ) 

#define SummaryProperties_get_DateCreated(This,pdtDateCreated)	\
    ( (This)->lpVtbl -> get_DateCreated(This,pdtDateCreated) ) 

#define SummaryProperties_get_DateLastSaved(This,pdtDateLastSaved)	\
    ( (This)->lpVtbl -> get_DateLastSaved(This,pdtDateLastSaved) ) 

#define SummaryProperties_get_PageCount(This,plPageCount)	\
    ( (This)->lpVtbl -> get_PageCount(This,plPageCount) ) 

#define SummaryProperties_get_WordCount(This,plWordCount)	\
    ( (This)->lpVtbl -> get_WordCount(This,plWordCount) ) 

#define SummaryProperties_get_CharacterCount(This,plCharacterCount)	\
    ( (This)->lpVtbl -> get_CharacterCount(This,plCharacterCount) ) 

#define SummaryProperties_get_Thumbnail(This,pvtThumbnail)	\
    ( (This)->lpVtbl -> get_Thumbnail(This,pvtThumbnail) ) 

#define SummaryProperties_get_ApplicationName(This,pbstrAppName)	\
    ( (This)->lpVtbl -> get_ApplicationName(This,pbstrAppName) ) 

#define SummaryProperties_get_DocumentSecurity(This,plDocSecurity)	\
    ( (This)->lpVtbl -> get_DocumentSecurity(This,plDocSecurity) ) 

#define SummaryProperties_get_Category(This,pbstrCategory)	\
    ( (This)->lpVtbl -> get_Category(This,pbstrCategory) ) 

#define SummaryProperties_put_Category(This,bstrCategory)	\
    ( (This)->lpVtbl -> put_Category(This,bstrCategory) ) 

#define SummaryProperties_get_PresentationFormat(This,pbstrPresFormat)	\
    ( (This)->lpVtbl -> get_PresentationFormat(This,pbstrPresFormat) ) 

#define SummaryProperties_get_ByteCount(This,plByteCount)	\
    ( (This)->lpVtbl -> get_ByteCount(This,plByteCount) ) 

#define SummaryProperties_get_LineCount(This,plLineCount)	\
    ( (This)->lpVtbl -> get_LineCount(This,plLineCount) ) 

#define SummaryProperties_get_ParagraphCount(This,plParagraphCount)	\
    ( (This)->lpVtbl -> get_ParagraphCount(This,plParagraphCount) ) 

#define SummaryProperties_get_SlideCount(This,plSlideCount)	\
    ( (This)->lpVtbl -> get_SlideCount(This,plSlideCount) ) 

#define SummaryProperties_get_NoteCount(This,plPresNotes)	\
    ( (This)->lpVtbl -> get_NoteCount(This,plPresNotes) ) 

#define SummaryProperties_get_HiddenSlideCount(This,plHiddenSlides)	\
    ( (This)->lpVtbl -> get_HiddenSlideCount(This,plHiddenSlides) ) 

#define SummaryProperties_get_MultimediaClipCount(This,plMultimediaClips)	\
    ( (This)->lpVtbl -> get_MultimediaClipCount(This,plMultimediaClips) ) 

#define SummaryProperties_get_Manager(This,pbstrManager)	\
    ( (This)->lpVtbl -> get_Manager(This,pbstrManager) ) 

#define SummaryProperties_put_Manager(This,bstrManager)	\
    ( (This)->lpVtbl -> put_Manager(This,bstrManager) ) 

#define SummaryProperties_get_Company(This,pbstrCompany)	\
    ( (This)->lpVtbl -> get_Company(This,pbstrCompany) ) 

#define SummaryProperties_put_Company(This,bstrCompany)	\
    ( (This)->lpVtbl -> put_Company(This,bstrCompany) ) 

#define SummaryProperties_get_CharacterCountWithSpaces(This,plCharCountWithSpaces)	\
    ( (This)->lpVtbl -> get_CharacterCountWithSpaces(This,plCharCountWithSpaces) ) 

#define SummaryProperties_get_SharedDocument(This,pbSharedDocument)	\
    ( (This)->lpVtbl -> get_SharedDocument(This,pbSharedDocument) ) 

#define SummaryProperties_get_Version(This,pbstrVersion)	\
    ( (This)->lpVtbl -> get_Version(This,pbstrVersion) ) 

#define SummaryProperties_get_DigitalSignature(This,pvtDigSig)	\
    ( (This)->lpVtbl -> get_DigitalSignature(This,pvtDigSig) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __SummaryProperties_INTERFACE_DEFINED__ */


#ifndef ___OleDocumentProperties_INTERFACE_DEFINED__
#define ___OleDocumentProperties_INTERFACE_DEFINED__

/* interface _OleDocumentProperties */
/* [object][oleautomation][nonextensible][dual][hidden][uuid] */ 


DEFINE_GUID(IID__OleDocumentProperties,0x58968145,0xCF01,0x4341,0x99,0x5F,0x2E,0xE0,0x93,0xF6,0xAB,0xA3);

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("58968145-CF01-4341-995F-2EE093F6ABA3")
    _OleDocumentProperties : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Open( 
            /* [in] */ BSTR sFileName,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL ReadOnly = 0,
            /* [defaultvalue][optional][in] */ dsoFileOpenOptions Options = dsoOptionDefault) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Close( 
            /* [defaultvalue][optional][in] */ VARIANT_BOOL SaveBeforeClose = 0) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_IsReadOnly( 
            /* [retval][out] */ VARIANT_BOOL *pbReadOnly) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_IsDirty( 
            /* [retval][out] */ VARIANT_BOOL *pbDirty) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Save( void) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_SummaryProperties( 
            /* [retval][out] */ SummaryProperties **ppSummaryProperties) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_CustomProperties( 
            /* [retval][out] */ CustomProperties **ppCustomProperties) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Icon( 
            /* [retval][out] */ IDispatch **ppicIcon) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR *pbstrName) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Path( 
            /* [retval][out] */ BSTR *pbstrPath) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_IsOleFile( 
            /* [retval][out] */ VARIANT_BOOL *pIsOleFile) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_CLSID( 
            /* [retval][out] */ BSTR *pbstrCLSID) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_ProgID( 
            /* [retval][out] */ BSTR *pbstrProgID) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_OleDocumentFormat( 
            /* [retval][out] */ BSTR *pbstrFormat) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_OleDocumentType( 
            /* [retval][out] */ BSTR *pbstrType) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct _OleDocumentPropertiesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _OleDocumentProperties * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _OleDocumentProperties * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _OleDocumentProperties * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _OleDocumentProperties * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _OleDocumentProperties * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _OleDocumentProperties * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _OleDocumentProperties * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Open )( 
            _OleDocumentProperties * This,
            /* [in] */ BSTR sFileName,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL ReadOnly,
            /* [defaultvalue][optional][in] */ dsoFileOpenOptions Options);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Close )( 
            _OleDocumentProperties * This,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL SaveBeforeClose);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_IsReadOnly )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ VARIANT_BOOL *pbReadOnly);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_IsDirty )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ VARIANT_BOOL *pbDirty);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Save )( 
            _OleDocumentProperties * This);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_SummaryProperties )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ SummaryProperties **ppSummaryProperties);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_CustomProperties )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ CustomProperties **ppCustomProperties);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Icon )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ IDispatch **ppicIcon);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Name )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ BSTR *pbstrName);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Path )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ BSTR *pbstrPath);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_IsOleFile )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ VARIANT_BOOL *pIsOleFile);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_CLSID )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ BSTR *pbstrCLSID);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_ProgID )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ BSTR *pbstrProgID);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_OleDocumentFormat )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ BSTR *pbstrFormat);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_OleDocumentType )( 
            _OleDocumentProperties * This,
            /* [retval][out] */ BSTR *pbstrType);
        
        END_INTERFACE
    } _OleDocumentPropertiesVtbl;

    interface _OleDocumentProperties
    {
        CONST_VTBL struct _OleDocumentPropertiesVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _OleDocumentProperties_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _OleDocumentProperties_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _OleDocumentProperties_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _OleDocumentProperties_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _OleDocumentProperties_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _OleDocumentProperties_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _OleDocumentProperties_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define _OleDocumentProperties_Open(This,sFileName,ReadOnly,Options)	\
    ( (This)->lpVtbl -> Open(This,sFileName,ReadOnly,Options) ) 

#define _OleDocumentProperties_Close(This,SaveBeforeClose)	\
    ( (This)->lpVtbl -> Close(This,SaveBeforeClose) ) 

#define _OleDocumentProperties_get_IsReadOnly(This,pbReadOnly)	\
    ( (This)->lpVtbl -> get_IsReadOnly(This,pbReadOnly) ) 

#define _OleDocumentProperties_get_IsDirty(This,pbDirty)	\
    ( (This)->lpVtbl -> get_IsDirty(This,pbDirty) ) 

#define _OleDocumentProperties_Save(This)	\
    ( (This)->lpVtbl -> Save(This) ) 

#define _OleDocumentProperties_get_SummaryProperties(This,ppSummaryProperties)	\
    ( (This)->lpVtbl -> get_SummaryProperties(This,ppSummaryProperties) ) 

#define _OleDocumentProperties_get_CustomProperties(This,ppCustomProperties)	\
    ( (This)->lpVtbl -> get_CustomProperties(This,ppCustomProperties) ) 

#define _OleDocumentProperties_get_Icon(This,ppicIcon)	\
    ( (This)->lpVtbl -> get_Icon(This,ppicIcon) ) 

#define _OleDocumentProperties_get_Name(This,pbstrName)	\
    ( (This)->lpVtbl -> get_Name(This,pbstrName) ) 

#define _OleDocumentProperties_get_Path(This,pbstrPath)	\
    ( (This)->lpVtbl -> get_Path(This,pbstrPath) ) 

#define _OleDocumentProperties_get_IsOleFile(This,pIsOleFile)	\
    ( (This)->lpVtbl -> get_IsOleFile(This,pIsOleFile) ) 

#define _OleDocumentProperties_get_CLSID(This,pbstrCLSID)	\
    ( (This)->lpVtbl -> get_CLSID(This,pbstrCLSID) ) 

#define _OleDocumentProperties_get_ProgID(This,pbstrProgID)	\
    ( (This)->lpVtbl -> get_ProgID(This,pbstrProgID) ) 

#define _OleDocumentProperties_get_OleDocumentFormat(This,pbstrFormat)	\
    ( (This)->lpVtbl -> get_OleDocumentFormat(This,pbstrFormat) ) 

#define _OleDocumentProperties_get_OleDocumentType(This,pbstrType)	\
    ( (This)->lpVtbl -> get_OleDocumentType(This,pbstrType) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* ___OleDocumentProperties_INTERFACE_DEFINED__ */


DEFINE_GUID(CLSID_OleDocumentProperties,0x58968145,0xCF05,0x4341,0x99,0x5F,0x2E,0xE0,0x93,0xF6,0xAB,0xA3);

#ifdef __cplusplus

class DECLSPEC_UUID("58968145-CF05-4341-995F-2EE093F6ABA3")
OleDocumentProperties;
#endif
#endif /* __DSOFile_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


