#pragma once

// MFCActiveXControl1Ctrl.h : Declaration of the CMFCActiveXControl1Ctrl ActiveX Control class.
#include <objsafe.h>

// CMFCActiveXControl1Ctrl : See MFCActiveXControl1Ctrl.cpp for implementation.

class CMFCActiveXControl1Ctrl : public COleControl
{
	DECLARE_DYNCREATE(CMFCActiveXControl1Ctrl)

// Constructor
public:
	CMFCActiveXControl1Ctrl();

// Overrides
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// Implementation
protected:
	~CMFCActiveXControl1Ctrl();

	DECLARE_OLECREATE_EX(CMFCActiveXControl1Ctrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CMFCActiveXControl1Ctrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CMFCActiveXControl1Ctrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CMFCActiveXControl1Ctrl)		// Type name and misc status

// Message maps
	afx_msg void OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// Event maps
	DECLARE_EVENT_MAP()

protected:

	DECLARE_INTERFACE_MAP()
	BEGIN_INTERFACE_PART(ObjSafe, IObjectSafety)
		STDMETHOD_(HRESULT, GetInterfaceSafetyOptions)(
			/* [in]  */ REFIID riid,
			/* [out] */ DWORD  __RPC_FAR *pdwSupportedOptions,
			/* [out] */ DWORD  __RPC_FAR *pdwEnabledOptions
			);

		STDMETHOD_(HRESULT, SetInterfaceSafetyOptions)(
			/* [in]  */ REFIID riid,
			/* [in]  */ DWORD  dwOptionSetMask,
			/* [in]  */ DWORD  dwEnabledOptions
			);
	END_INTERFACE_PART(ObjSafe);

	BEGIN_INTERFACE_PART(PersistStorage, IPersistStorage)
		INIT_INTERFACE_PART(CMFCActiveXControl1Ctrl, PersistStorage)
		STDMETHOD(GetClassID)(LPCLSID);
		STDMETHOD(IsDirty)();
		STDMETHOD(InitNew)(LPSTORAGE);
		STDMETHOD(Load)(LPSTORAGE);
		STDMETHOD(Save)(LPSTORAGE, BOOL);
		STDMETHOD(SaveCompleted)(LPSTORAGE);
		STDMETHOD(HandsOffStorage)();
	END_INTERFACE_PART(PersistStorage)

public:
	BOOL OnRenderData(LPFORMATETC pfmt, LPSTGMEDIUM pmedium);

private:
		BOOL RenderObjectDescriptor(LPSTGMEDIUM pmedium);
		BOOL RenderEmbeddedObject(LPSTGMEDIUM pmedium);
		BOOL AllocateStgMedium(LPSTGMEDIUM pmedium, TYMED tymed, DWORD cbSize = 0);
		void OnCopy(BOOL bActiveX = TRUE);

// Dispatch and event IDs
public:
	enum {
	};
};

