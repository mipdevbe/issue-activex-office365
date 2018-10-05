// MFCActiveXControl1Ctrl.cpp : Implementation of the CMFCActiveXControl1Ctrl ActiveX Control class.

#include "stdafx.h"
#include "MFCActiveXControl1.h"
#include "MFCActiveXControl1Ctrl.h"
#include "MFCActiveXControl1PropPage.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CMFCActiveXControl1Ctrl, COleControl)

// Message map

BEGIN_MESSAGE_MAP(CMFCActiveXControl1Ctrl, COleControl)
	ON_WM_KEYUP()
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()

// Dispatch map

BEGIN_DISPATCH_MAP(CMFCActiveXControl1Ctrl, COleControl)
	DISP_FUNCTION_ID(CMFCActiveXControl1Ctrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CMFCActiveXControl1Ctrl, COleControl)
	INTERFACE_PART(CMFCActiveXControl1Ctrl, IID_IObjectSafety, ObjSafe)
	INTERFACE_PART(CMFCActiveXControl1Ctrl, IID_IPersistStorage, PersistStorage)
END_INTERFACE_MAP()

// Event map

BEGIN_EVENT_MAP(CMFCActiveXControl1Ctrl, COleControl)
END_EVENT_MAP()

// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CMFCActiveXControl1Ctrl, 1)
	PROPPAGEID(CMFCActiveXControl1PropPage::guid)
END_PROPPAGEIDS(CMFCActiveXControl1Ctrl)

// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMFCActiveXControl1Ctrl, "MFCACTIVEXCONTRO.MFCActiveXControl1Ctrl.1",
	0x23646dca,0xd367,0x481d,0xbf,0x77,0xd9,0x4b,0x46,0x4a,0x1e,0xa5)

// Type library ID and version

IMPLEMENT_OLETYPELIB(CMFCActiveXControl1Ctrl, _tlid, _wVerMajor, _wVerMinor)

// Interface IDs

const IID IID_DMFCActiveXControl1 = {0xc2abec10,0x0924,0x4d4c,{0xbf,0xb4,0x7c,0xb0,0xc8,0xcb,0xc5,0x76}};
const IID IID_DMFCActiveXControl1Events = {0xbd05f7c0,0x1f73,0x4fbc,{0xbc,0xbc,0x3e,0xc2,0x2e,0x24,0xa7,0x73}};

// Control type information

static const DWORD _dwMFCActiveXControl1OleMisc =
	OLEMISC_SIMPLEFRAME |
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CMFCActiveXControl1Ctrl, IDS_MFCACTIVEXCONTROL1, _dwMFCActiveXControl1OleMisc)

// CMFCActiveXControl1Ctrl::CMFCActiveXControl1CtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CMFCActiveXControl1Ctrl

BOOL CMFCActiveXControl1Ctrl::CMFCActiveXControl1CtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_MFCACTIVEXCONTROL1,
			IDB_MFCACTIVEXCONTROL1,
			afxRegApartmentThreading,
			_dwMFCActiveXControl1OleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


// CMFCActiveXControl1Ctrl::CMFCActiveXControl1Ctrl - Constructor

CMFCActiveXControl1Ctrl::CMFCActiveXControl1Ctrl()
{
	InitializeIIDs(&IID_DMFCActiveXControl1, &IID_DMFCActiveXControl1Events);

	EnableSimpleFrame();
	// TODO: Initialize your control's instance data here.
}

// CMFCActiveXControl1Ctrl::~CMFCActiveXControl1Ctrl - Destructor

CMFCActiveXControl1Ctrl::~CMFCActiveXControl1Ctrl()
{
	// TODO: Cleanup your control's instance data here.
}

// CMFCActiveXControl1Ctrl::OnDraw - Drawing function

void CMFCActiveXControl1Ctrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /* rcInvalid */)
{
	if (!pdc)
		return;

	// TODO: Replace the following code with your own drawing code.
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	pdc->Ellipse(rcBounds);
}

// CMFCActiveXControl1Ctrl::DoPropExchange - Persistence support

void CMFCActiveXControl1Ctrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
}


// CMFCActiveXControl1Ctrl::OnResetState - Reset control to default state

void CMFCActiveXControl1Ctrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


// CMFCActiveXControl1Ctrl::AboutBox - Display an "About" box to the user

void CMFCActiveXControl1Ctrl::AboutBox()
{
	CDialogEx dlgAbout(IDD_ABOUTBOX_MFCACTIVEXCONTROL1);
	dlgAbout.DoModal();
}


// CMFCActiveXControl1Ctrl message handlers
#define CF_EMBEDDEDOBJECT   TEXT("Embedded Object")
#define CF_OBJECTDESCRIPTOR TEXT("Object Descriptor")

//OnCopy() : Puts the object on the clipboard.
void CMFCActiveXControl1Ctrl::OnCopy(BOOL bActiveX)
{
	BeginWaitCursor();

	// Specify formats for which there is a default implementation
	if (CControlDataSource* pdata = new CControlDataSource(this)) {

		FORMATETC formatEtc;
		STGMEDIUM StgMedium = { 0 };

		if (bActiveX)
		{
			// Can render embedded object as persitent storage.
			formatEtc.cfFormat = ::RegisterClipboardFormat(CF_EMBEDDEDOBJECT);
			formatEtc.ptd = NULL;
			formatEtc.dwAspect = DVASPECT_CONTENT;
			formatEtc.lindex = -1;
			formatEtc.tymed = TYMED_ISTORAGE;
			pdata->DelayRenderData(0, &formatEtc);

			// Can render embedded object descriptor on global memory.
			formatEtc.cfFormat = ::RegisterClipboardFormat(CF_OBJECTDESCRIPTOR);
			formatEtc.ptd = NULL;
			formatEtc.dwAspect = DVASPECT_CONTENT;
			formatEtc.lindex = -1;
			formatEtc.tymed = TYMED_HGLOBAL;
			pdata->DelayRenderData(0, &formatEtc);
		}

		// Can render as windows metafile. Required by Office 95.
		formatEtc.cfFormat = CF_METAFILEPICT;
		formatEtc.ptd = NULL;
		formatEtc.dwAspect = DVASPECT_CONTENT;
		formatEtc.lindex = -1;
		formatEtc.tymed = TYMED_MFPICT;

		pdata->DelayRenderData(0, &formatEtc);

		pdata->SetClipboard();


	}
	EndWaitCursor();
}


BOOL CMFCActiveXControl1Ctrl::OnRenderData(LPFORMATETC pfmt, LPSTGMEDIUM pmedium)
{

	if (pfmt->cfFormat == ::RegisterClipboardFormat(CF_OBJECTDESCRIPTOR)) {
		if (pfmt->tymed & TYMED_HGLOBAL) {
			if (AllocateStgMedium(pmedium, TYMED_HGLOBAL, sizeof(OBJECTDESCRIPTOR) + 200)) {
				if (RenderObjectDescriptor(pmedium)) {
					return TRUE;
				}
			}
		}
	}

	if (pfmt->cfFormat == ::RegisterClipboardFormat(CF_EMBEDDEDOBJECT)) {
		if (pfmt->dwAspect == DVASPECT_CONTENT) {
			if (pfmt->tymed & TYMED_ISTORAGE) {
				if (AllocateStgMedium(pmedium, TYMED_ISTORAGE)) {
					if (RenderEmbeddedObject(pmedium)) {
						return TRUE;
					}
				}
			}
		}
	}

	return COleControl::OnRenderData(pfmt, pmedium);
}


BOOL CMFCActiveXControl1Ctrl::RenderEmbeddedObject(LPSTGMEDIUM pmedium)
{
	BOOL DataRendered = FALSE;

	IPersistStorage* ppersist = NULL;
	if (SUCCEEDED(ExternalQueryInterface(&IID_IPersistStorage, (LPVOID*)&ppersist))) {
		if (SUCCEEDED(OleSave(ppersist, pmedium->pstg, FALSE))) {
			DataRendered = TRUE;
		}
		ppersist->Release();
	}

	return DataRendered;
}


BOOL CMFCActiveXControl1Ctrl::RenderObjectDescriptor(LPSTGMEDIUM pmedium)
{
	BOOL DataRendered = FALSE;

	IOleObject* pobject = NULL;
	if (SUCCEEDED(ExternalQueryInterface(&IID_IOleObject, (LPVOID*)&pobject))) {

		CLSID clsid = CLSID_NULL;
		pobject->GetUserClassID(&clsid);

		SIZEL sizel = { 0, 0 };
		pobject->GetExtent(DVASPECT_CONTENT, &sizel);

		LPOLESTR pszUserType = NULL;
		ULONG    cbUserType = 0;
		if (SUCCEEDED(pobject->GetUserType(USERCLASSTYPE_SHORT, &pszUserType))) {
			cbUserType = (lstrlenW(pszUserType) + 1) * sizeof(OLECHAR);
		}

		LPOLESTR pszAppName = NULL;
		ULONG    cbAppName = 0;
		/*
		Do not provide an application name, because this is identical
		to the full user type name. If we do, everything appears twice
		in the Insert Object dialog box, which looks silly.

		if (SUCCEEDED (pobject->GetUserType (USERCLASSTYPE_APPNAME, &pszAppName))) {
		cbAppName = (lstrlenW(pszAppName) + 1) * sizeof (OLECHAR);
		}
		*/

		DWORD dwStatus = 0;
		pobject->GetMiscStatus(DVASPECT_CONTENT, &dwStatus);

		unsigned long cbSize = sizeof(OBJECTDESCRIPTOR) + cbUserType + cbAppName;

		if (cbSize <= GlobalSize(pmedium->hGlobal)) {
			if (OBJECTDESCRIPTOR* pdesc = (OBJECTDESCRIPTOR*)GlobalLock(pmedium->hGlobal)) {

				memset(pdesc, 0, cbSize);

				pdesc->cbSize = cbSize;
				pdesc->clsid = clsid;
				pdesc->sizel = sizel;
				pdesc->dwDrawAspect = DVASPECT_CONTENT;
				pdesc->dwStatus = dwStatus;

				if (cbUserType) {
					pdesc->dwFullUserTypeName = sizeof(OBJECTDESCRIPTOR);
					memcpy((BYTE*)pdesc + (pdesc->dwFullUserTypeName), pszUserType, cbUserType);
				}

				if (cbAppName) {
					pdesc->dwSrcOfCopy = sizeof(OBJECTDESCRIPTOR) + cbUserType;
					memcpy((BYTE*)pdesc + (pdesc->dwSrcOfCopy), pszAppName, cbAppName);
				}

				DataRendered = TRUE;
				GlobalUnlock(pmedium->hGlobal);
			}
		}

		CoTaskMemFree(pszUserType);
		CoTaskMemFree(pszAppName);
		pobject->Release();
	}

	return DataRendered;
}

BOOL CMFCActiveXControl1Ctrl::AllocateStgMedium(LPSTGMEDIUM pmedium, TYMED tymedWanted, DWORD cbSize)
{

	if (pmedium->tymed != TYMED_NULL) {
		// Caller (or the MFC framework) already allocated the
		// storage medium. This is a common optimization technique
		// in OLE Uniform Data Transfers.
		return (pmedium->tymed == DWORD(tymedWanted));
	}


	if (tymedWanted == TYMED_HGLOBAL) {
		if (HGLOBAL hglobal = GlobalAlloc(GMEM_SHARE, cbSize)) {
			pmedium->tymed = TYMED_HGLOBAL;
			pmedium->hGlobal = hglobal;
			pmedium->pUnkForRelease = NULL;
			return TRUE;
		}
	}


	if (tymedWanted == TYMED_ISTORAGE) {
		IStorage* pstgTemp = NULL;
		HRESULT hr = StgCreateDocfile(NULL, STGM_TRANSACTED |
			STGM_READWRITE | STGM_CREATE | STGM_SHARE_EXCLUSIVE |
			STGM_DELETEONRELEASE, 0, &pstgTemp);
		if (SUCCEEDED(hr)) {
			pmedium->tymed = TYMED_ISTORAGE;
			pmedium->pstg = pstgTemp;
			pmedium->pUnkForRelease = NULL;
			return TRUE;
		}
	}

	// Ran out of options. Debug version asserts, so we can add
	// whatever type of allocation is needed.
	DEBUG_ONLY(AfxAssertFailedLine(__FILE__, __LINE__));

	return FALSE;
}

void CMFCActiveXControl1Ctrl::OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	BOOL bControl = (BOOL)HIBYTE(::GetKeyState(VK_CONTROL));
	if (!bControl) {
		COleControl::OnKeyUp(nChar, nRepCnt, nFlags);
		return;
	}


	switch (nChar)
	{
	case	'C':
	case	VK_INSERT:
			OnCopy();
		break;
	}

	COleControl::OnKeyUp(nChar, nRepCnt, nFlags);
}