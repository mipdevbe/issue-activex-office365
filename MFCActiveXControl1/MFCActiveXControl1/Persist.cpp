#include "StdAfx.h"
#include <afxole.h>
#include <stdlib.h> 
#include <wtypes.h>
#include "MFCActiveXControl1.h"
#include "MFCActiveXControl1Ctrl.h"

static LPCOLESTR s_szContents = OLESTR("Contents");
static char s_szVersion[] = "Version";
//---------------------------------------------------------------------------------------------
ULONG EXPORT CMFCActiveXControl1Ctrl::XObjSafe::AddRef()
{
	METHOD_PROLOGUE(CMFCActiveXControl1Ctrl, ObjSafe)

	return pThis->ExternalAddRef();
}

ULONG EXPORT CMFCActiveXControl1Ctrl::XObjSafe::Release()
{
	METHOD_PROLOGUE(CMFCActiveXControl1Ctrl, ObjSafe)

	return pThis->ExternalRelease();
}

HRESULT EXPORT CMFCActiveXControl1Ctrl::XObjSafe::QueryInterface(REFIID iid, void** ppvObj)
{
	METHOD_PROLOGUE(CMFCActiveXControl1Ctrl, ObjSafe)

	return (HRESULT)pThis->ExternalQueryInterface(&iid, ppvObj);
}


STDMETHODIMP_(ULONG) CMFCActiveXControl1Ctrl::XPersistStorage::AddRef()
{
	METHOD_PROLOGUE_EX_(CMFCActiveXControl1Ctrl, PersistStorage)
		return (ULONG)pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CMFCActiveXControl1Ctrl::XPersistStorage::Release()
{
	METHOD_PROLOGUE_EX_(CMFCActiveXControl1Ctrl, PersistStorage)
		return (ULONG)pThis->ExternalRelease();
}

STDMETHODIMP CMFCActiveXControl1Ctrl::XPersistStorage::QueryInterface( REFIID iid, LPVOID* ppvObj)
{
	METHOD_PROLOGUE_EX_(CMFCActiveXControl1Ctrl, PersistStorage)
		return (HRESULT)pThis->ExternalQueryInterface(&iid, ppvObj);
}

const DWORD dwSupportedBits = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
const DWORD dwNotSupportedBits = ~dwSupportedBits;

HRESULT STDMETHODCALLTYPE CMFCActiveXControl1Ctrl::XObjSafe::GetInterfaceSafetyOptions(
	/* [in] */ REFIID riid,
	/* [out]*/ DWORD __RPC_FAR *pdwSupportedOptions,
	/* [out]*/ DWORD __RPC_FAR *pdwEnabledOptions)
{
	METHOD_PROLOGUE(CMFCActiveXControl1Ctrl, ObjSafe)

		HRESULT retval = ResultFromScode(S_OK);
	/* does the interface exist ? */
	IUnknown* pUnkInterface = NULL;
	retval = pThis->ExternalQueryInterface(&riid, (void **)&pUnkInterface);
	if (retval != E_NOINTERFACE) {
		/* Interface exists */
		pUnkInterface->Release(); /* Release it just checking */
	}

	*pdwSupportedOptions = *pdwEnabledOptions = dwSupportedBits;
	return retval;
}

HRESULT STDMETHODCALLTYPE CMFCActiveXControl1Ctrl::XObjSafe::SetInterfaceSafetyOptions(
	/* [in] */ REFIID riid,
	/* [in] */ DWORD dwOptionSetMask,
	/* [in] */ DWORD dwEnabledOptions)
{
	METHOD_PROLOGUE(CMFCActiveXControl1Ctrl, ObjSafe)

	/* does the interface exist ? */
	IUnknown* pUnkInterface = NULL;
	pThis->ExternalQueryInterface(&riid,(void **)&pUnkInterface);
	if (pUnkInterface) {
		/* Interface exists */
		pUnkInterface->Release(); /* Release it just checking */
	}
	else {
		return ResultFromScode(E_NOINTERFACE);
	}

	if (dwOptionSetMask & dwNotSupportedBits) {
		return ResultFromScode(E_FAIL);
	}

	dwEnabledOptions &= dwSupportedBits;
	if ((dwOptionSetMask & dwEnabledOptions) != dwOptionSetMask) {
		return ResultFromScode(E_FAIL);
	}

	return ResultFromScode(S_OK);
}

//---------------------------------------------------------------------------------------------

STDMETHODIMP CMFCActiveXControl1Ctrl::XPersistStorage::GetClassID(LPCLSID lpClassID)
{
	METHOD_PROLOGUE_EX_(CMFCActiveXControl1Ctrl, PersistStorage)
		return pThis->GetClassID(lpClassID);
}

STDMETHODIMP CMFCActiveXControl1Ctrl::XPersistStorage::IsDirty()
{
	METHOD_PROLOGUE_EX_(CMFCActiveXControl1Ctrl, PersistStorage)

		if (pThis->m_pDefIPersistStorage == NULL)
			pThis->m_pDefIPersistStorage =
			(LPPERSISTSTORAGE)pThis->QueryDefHandler(IID_IPersistStorage);

	BOOL bDefModified = (pThis->m_pDefIPersistStorage->IsDirty() == S_OK);
	return (bDefModified || pThis->m_bModified) ? S_OK : S_FALSE;

}

STDMETHODIMP CMFCActiveXControl1Ctrl::XPersistStorage::InitNew(LPSTORAGE pStg)
{
	METHOD_PROLOGUE_EX(CMFCActiveXControl1Ctrl, PersistStorage)

		if (pThis->m_pDefIPersistStorage == NULL)
			pThis->m_pDefIPersistStorage =
			(LPPERSISTSTORAGE)pThis->QueryDefHandler(IID_IPersistStorage);

	pThis->m_pDefIPersistStorage->InitNew(pStg);

	// Delegate to OnResetState.
	pThis->OnResetState();

	// Unless IOleObject::SetClientSite is called after this, we can
	// count on ambient properties being available while loading.
	pThis->m_bCountOnAmbients = TRUE;

	// Properties have been initialized
	pThis->m_bInitialized = TRUE;

	return S_OK;
}

STDMETHODIMP CMFCActiveXControl1Ctrl::XPersistStorage::Load(LPSTORAGE pStg)
{
	ASSERT(pStg != NULL);
	METHOD_PROLOGUE_EX(CMFCActiveXControl1Ctrl, PersistStorage)
		pThis->ExternalAddRef();

	LPSTREAM pStm = NULL;
	HRESULT hr = pStg->OpenStream(s_szContents, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStm);

	ASSERT(FAILED(hr) || pStm != NULL);

	if (pStm != NULL)
	{
		char szVersion[1024];
		ULONG cbWritten = 0;
		pStm->Read(szVersion, sizeof(szVersion), &cbWritten);
		szVersion[sizeof(szVersion) - 1] = 0x0;
		pStm->Release();
		pStm = NULL;

		::MessageBoxA(NULL, szVersion, s_szVersion, MB_OK);
	}


	pThis->ExternalRelease();

	return S_OK;
}

STDMETHODIMP CMFCActiveXControl1Ctrl::XPersistStorage::Save(LPSTORAGE pStg,
	BOOL fSameAsLoad)
{

	METHOD_PROLOGUE_EX(CMFCActiveXControl1Ctrl, PersistStorage)
		ASSERT(pStg != NULL);

	pThis->ExternalAddRef();

	LPSTREAM pStm = NULL;
	HRESULT hr = pStg->CreateStream(s_szContents, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, &pStm);

	ASSERT(FAILED(hr) || pStm != NULL);

	if (pStm != NULL)
	{
		char szVersion[1024];
		sprintf_s(szVersion, "%s:03102018", s_szVersion);
		ULONG cbWritten = 0;
		pStm->Write(szVersion, sizeof(szVersion), &cbWritten);
		pStm->Commit(STGC_DEFAULT);
		pStm->Release();
		pStm = NULL;
	}

	// Delegate to default handler (for cache).

	hr = pStg->Commit(STGC_DEFAULT);

	pThis->ExternalRelease();

	return hr;
}

STDMETHODIMP CMFCActiveXControl1Ctrl::XPersistStorage::SaveCompleted(LPSTORAGE pStgSaved)
{
	METHOD_PROLOGUE_EX_(CMFCActiveXControl1Ctrl, PersistStorage)

		if (pThis->m_pDefIPersistStorage == NULL)
			pThis->m_pDefIPersistStorage =
			(LPPERSISTSTORAGE)pThis->QueryDefHandler(IID_IPersistStorage);

	if (pStgSaved != NULL)
		pThis->SetModifiedFlag(FALSE);

	return pThis->m_pDefIPersistStorage->SaveCompleted(pStgSaved);

}

STDMETHODIMP CMFCActiveXControl1Ctrl::XPersistStorage::HandsOffStorage()
{
	METHOD_PROLOGUE_EX_(CMFCActiveXControl1Ctrl, PersistStorage)

		if (pThis->m_pDefIPersistStorage == NULL)
			pThis->m_pDefIPersistStorage =
			(LPPERSISTSTORAGE)pThis->QueryDefHandler(IID_IPersistStorage);

	return pThis->m_pDefIPersistStorage->HandsOffStorage();
}

//---------------------------------------------------------------------------------------------
