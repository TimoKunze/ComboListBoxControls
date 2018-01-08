// APIWrapper.cpp: A wrapper class for API functions not available on all supported systems

#include "stdafx.h"
#include "APIWrapper.h"


BOOL APIWrapper::checkedSupport_SHGetImageList = FALSE;
BOOL APIWrapper::checkedSupport_SHGetKnownFolderIDList = FALSE;
SHGetImageListFn* APIWrapper::pfnSHGetImageList = NULL;
SHGetKnownFolderIDListFn* APIWrapper::pfnSHGetKnownFolderIDList = NULL;


APIWrapper::APIWrapper(void)
{
}


BOOL APIWrapper::IsSupported_SHGetImageList(void)
{
	if(!checkedSupport_SHGetImageList) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHGetImageList = reinterpret_cast<SHGetImageListFn*>(GetProcAddress(hShell32, "SHGetImageList"));
			if(!pfnSHGetImageList) {
				pfnSHGetImageList = reinterpret_cast<SHGetImageListFn*>(GetProcAddress(hShell32, MAKEINTRESOURCEA(727)));
			}
		}
		checkedSupport_SHGetImageList = TRUE;
	}

	return (pfnSHGetImageList != NULL);
}

BOOL APIWrapper::IsSupported_SHGetKnownFolderIDList(void)
{
	if(!checkedSupport_SHGetKnownFolderIDList) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHGetKnownFolderIDList = reinterpret_cast<SHGetKnownFolderIDListFn*>(GetProcAddress(hShell32, "SHGetKnownFolderIDList"));
		}
		checkedSupport_SHGetKnownFolderIDList = TRUE;
	}

	return (pfnSHGetKnownFolderIDList != NULL);
}

HRESULT APIWrapper::SHGetImageList(int imageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHGetImageList()) {
		*pReturnValue = pfnSHGetImageList(imageList, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHGetKnownFolderIDList(REFKNOWNFOLDERID folderID, DWORD flags, HANDLE hToken, PIDLIST_ABSOLUTE* ppIDL, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHGetKnownFolderIDList()) {
		*pReturnValue = pfnSHGetKnownFolderIDList(folderID, flags, hToken, ppIDL);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}