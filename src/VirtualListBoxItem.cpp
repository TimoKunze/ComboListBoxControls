// VirtualListBoxItem.cpp: A wrapper for non-existing list box items.

#include "stdafx.h"
#include "VirtualListBoxItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualListBoxItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualListBoxItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualListBoxItem::Properties::~Properties()
{
	if(dataIsString && data) {
		free(reinterpret_cast<LPVOID>(data));
	}
	if(pOwnerLBox) {
		pOwnerLBox->Release();
	}
}

HWND VirtualListBoxItem::Properties::GetLBoxHWnd(void)
{
	ATLASSUME(pOwnerLBox);

	OLE_HANDLE handle = NULL;
	pOwnerLBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT VirtualListBoxItem::LoadState(int itemIndex, LPARAM data, LONG itemData)
{
	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	if(properties.dataIsString && properties.data) {
		free(reinterpret_cast<LPVOID>(properties.data));
	}
	properties.data = 0;

	DWORD style = CWindow(hWndLBox).GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		properties.dataIsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		properties.dataIsString = TRUE;
	}

	if(properties.dataIsString) {
		// duplicate the item's text
		int length = lstrlen(reinterpret_cast<LPCTSTR>(data));
		properties.data = reinterpret_cast<LPARAM>(malloc((length + 1) * sizeof(TCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopy(reinterpret_cast<LPTSTR>(properties.data), length + 1, reinterpret_cast<LPCTSTR>(data))));
	} else {
		properties.data = data;
	}
	properties.itemIndex = itemIndex;
	properties.itemData = itemData;

	return S_OK;
}

HRESULT VirtualListBoxItem::SaveState(int& itemIndex, LPARAM& data, LONG& itemData)
{
	if(properties.dataIsString) {
		// duplicate the item's text
		int length = lstrlen(reinterpret_cast<LPCTSTR>(properties.data));
		data = reinterpret_cast<LPARAM>(malloc((length + 1) * sizeof(TCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopy(reinterpret_cast<LPTSTR>(data), length + 1, reinterpret_cast<LPCTSTR>(properties.data))));
	} else {
		data = properties.data;
	}
	itemIndex = properties.itemIndex;
	itemData = properties.itemData;

	return S_OK;
}

void VirtualListBoxItem::SetOwner(ListBox* pOwner)
{
	if(properties.pOwnerLBox) {
		properties.pOwnerLBox->Release();
	}
	properties.pOwnerLBox = pOwner;
	if(properties.pOwnerLBox) {
		properties.pOwnerLBox->AddRef();
	}
}


STDMETHODIMP VirtualListBoxItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex;
	return S_OK;
}

STDMETHODIMP VirtualListBoxItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemData;
	return S_OK;
}

STDMETHODIMP VirtualListBoxItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.dataIsString) {
		if(properties.data) {
			*pValue = _bstr_t(reinterpret_cast<LPCTSTR>(properties.data)).Detach();
		}
	} else {
		*pValue = SysAllocString(L"");
	}

	return S_OK;
}