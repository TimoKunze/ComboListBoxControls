// VirtualComboBoxItem.cpp: A wrapper for non-existing combo box items.

#include "stdafx.h"
#include "VirtualComboBoxItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualComboBoxItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualComboBoxItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualComboBoxItem::Properties::~Properties()
{
	if(dataIsString && data) {
		free(reinterpret_cast<LPVOID>(data));
	}
	if(pOwnerCBox) {
		pOwnerCBox->Release();
	}
}

HWND VirtualComboBoxItem::Properties::GetCBoxHWnd(void)
{
	ATLASSUME(pOwnerCBox);

	OLE_HANDLE handle = NULL;
	pOwnerCBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT VirtualComboBoxItem::LoadState(int itemIndex, LPARAM data, LONG itemData)
{
	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));

	if(properties.dataIsString && properties.data) {
		free(reinterpret_cast<LPVOID>(properties.data));
	}
	properties.data = 0;

	DWORD style = CWindow(hWndCBox).GetStyle();
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		properties.dataIsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
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

HRESULT VirtualComboBoxItem::SaveState(int& itemIndex, LPARAM& data, LONG& itemData)
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

void VirtualComboBoxItem::SetOwner(ComboBox* pOwner)
{
	if(properties.pOwnerCBox) {
		properties.pOwnerCBox->Release();
	}
	properties.pOwnerCBox = pOwner;
	if(properties.pOwnerCBox) {
		properties.pOwnerCBox->AddRef();
	}
}


STDMETHODIMP VirtualComboBoxItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex;
	return S_OK;
}

STDMETHODIMP VirtualComboBoxItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemData;
	return S_OK;
}

STDMETHODIMP VirtualComboBoxItem::get_Text(BSTR* pValue)
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