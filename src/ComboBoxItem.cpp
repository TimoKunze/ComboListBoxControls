// ComboBoxItem.cpp: A wrapper for existing combo box items.

#include "stdafx.h"
#include "ComboBoxItem.h"
#include "ClassFactory.h"
#include "CWindowEx2.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ComboBoxItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IComboBoxItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ComboBoxItem::Properties::~Properties()
{
	if(pOwnerCBox) {
		pOwnerCBox->Release();
	}
}

HWND ComboBoxItem::Properties::GetCBoxHWnd(void)
{
	ATLASSUME(pOwnerCBox);

	OLE_HANDLE handle = NULL;
	pOwnerCBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HWND ComboBoxItem::Properties::GetLBoxHWnd(void)
{
	ATLASSUME(pOwnerCBox);

	OLE_HANDLE handle = NULL;
	pOwnerCBox->get_hWndListBox(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ComboBoxItem::Attach(int itemIndex, ULONG_PTR itemData/* = NULL*/)
{
	properties.itemIndex = itemIndex;
	properties.itemData = itemData;
}

void ComboBoxItem::Detach(void)
{
	properties.itemIndex = -2;
}

HRESULT ComboBoxItem::SaveState(VirtualComboBoxItem* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));

	BOOL dataIsString;
	DWORD style = CWindow(hWndCBox).GetStyle();
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		dataIsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
	} else {
		dataIsString = TRUE;
	}

	LPARAM data = 0;
	if(dataIsString) {
		int textLength = static_cast<int>(SendMessage(hWndCBox, CB_GETLBTEXTLEN, properties.itemIndex, 0));
		if(textLength == CB_ERR) {
			return DISP_E_BADINDEX;
		}
		data = reinterpret_cast<LPARAM>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
		if(!data) {
			return E_OUTOFMEMORY;
		}
		ZeroMemory(reinterpret_cast<LPVOID>(data), (textLength + 1) * sizeof(TCHAR));
		SendMessage(hWndCBox, CB_GETLBTEXT, properties.itemIndex, data);
	} else {
		SendMessage(hWndCBox, CB_GETLBTEXT, properties.itemIndex, reinterpret_cast<LPARAM>(&data));
	}

	pTarget->SetOwner(properties.pOwnerCBox);
	LONG itemData = 0;
	get_ItemData(&itemData);
	pTarget->LoadState(properties.itemIndex, data, itemData);
	if(dataIsString) {
		HeapFree(GetProcessHeap(), 0, reinterpret_cast<LPVOID>(data));
	}
	return S_OK;
}

void ComboBoxItem::SetOwner(ComboBox* pOwner)
{
	if(properties.pOwnerCBox) {
		properties.pOwnerCBox->Release();
	}
	properties.pOwnerCBox = pOwner;
	if(properties.pOwnerCBox) {
		properties.pOwnerCBox->AddRef();
	}
}


STDMETHODIMP ComboBoxItem::get_Height(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));

	*pValue = static_cast<OLE_YSIZE_PIXELS>(SendMessage(hWndCBox, CB_GETITEMHEIGHT, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP ComboBoxItem::put_Height(LONG newValue)
{
	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));

	if(SendMessage(hWndCBox, CB_SETITEMHEIGHT, properties.itemIndex, newValue) != CB_ERR) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBoxItem::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	if(properties.pOwnerCBox) {
		*pValue = properties.pOwnerCBox->ItemIndexToID(properties.itemIndex);
		if(*pValue != -1) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ComboBoxItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex;
	return S_OK;
}

STDMETHODIMP ComboBoxItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}
	if(properties.itemIndex == -1) {
		*pValue = static_cast<LONG>(properties.itemData);
		return S_OK;
	}

	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));

	*pValue = static_cast<LONG>(SendMessage(hWndCBox, CB_GETITEMDATA, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP ComboBoxItem::put_ItemData(LONG newValue)
{
	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));

	if(SendMessage(hWndCBox, CB_SETITEMDATA, properties.itemIndex, newValue) != CB_ERR) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBoxItem::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		*pValue = VARIANT_FALSE;
		return S_OK;
	}

	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));

	*pValue = BOOL2VARIANTBOOL(properties.itemIndex == static_cast<int>(SendMessage(hWndCBox, CB_GETCURSEL, 0, 0)));
	return S_OK;
}

STDMETHODIMP ComboBoxItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));

	BOOL expectsString;
	DWORD style = CWindow(hWndCBox).GetStyle();
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	if(!expectsString) {
		*pValue = SysAllocString(L"");
		return S_OK;
	}

	int textLength = static_cast<int>(SendMessage(hWndCBox, CB_GETLBTEXTLEN, properties.itemIndex, 0));
	if(textLength == CB_ERR) {
		return DISP_E_BADINDEX;
	}
	LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
	if(!pBuffer) {
		return E_OUTOFMEMORY;
	}
	ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
	SendMessage(hWndCBox, CB_GETLBTEXT, properties.itemIndex, reinterpret_cast<LPARAM>(pBuffer));

	*pValue = _bstr_t(pBuffer).Detach();
	HeapFree(GetProcessHeap(), 0, pBuffer);
	return S_OK;
}

STDMETHODIMP ComboBoxItem::put_Text(BSTR newValue)
{
	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndCBox = properties.GetCBoxHWnd();
	ATLASSERT(IsWindow(hWndCBox));
	CWindowEx2 wnd = hWndCBox;
	BOOL expectsString;
	DWORD style = wnd.GetStyle();
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	if(!expectsString) {
		return E_FAIL;
	}

	COLE2T converter(newValue);
	LPTSTR pBuffer = converter;

	// there really is no CB_SETTEXT
	HRESULT hr = E_FAIL;
	wnd.InternalSetRedraw(FALSE);
	int currentSelection = static_cast<int>(SendMessage(hWndCBox, CB_GETCURSEL, 0, 0));
	LPARAM itemData = static_cast<LPARAM>(SendMessage(hWndCBox, CB_GETITEMDATA, properties.itemIndex, 0));

	LONG id = 0;
	get_ID(&id);

	properties.pOwnerCBox->EnterSilentItemInsertionSection();
	if(wnd.GetStyle() & CBS_SORT) {
		// CB_INSERTSTRING won't resort the list
		int newIndex = static_cast<int>(SendMessage(hWndCBox, CB_ADDSTRING, properties.itemIndex, reinterpret_cast<LPARAM>(pBuffer)));
		if(newIndex > CB_ERR) {
			hr = S_OK;
			int indexToDelete = properties.pOwnerCBox->IDToItemIndex(id);
			ATLASSERT(indexToDelete != newIndex);
			properties.pOwnerCBox->EnterSilentItemDeletionSection();
			SendMessage(hWndCBox, CB_DELETESTRING, indexToDelete, 0);
			properties.pOwnerCBox->LeaveSilentItemDeletionSection();
			if(newIndex > indexToDelete) {
				newIndex--;
			}
			properties.itemIndex = newIndex;
		}
	} else {
		if(static_cast<int>(SendMessage(hWndCBox, CB_INSERTSTRING, properties.itemIndex, reinterpret_cast<LPARAM>(pBuffer))) > CB_ERR) {
			hr = S_OK;
			properties.pOwnerCBox->EnterSilentItemDeletionSection();
			SendMessage(hWndCBox, CB_DELETESTRING, properties.itemIndex + 1, 0);
			properties.pOwnerCBox->LeaveSilentItemDeletionSection();
		}
	}
	properties.pOwnerCBox->LeaveSilentItemInsertionSection();
	if(SUCCEEDED(hr)) {
		properties.pOwnerCBox->SetItemID(properties.itemIndex, id);
		properties.pOwnerCBox->DecrementNextItemID();
	}

	if(SUCCEEDED(hr)) {
		// NOTE: itemData can be CB_ERR
		SendMessage(hWndCBox, CB_SETITEMDATA, properties.itemIndex, itemData);
		SendMessage(hWndCBox, CB_SETCURSEL, currentSelection, 0);
	}
	wnd.InternalSetRedraw(TRUE);
	return hr;
}


STDMETHODIMP ComboBoxItem::GetRectangle(ItemRectangleTypeConstants /*rectangleType*/, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	HRESULT hr = E_FAIL;
	RECT rc = {0};
	if(SendMessage(hWndLBox, LB_GETITEMRECT, properties.itemIndex, reinterpret_cast<LPARAM>(&rc)) != LB_ERR) {
		hr = S_OK;
	}
	if(SUCCEEDED(hr)) {
		if(pXLeft) {
			*pXLeft = rc.left;
		}
		if(pYTop) {
			*pYTop = rc.top;
		}
		if(pXRight) {
			*pXRight = rc.right;
		}
		if(pYBottom) {
			*pYBottom = rc.bottom;
		}
	}
	return hr;
}