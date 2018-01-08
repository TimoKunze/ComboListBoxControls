// ListBoxItem.cpp: A wrapper for existing list box items.

#include "stdafx.h"
#include "ListBoxItem.h"
#include "ClassFactory.h"
#include "CWindowEx.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListBoxItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListBoxItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ListBoxItem::Properties::~Properties()
{
	if(pOwnerLBox) {
		pOwnerLBox->Release();
	}
}

HWND ListBoxItem::Properties::GetLBoxHWnd(void)
{
	ATLASSUME(pOwnerLBox);

	OLE_HANDLE handle = NULL;
	pOwnerLBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ListBoxItem::Attach(int itemIndex, ULONG_PTR itemData/* = NULL*/)
{
	properties.itemIndex = itemIndex;
	properties.itemData = itemData;
}

void ListBoxItem::Detach(void)
{
	properties.itemIndex = -2;
}

HRESULT ListBoxItem::SaveState(VirtualListBoxItem* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	BOOL dataIsString;
	DWORD style = CWindow(hWndLBox).GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		dataIsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		dataIsString = TRUE;
	}

	LPARAM data = 0;
	if(dataIsString) {
		int textLength = static_cast<int>(SendMessage(hWndLBox, LB_GETTEXTLEN, properties.itemIndex, 0));
		if(textLength == LB_ERR) {
			return DISP_E_BADINDEX;
		}
		data = reinterpret_cast<LPARAM>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
		if(!data) {
			return E_OUTOFMEMORY;
		}
		SendMessage(hWndLBox, LB_GETTEXT, properties.itemIndex, data);
	} else {
		SendMessage(hWndLBox, LB_GETTEXT, properties.itemIndex, reinterpret_cast<LPARAM>(&data));
	}

	pTarget->SetOwner(properties.pOwnerLBox);
	LONG itemData = 0;
	get_ItemData(&itemData);
	pTarget->LoadState(properties.itemIndex, data, itemData);
	if(dataIsString) {
		HeapFree(GetProcessHeap(), 0, reinterpret_cast<LPVOID>(data));
	}
	return S_OK;
}

void ListBoxItem::SetOwner(ListBox* pOwner)
{
	if(properties.pOwnerLBox) {
		properties.pOwnerLBox->Release();
	}
	properties.pOwnerLBox = pOwner;
	if(properties.pOwnerLBox) {
		properties.pOwnerLBox->AddRef();
	}
}


STDMETHODIMP ListBoxItem::get_Anchor(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	*pValue = BOOL2VARIANTBOOL(properties.itemIndex == static_cast<int>(SendMessage(hWndLBox, LB_GETANCHORINDEX, 0, 0)));
	return S_OK;
}

STDMETHODIMP ListBoxItem::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	*pValue = BOOL2VARIANTBOOL(properties.itemIndex == static_cast<int>(SendMessage(hWndLBox, LB_GETCARETINDEX, 0, 0)));
	return S_OK;
}

STDMETHODIMP ListBoxItem::get_Height(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	*pValue = static_cast<OLE_YSIZE_PIXELS>(SendMessage(hWndLBox, LB_GETITEMHEIGHT, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP ListBoxItem::put_Height(LONG newValue)
{
	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	if(SendMessage(hWndLBox, LB_SETITEMHEIGHT, properties.itemIndex, newValue) != LB_ERR) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBoxItem::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	if(properties.pOwnerLBox) {
		*pValue = properties.pOwnerLBox->ItemIndexToID(properties.itemIndex);
		if(*pValue != -1) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListBoxItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex;
	return S_OK;
}

STDMETHODIMP ListBoxItem::get_ItemData(LONG* pValue)
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

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	*pValue = static_cast<LONG>(SendMessage(hWndLBox, LB_GETITEMDATA, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP ListBoxItem::put_ItemData(LONG newValue)
{
	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	if(SendMessage(hWndLBox, LB_SETITEMDATA, properties.itemIndex, newValue) != LB_ERR) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBoxItem::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		*pValue = VARIANT_FALSE;
		return S_OK;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	*pValue = BOOL2VARIANTBOOL(SendMessage(hWndLBox, LB_GETSEL, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP ListBoxItem::put_Selected(VARIANT_BOOL newValue)
{
	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	CWindowEx wnd = hWndLBox;
	if(!(wnd.GetStyle() & (LBS_MULTIPLESEL | LBS_EXTENDEDSEL))) {
		return E_FAIL;
	}

	SendMessage(hWndLBox, LB_SETSEL, VARIANTBOOL2BOOL(newValue), properties.itemIndex);
	return S_OK;
}

STDMETHODIMP ListBoxItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	BOOL expectsString;
	DWORD style = CWindow(hWndLBox).GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	if(!expectsString) {
		*pValue = SysAllocString(L"");
		return S_OK;
	}

	int textLength = static_cast<int>(SendMessage(hWndLBox, LB_GETTEXTLEN, properties.itemIndex, 0));
	if(textLength == LB_ERR) {
		return DISP_E_BADINDEX;
	}
	LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
	if(!pBuffer) {
		return E_OUTOFMEMORY;
	}
	SendMessage(hWndLBox, LB_GETTEXT, properties.itemIndex, reinterpret_cast<LPARAM>(pBuffer));

	*pValue = _bstr_t(pBuffer).Detach();
	HeapFree(GetProcessHeap(), 0, pBuffer);
	return S_OK;
}

STDMETHODIMP ListBoxItem::put_Text(BSTR newValue)
{
	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));
	CWindowEx wnd = hWndLBox;
	BOOL expectsString;
	DWORD style = wnd.GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	if(!expectsString) {
		return E_FAIL;
	}

	COLE2T converter(newValue);
	LPTSTR pBuffer = converter;

	// there really is no LB_SETTEXT
	HRESULT hr = E_FAIL;
	wnd.InternalSetRedraw(FALSE);
	int currentSelection = static_cast<int>(SendMessage(hWndLBox, LB_GETCURSEL, 0, 0));
	LPARAM itemData = static_cast<LPARAM>(SendMessage(hWndLBox, LB_GETITEMDATA, properties.itemIndex, 0));

	LONG id = 0;
	get_ID(&id);

	properties.pOwnerLBox->EnterSilentItemInsertionSection();
	if(wnd.GetStyle() & LBS_SORT) {
		// LB_INSERTSTRING won't resort the list
		int newIndex = static_cast<int>(SendMessage(hWndLBox, LB_ADDSTRING, properties.itemIndex, reinterpret_cast<LPARAM>(pBuffer)));
		if(newIndex > LB_ERR) {
			hr = S_OK;
			int indexToDelete = properties.pOwnerLBox->IDToItemIndex(id);
			ATLASSERT(indexToDelete != newIndex);
			properties.pOwnerLBox->EnterSilentItemDeletionSection();
			SendMessage(hWndLBox, LB_DELETESTRING, indexToDelete, 0);
			properties.pOwnerLBox->LeaveSilentItemDeletionSection();
			if(newIndex > indexToDelete) {
				newIndex--;
			}
			properties.itemIndex = newIndex;
		}
	} else {
		if(static_cast<int>(SendMessage(hWndLBox, LB_INSERTSTRING, properties.itemIndex, reinterpret_cast<LPARAM>(pBuffer))) > LB_ERR) {
			hr = S_OK;
			properties.pOwnerLBox->EnterSilentItemDeletionSection();
			SendMessage(hWndLBox, LB_DELETESTRING, properties.itemIndex + 1, 0);
			properties.pOwnerLBox->LeaveSilentItemDeletionSection();
		}
	}
	properties.pOwnerLBox->LeaveSilentItemInsertionSection();
	if(SUCCEEDED(hr)) {
		properties.pOwnerLBox->SetItemID(properties.itemIndex, id);
		properties.pOwnerLBox->DecrementNextItemID();
	}

	if(SUCCEEDED(hr)) {
		// NOTE: itemData can be LB_ERR
		SendMessage(hWndLBox, LB_SETITEMDATA, properties.itemIndex, itemData);
		SendMessage(hWndLBox, LB_SETCURSEL, currentSelection, 0);
	}
	wnd.InternalSetRedraw(TRUE);
	return hr;
}


STDMETHODIMP ListBoxItem::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	ATLASSUME(properties.pOwnerLBox);

	POINT upperLeftPoint = {0};
	*phImageList = HandleToLong(properties.pOwnerLBox->CreateLegacyDragImage(properties.itemIndex, &upperLeftPoint, NULL));

	if(*phImageList) {
		if(pXUpperLeft) {
			*pXUpperLeft = upperLeftPoint.x;
		}
		if(pYUpperLeft) {
			*pYUpperLeft = upperLeftPoint.y;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBoxItem::GetRectangle(ItemRectangleTypeConstants /*rectangleType*/, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
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