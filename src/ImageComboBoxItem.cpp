// ImageComboBoxItem.cpp: A wrapper for existing image combo box items.

#include "stdafx.h"
#include "ImageComboBoxItem.h"
#include "ClassFactory.h"
#include "CWindowEx2.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ImageComboBoxItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IImageComboBoxItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ImageComboBoxItem::Properties::~Properties()
{
	if(pOwnerICBox) {
		pOwnerICBox->Release();
	}
}

HWND ImageComboBoxItem::Properties::GetICBoxHWnd(void)
{
	ATLASSUME(pOwnerICBox);

	OLE_HANDLE handle = NULL;
	pOwnerICBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HWND ImageComboBoxItem::Properties::GetLBoxHWnd(void)
{
	ATLASSUME(pOwnerICBox);

	OLE_HANDLE handle = NULL;
	pOwnerICBox->get_hWndListBox(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

void ImageComboBoxItem::Properties::InvalidateComboBox(void)
{
	ATLASSUME(pOwnerICBox);

	OLE_HANDLE handle = NULL;
	pOwnerICBox->get_hWndComboBox(&handle);
	InvalidateRect(static_cast<HWND>(LongToHandle(handle)), NULL, TRUE);
}

void ImageComboBoxItem::Properties::InvalidateListBox(void)
{
	ATLASSUME(pOwnerICBox);

	OLE_HANDLE handle = NULL;
	pOwnerICBox->get_hWndListBox(&handle);
	InvalidateRect(static_cast<HWND>(LongToHandle(handle)), NULL, TRUE);
}


HRESULT ImageComboBoxItem::SaveState(int itemIndex, PCOMBOBOXEXITEM pTarget, HWND hWndICBox/* = NULL*/)
{
	if(!hWndICBox) {
		hWndICBox = properties.GetICBoxHWnd();
	}
	ATLASSERT(IsWindow(hWndICBox));
	if(itemIndex < 0 || itemIndex >= static_cast<int>(SendMessage(hWndICBox, CB_GETCOUNT, 0, 0))) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTarget, COMBOBOXEXITEM);
	if(!pTarget) {
		return E_POINTER;
	}

	ZeroMemory(pTarget, sizeof(COMBOBOXEXITEM));
	pTarget->iItem = itemIndex;
	// NOTE: cchTextMax includes the terminating null character
	pTarget->cchTextMax = static_cast<int>(SendMessage(hWndICBox, CB_GETLBTEXTLEN, properties.itemIndex, 0)) + 1;
	pTarget->pszText = reinterpret_cast<LPTSTR>(malloc(pTarget->cchTextMax * sizeof(TCHAR)));
	pTarget->mask = CBEIF_IMAGE | CBEIF_INDENT | CBEIF_LPARAM | CBEIF_OVERLAY | CBEIF_SELECTEDIMAGE | CBEIF_TEXT;

	return static_cast<HRESULT>(SendMessage(hWndICBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(pTarget)));
}


void ImageComboBoxItem::Attach(int itemIndex)
{
	properties.itemIndex = itemIndex;
}

void ImageComboBoxItem::Detach(void)
{
	properties.itemIndex = -2;
}

HRESULT ImageComboBoxItem::SaveState(PCOMBOBOXEXITEM pTarget, HWND hWndICBox/* = NULL*/)
{
	return SaveState(properties.itemIndex, pTarget, hWndICBox);
}

HRESULT ImageComboBoxItem::SaveState(VirtualImageComboBoxItem* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerICBox);
	COMBOBOXEXITEM item = {0};
	HRESULT hr = SaveState(&item);
	pTarget->LoadState(&item);
	SECUREFREE(item.pszText);

	return hr;
}

void ImageComboBoxItem::SetOwner(ImageComboBox* pOwner)
{
	if(properties.pOwnerICBox) {
		properties.pOwnerICBox->Release();
	}
	properties.pOwnerICBox = pOwner;
	if(properties.pOwnerICBox) {
		properties.pOwnerICBox->AddRef();
	}
}


STDMETHODIMP ImageComboBoxItem::get_Height(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	*pValue = static_cast<OLE_YSIZE_PIXELS>(SendMessage(hWndICBox, CB_GETITEMHEIGHT, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP ImageComboBoxItem::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	int itemIndex = properties.itemIndex;
	if(itemIndex == -1) {
		itemIndex = static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0));
	}

	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_IMAGE;
	if(SendMessage(hWndICBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::put_IconIndex(LONG newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_IMAGE;
	item.iImage = newValue;
	if(SendMessage(hWndICBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(properties.itemIndex == static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0))) {
			properties.InvalidateComboBox();
		}
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	*pValue = properties.pOwnerICBox->ItemIndexToID(properties.itemIndex);
	if(*pValue != -1) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::get_Indent(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	int itemIndex = properties.itemIndex;
	if(itemIndex == -1) {
		itemIndex = static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0));
	}

	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_INDENT;
	if(SendMessage(hWndICBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iIndent;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::put_Indent(OLE_XSIZE_PIXELS newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_INDENT;
	item.iIndent = newValue;
	if(SendMessage(hWndICBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex;
	return S_OK;
}

STDMETHODIMP ImageComboBoxItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	*pValue = static_cast<LONG>(SendMessage(hWndICBox, CB_GETITEMDATA, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP ImageComboBoxItem::put_ItemData(LONG newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	if(SendMessage(hWndICBox, CB_SETITEMDATA, properties.itemIndex, newValue) != CB_ERR) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::get_OverlayIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	int itemIndex = properties.itemIndex;
	if(itemIndex == -1) {
		itemIndex = static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0));
	}

	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_OVERLAY;
	if(SendMessage(hWndICBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iOverlay;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::put_OverlayIndex(LONG newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_OVERLAY;
	item.iOverlay = newValue;
	if(SendMessage(hWndICBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(properties.itemIndex == static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0))) {
			properties.InvalidateComboBox();
		}
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}
	if(properties.itemIndex == -1) {
		*pValue = VARIANT_TRUE;
		return S_OK;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	*pValue = BOOL2VARIANTBOOL(properties.itemIndex == static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0)));
	return S_OK;
}

STDMETHODIMP ImageComboBoxItem::get_SelectedIconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	int itemIndex = properties.itemIndex;
	if(itemIndex == -1) {
		itemIndex = static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0));
	}

	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_SELECTEDIMAGE;
	if(SendMessage(hWndICBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iSelectedImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::put_SelectedIconIndex(LONG newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_SELECTEDIMAGE;
	item.iSelectedImage = newValue;
	if(SendMessage(hWndICBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(properties.itemIndex == static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0))) {
			properties.InvalidateComboBox();
		}
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	int itemIndex = properties.itemIndex;

	if(properties.itemIndex == -1 && (CWindow(hWndICBox).GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
		itemIndex = static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0));
		if(itemIndex == -1) {
			*pValue = SysAllocString(L"");
			return S_OK;
		}
	}

	int textLength = static_cast<int>(SendMessage(hWndICBox, CB_GETLBTEXTLEN, itemIndex, 0));
	if(textLength == CB_ERR) {
		return DISP_E_BADINDEX;
	}
	LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
	if(!pBuffer) {
		return E_OUTOFMEMORY;
	}
	ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
	if(SendMessage(hWndICBox, CB_GETLBTEXT, itemIndex, reinterpret_cast<LPARAM>(pBuffer)) != CB_ERR) {
		*pValue = _bstr_t(pBuffer).Detach();
	}
	HeapFree(GetProcessHeap(), 0, pBuffer);
	return S_OK;
}

STDMETHODIMP ImageComboBoxItem::put_Text(BSTR newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_TEXT;
	COLE2T converter(newValue);
	if(newValue) {
		item.pszText = converter;
	} else {
		item.pszText = LPSTR_TEXTCALLBACK;
	}
	if(SendMessage(hWndICBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(properties.itemIndex == static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0))) {
			properties.InvalidateComboBox();
		}
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}


STDMETHODIMP ImageComboBoxItem::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	ATLASSUME(properties.pOwnerICBox);

	POINT upperLeftPoint = {0};
	*phImageList = HandleToLong(properties.pOwnerICBox->CreateLegacyDragImage(properties.itemIndex, &upperLeftPoint, NULL));

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

STDMETHODIMP ImageComboBoxItem::GetRectangle(ItemRectangleTypeConstants /*rectangleType*/, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
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