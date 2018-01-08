// DriveComboBoxItem.cpp: A wrapper for existing drive combo box items.

#include "stdafx.h"
#include "DriveComboBoxItem.h"
#include "ClassFactory.h"
#include "CWindowEx.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP DriveComboBoxItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IDriveComboBoxItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


DriveComboBoxItem::Properties::~Properties()
{
	if(pOwnerDCBox) {
		pOwnerDCBox->Release();
	}
}

HWND DriveComboBoxItem::Properties::GetDCBoxHWnd(void)
{
	ATLASSUME(pOwnerDCBox);

	OLE_HANDLE handle = NULL;
	pOwnerDCBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HWND DriveComboBoxItem::Properties::GetLBoxHWnd(void)
{
	ATLASSUME(pOwnerDCBox);

	OLE_HANDLE handle = NULL;
	pOwnerDCBox->get_hWndListBox(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

void DriveComboBoxItem::Properties::InvalidateComboBox(void)
{
	ATLASSUME(pOwnerDCBox);

	OLE_HANDLE handle = NULL;
	pOwnerDCBox->get_hWndComboBox(&handle);
	InvalidateRect(static_cast<HWND>(LongToHandle(handle)), NULL, TRUE);
}

void DriveComboBoxItem::Properties::InvalidateListBox(void)
{
	ATLASSUME(pOwnerDCBox);

	OLE_HANDLE handle = NULL;
	pOwnerDCBox->get_hWndListBox(&handle);
	InvalidateRect(static_cast<HWND>(LongToHandle(handle)), NULL, TRUE);
}


HRESULT DriveComboBoxItem::SaveState(int itemIndex, PCOMBOBOXEXITEM pTarget, HWND hWndDCBox/* = NULL*/)
{
	if(!hWndDCBox) {
		hWndDCBox = properties.GetDCBoxHWnd();
	}
	ATLASSERT(IsWindow(hWndDCBox));
	if(itemIndex < 0 || itemIndex >= static_cast<int>(SendMessage(hWndDCBox, CB_GETCOUNT, 0, 0))) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTarget, COMBOBOXEXITEM);
	if(!pTarget) {
		return E_POINTER;
	}

	ZeroMemory(pTarget, sizeof(COMBOBOXEXITEM));
	pTarget->iItem = itemIndex;
	// NOTE: cchTextMax includes the terminating null character
	pTarget->cchTextMax = static_cast<int>(SendMessage(hWndDCBox, CB_GETLBTEXTLEN, properties.itemIndex, 0)) + 1;
	pTarget->pszText = reinterpret_cast<LPTSTR>(malloc(pTarget->cchTextMax * sizeof(TCHAR)));
	pTarget->mask = CBEIF_IMAGE | CBEIF_INDENT | CBEIF_LPARAM | CBEIF_OVERLAY | CBEIF_SELECTEDIMAGE | CBEIF_TEXT;

	return static_cast<HRESULT>(SendMessage(hWndDCBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(pTarget)));
}


void DriveComboBoxItem::Attach(int itemIndex)
{
	properties.itemIndex = itemIndex;
}

void DriveComboBoxItem::Detach(void)
{
	properties.itemIndex = -2;
}

HRESULT DriveComboBoxItem::SaveState(PCOMBOBOXEXITEM pTarget, HWND hWndDCBox/* = NULL*/)
{
	return SaveState(properties.itemIndex, pTarget, hWndDCBox);
}

HRESULT DriveComboBoxItem::SaveState(VirtualDriveComboBoxItem* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerDCBox);
	COMBOBOXEXITEM item = {0};
	HRESULT hr = SaveState(&item);
	pTarget->LoadState(&item);
	SECUREFREE(item.pszText);

	return hr;
}

void DriveComboBoxItem::SetOwner(DriveComboBox* pOwner)
{
	if(properties.pOwnerDCBox) {
		properties.pOwnerDCBox->Release();
	}
	properties.pOwnerDCBox = pOwner;
	if(properties.pOwnerDCBox) {
		properties.pOwnerDCBox->AddRef();
	}
}


STDMETHODIMP DriveComboBoxItem::get_DriveType(DriveTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DriveTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	int itemIndex = properties.itemIndex;
	if(properties.itemIndex == -1 && (CWindow(hWndDCBox).GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
		itemIndex = static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0));
		if(itemIndex == -1) {
			return E_FAIL;
		}
	}

	int driveIndex = properties.pOwnerDCBox->GetDriveIndex(itemIndex);
	ATLASSERT(driveIndex >= -1);
	if(driveIndex == -1) {
		return E_FAIL;
	} else {
		TCHAR pDrive[4] = TEXT("A:\\");
		pDrive[0] += static_cast<TCHAR>(driveIndex);
		switch(GetDriveType(pDrive)) {
			case DRIVE_UNKNOWN:
				*pValue = dtUnknown;
				break;
			case DRIVE_NO_ROOT_DIR:
				*pValue = dtOther;
				break;
			case DRIVE_REMOVABLE:
				*pValue = dtRemovable;
				break;
			case DRIVE_FIXED:
				*pValue = dtFixed;
				break;
			case DRIVE_REMOTE:
				*pValue = dtRemote;
				break;
			case DRIVE_CDROM:
				*pValue = dtOptical;
				break;
			case DRIVE_RAMDISK:
				*pValue = dtRAMDisk;
				break;
		}
	}
	return S_OK;
}

STDMETHODIMP DriveComboBoxItem::get_Height(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	*pValue = static_cast<OLE_YSIZE_PIXELS>(SendMessage(hWndDCBox, CB_GETITEMHEIGHT, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP DriveComboBoxItem::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	int itemIndex = properties.itemIndex;
	if(itemIndex == -1) {
		itemIndex = static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0));
	}

	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_IMAGE;
	if(SendMessage(hWndDCBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::put_IconIndex(LONG newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_IMAGE;
	item.iImage = newValue;
	if(SendMessage(hWndDCBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(properties.itemIndex == static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0))) {
			properties.InvalidateComboBox();
		}
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	*pValue = properties.pOwnerDCBox->ItemIndexToID(properties.itemIndex);
	if(*pValue != -1) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::get_Indent(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	int itemIndex = properties.itemIndex;
	if(itemIndex == -1) {
		itemIndex = static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0));
	}

	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_INDENT;
	if(SendMessage(hWndDCBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iIndent;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::put_Indent(OLE_XSIZE_PIXELS newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_INDENT;
	item.iIndent = newValue;
	if(SendMessage(hWndDCBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex;
	return S_OK;
}

STDMETHODIMP DriveComboBoxItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < 0) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	*pValue = static_cast<LONG>(SendMessage(hWndDCBox, CB_GETITEMDATA, properties.itemIndex, 0));
	return S_OK;
}

STDMETHODIMP DriveComboBoxItem::put_ItemData(LONG newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	if(SendMessage(hWndDCBox, CB_SETITEMDATA, properties.itemIndex, newValue) != CB_ERR) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::get_OverlayIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	int itemIndex = properties.itemIndex;
	if(itemIndex == -1) {
		itemIndex = static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0));
	}

	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_OVERLAY;
	if(SendMessage(hWndDCBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iOverlay;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::put_OverlayIndex(LONG newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_OVERLAY;
	item.iOverlay = newValue;
	if(SendMessage(hWndDCBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(properties.itemIndex == static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0))) {
			properties.InvalidateComboBox();
		}
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::get_Path(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	int itemIndex = properties.itemIndex;
	if(properties.itemIndex == -1 && (CWindow(hWndDCBox).GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
		itemIndex = static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0));
		if(itemIndex == -1) {
			*pValue = SysAllocString(L"");
			return S_OK;
		}
	}

	int driveIndex = properties.pOwnerDCBox->GetDriveIndex(itemIndex);
	ATLASSERT(driveIndex >= -1);
	if(driveIndex == -1) {
		*pValue = SysAllocString(L"");
	} else {
		TCHAR pDrive[4] = TEXT("A:\\");
		pDrive[0] += static_cast<TCHAR>(driveIndex);
		*pValue = _bstr_t(pDrive).Detach();
	}
	return S_OK;
}

STDMETHODIMP DriveComboBoxItem::get_Selected(VARIANT_BOOL* pValue)
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

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	*pValue = BOOL2VARIANTBOOL(properties.itemIndex == static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0)));
	return S_OK;
}

STDMETHODIMP DriveComboBoxItem::get_SelectedIconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	int itemIndex = properties.itemIndex;
	if(itemIndex == -1) {
		itemIndex = static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0));
	}

	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_SELECTEDIMAGE;
	if(SendMessage(hWndDCBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iSelectedImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::put_SelectedIconIndex(LONG newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_SELECTEDIMAGE;
	item.iSelectedImage = newValue;
	if(SendMessage(hWndDCBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(properties.itemIndex == static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0))) {
			properties.InvalidateComboBox();
		}
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBoxItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	int itemIndex = properties.itemIndex;

	if(properties.itemIndex == -1 && (CWindow(hWndDCBox).GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
		itemIndex = static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0));
		if(itemIndex == -1) {
			*pValue = SysAllocString(L"");
			return S_OK;
		}
	}

	int textLength = static_cast<int>(SendMessage(hWndDCBox, CB_GETLBTEXTLEN, itemIndex, 0));
	if(textLength == CB_ERR) {
		return DISP_E_BADINDEX;
	}
	LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
	if(!pBuffer) {
		return E_OUTOFMEMORY;
	}
	ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
	if(SendMessage(hWndDCBox, CB_GETLBTEXT, itemIndex, reinterpret_cast<LPARAM>(pBuffer)) != CB_ERR) {
		*pValue = _bstr_t(pBuffer).Detach();
	}
	HeapFree(GetProcessHeap(), 0, pBuffer);
	return S_OK;
}

STDMETHODIMP DriveComboBoxItem::put_Text(BSTR newValue)
{
	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	COMBOBOXEXITEM item = {0};
	item.iItem = properties.itemIndex;
	item.mask = CBEIF_TEXT;
	COLE2T converter(newValue);
	if(newValue) {
		item.pszText = converter;
	} else {
		item.pszText = LPSTR_TEXTCALLBACK;
	}
	if(SendMessage(hWndDCBox, CBEM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		if(properties.itemIndex == static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0))) {
			properties.InvalidateComboBox();
		}
		properties.InvalidateListBox();
		return S_OK;
	}
	return E_FAIL;
}


STDMETHODIMP DriveComboBoxItem::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	if(properties.itemIndex < -1) {
		return E_FAIL;
	}

	ATLASSUME(properties.pOwnerDCBox);

	POINT upperLeftPoint = {0};
	*phImageList = HandleToLong(properties.pOwnerDCBox->CreateLegacyDragImage(properties.itemIndex, &upperLeftPoint, NULL));

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

STDMETHODIMP DriveComboBoxItem::GetRectangle(ItemRectangleTypeConstants /*rectangleType*/, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
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