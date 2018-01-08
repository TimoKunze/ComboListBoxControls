// VirtualDriveComboBoxItem.cpp: A wrapper for non-existing drive combo box items.

#include "stdafx.h"
#include "VirtualDriveComboBoxItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualDriveComboBoxItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualDriveComboBoxItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualDriveComboBoxItem::Properties::~Properties()
{
	if(settings.pszText != LPSTR_TEXTCALLBACK) {
		SECUREFREE(settings.pszText);
	}
	if(pOwnerDCBox) {
		pOwnerDCBox->Release();
	}
}

HWND VirtualDriveComboBoxItem::Properties::GetDCBoxHWnd(void)
{
	ATLASSUME(pOwnerDCBox);

	OLE_HANDLE handle = NULL;
	pOwnerDCBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT VirtualDriveComboBoxItem::LoadState(PCOMBOBOXEXITEM pSource)
{
	ATLASSERT_POINTER(pSource, COMBOBOXEXITEM);

	SECUREFREE(properties.settings.pszText);
	properties.settings = *pSource;
	if(properties.settings.mask & CBEIF_TEXT) {
		// duplicate the item's text
		if(properties.settings.pszText != LPSTR_TEXTCALLBACK) {
			properties.settings.cchTextMax = lstrlen(pSource->pszText);
			properties.settings.pszText = reinterpret_cast<LPTSTR>(malloc((properties.settings.cchTextMax + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(properties.settings.pszText, properties.settings.cchTextMax + 1, pSource->pszText)));
		}
	}
	if(properties.settings.mask & CBEIF_NOINTERCEPTION) {
		if(properties.settings.mask & CBEIF_LPARAM) {
			properties.driveIndex = properties.settings.lParam - 1;
			properties.settings.lParam = 0;
		}
	}

	return S_OK;
}

HRESULT VirtualDriveComboBoxItem::SaveState(PCOMBOBOXEXITEM pTarget)
{
	ATLASSERT_POINTER(pTarget, COMBOBOXEXITEM);

	SECUREFREE(pTarget->pszText);
	*pTarget = properties.settings;
	if(pTarget->mask & CBEIF_TEXT) {
		// duplicate the item's text
		if(pTarget->pszText != LPSTR_TEXTCALLBACK) {
			pTarget->pszText = reinterpret_cast<LPTSTR>(malloc((pTarget->cchTextMax + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(pTarget->pszText, pTarget->cchTextMax + 1, properties.settings.pszText)));
		}
	}
	if(pTarget->mask & CBEIF_NOINTERCEPTION) {
		if(pTarget->mask & CBEIF_LPARAM) {
			pTarget->lParam = properties.driveIndex + 1;
		}
	}

	return S_OK;
}

void VirtualDriveComboBoxItem::SetOwner(DriveComboBox* pOwner)
{
	if(properties.pOwnerDCBox) {
		properties.pOwnerDCBox->Release();
	}
	properties.pOwnerDCBox = pOwner;
	if(properties.pOwnerDCBox) {
		properties.pOwnerDCBox->AddRef();
	}
}


STDMETHODIMP VirtualDriveComboBoxItem::get_DriveType(DriveTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DriveTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	int driveIndex = properties.driveIndex;
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

STDMETHODIMP VirtualDriveComboBoxItem::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & CBEIF_IMAGE) {
		*pValue = properties.settings.iImage;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualDriveComboBoxItem::get_Indent(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & CBEIF_INDENT) {
		*pValue = properties.settings.iIndent;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualDriveComboBoxItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.settings.iItem;
	return S_OK;
}

STDMETHODIMP VirtualDriveComboBoxItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & CBEIF_LPARAM) {
		*pValue = static_cast<LONG>(properties.settings.lParam);
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualDriveComboBoxItem::get_OverlayIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & CBEIF_OVERLAY) {
		*pValue = properties.settings.iOverlay;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualDriveComboBoxItem::get_Path(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	int driveIndex = properties.driveIndex;
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

STDMETHODIMP VirtualDriveComboBoxItem::get_SelectedIconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & CBEIF_SELECTEDIMAGE) {
		*pValue = properties.settings.iSelectedImage;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualDriveComboBoxItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & CBEIF_TEXT) {
		if(properties.settings.pszText == LPSTR_TEXTCALLBACK) {
			*pValue = NULL;
		} else {
			*pValue = _bstr_t(properties.settings.pszText).Detach();
		}
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}