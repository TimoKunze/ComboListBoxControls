// VirtualImageComboBoxItem.cpp: A wrapper for non-existing image combo box items.

#include "stdafx.h"
#include "VirtualImageComboBoxItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualImageComboBoxItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualImageComboBoxItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualImageComboBoxItem::Properties::~Properties()
{
	if(settings.pszText != LPSTR_TEXTCALLBACK) {
		SECUREFREE(settings.pszText);
	}
	if(pOwnerICBox) {
		pOwnerICBox->Release();
	}
}

HWND VirtualImageComboBoxItem::Properties::GetICBoxHWnd(void)
{
	ATLASSUME(pOwnerICBox);

	OLE_HANDLE handle = NULL;
	pOwnerICBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT VirtualImageComboBoxItem::LoadState(PCOMBOBOXEXITEM pSource)
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

	return S_OK;
}

HRESULT VirtualImageComboBoxItem::SaveState(PCOMBOBOXEXITEM pTarget)
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

	return S_OK;
}

void VirtualImageComboBoxItem::SetOwner(ImageComboBox* pOwner)
{
	if(properties.pOwnerICBox) {
		properties.pOwnerICBox->Release();
	}
	properties.pOwnerICBox = pOwner;
	if(properties.pOwnerICBox) {
		properties.pOwnerICBox->AddRef();
	}
}


STDMETHODIMP VirtualImageComboBoxItem::get_IconIndex(LONG* pValue)
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

STDMETHODIMP VirtualImageComboBoxItem::get_Indent(OLE_XSIZE_PIXELS* pValue)
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

STDMETHODIMP VirtualImageComboBoxItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.settings.iItem;
	return S_OK;
}

STDMETHODIMP VirtualImageComboBoxItem::get_ItemData(LONG* pValue)
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

STDMETHODIMP VirtualImageComboBoxItem::get_OverlayIndex(LONG* pValue)
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

STDMETHODIMP VirtualImageComboBoxItem::get_SelectedIconIndex(LONG* pValue)
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

STDMETHODIMP VirtualImageComboBoxItem::get_Text(BSTR* pValue)
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