// DriveComboBox.cpp: Superclasses ComboBoxEx32.

#include "stdafx.h"
#include "DriveComboBox.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
IMEModeConstants DriveComboBox::IMEFlags::chineseIMETable[10] = {
    imeOff,
    imeOff,
    imeOff,
    imeOff,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOff
};

IMEModeConstants DriveComboBox::IMEFlags::japaneseIMETable[10] = {
    imeDisable,
    imeDisable,
    imeOff,
    imeOff,
    imeHiragana,
    imeHiragana,
    imeKatakana,
    imeKatakanaHalf,
    imeAlphaFull,
    imeAlpha
};

IMEModeConstants DriveComboBox::IMEFlags::koreanIMETable[10] = {
    imeDisable,
    imeDisable,
    imeAlpha,
    imeAlpha,
    imeHangulFull,
    imeHangul,
    imeHangulFull,
    imeHangul,
    imeAlphaFull,
    imeAlpha
};

FONTDESC DriveComboBox::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


#pragma warning(push)
#pragma warning(disable: 4355)     // passing "this" to a member constructor
DriveComboBox::DriveComboBox() :
    containedComboBox(WC_COMBOBOX, this, 1),
    containedListBox(WC_LISTBOX, this, 2)
{
	properties.font.InitializePropertyWatcher(this, DISPID_DCB_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_DCB_MOUSEICON);

	SIZEL size = {100, 20};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
	lastItemID = 0;
	currentSelectedItem = -1;
	selectedItemOnDropDown = -1;
	itemUnderMouse = -1;

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	if(RunTimeHelper::IsVista()) {
		CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
		ATLASSUME(pWICImagingFactory);
	}
}
#pragma warning(pop)


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP DriveComboBox::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IDriveComboBox, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP DriveComboBox::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<DriveComboBox>::Load(pPropertyBag, pErrorLog);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);

		CComBSTR bstr;
		hr = pPropertyBag->Read(OLESTR("CueBanner"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_CueBanner(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP DriveComboBox::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<DriveComboBox>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.cueBanner.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("CueBanner"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP DriveComboBox::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 21 VT_I4 properties...
	pSize->QuadPart += 21 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 11 VT_BOOL properties...
	pSize->QuadPart += 11 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 1 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.cueBanner.ByteLength() + sizeof(OLECHAR);

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP DriveComboBox::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x10101010/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version > 0x0102) {
		return E_FAIL;
	}
	LONG subSignature = 0;
	if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}
	if(subSignature != 0x02020202/*4x 0x02 (-> DriveComboBox)*/) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL caseSensitiveItemSearching = propertyValue.boolVal;
	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR cueBanner;
	if(FAILED(hr = cueBanner.ReadFromStream(pStream))) {
		return hr;
	}
	if(!cueBanner) {
		cueBanner = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DriveTypeConstants displayedDriveTypes = static_cast<DriveTypeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisplayNameStyleConstants displayNameStyle = static_cast<DisplayNameStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayOverlayImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL doOEMConversion = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragDropDownTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DropDownKeyConstants dropDownKey = static_cast<DropDownKeyConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR foreColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HandleOLEDragDropConstants handleOLEDragDrop = static_cast<HandleOLEDragDropConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IconVisibilityConstants iconVisibility = static_cast<IconVisibilityConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants IMEMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS itemHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL listAlwaysShowVerticalScrollBar = propertyValue.boolVal;
	/*if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR listBackColor = propertyValue.lVal;*/
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG listDragScrollTimeBase = propertyValue.lVal;
	/*if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR listForeColor = propertyValue.lVal;*/
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS listHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS listWidth = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLEDragImageStyleConstants oleDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS selectionFieldHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(version < 0x0102) {
		if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
			return hr;
		}
		CComBSTR text;
		if(FAILED(hr = text.ReadFromStream(pStream))) {
			return hr;
		}
		if(!text) {
			text = L"";
		}
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL textEndEllipsis = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	UseSystemImageListConstants useSystemImageList = static_cast<UseSystemImageListConstants>(propertyValue.lVal);

	DriveTypeConstants driveTypesWithVolumeName = static_cast<DriveTypeConstants>(dtFixed | dtRAMDisk);
	if(version >= 0x0101) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		driveTypesWithVolumeName = static_cast<DriveTypeConstants>(propertyValue.lVal);
	}


	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CaseSensitiveItemSearching(caseSensitiveItemSearching);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CueBanner(cueBanner);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayedDriveTypes(displayedDriveTypes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayNameStyle(displayNameStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayOverlayImages(displayOverlayImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragDropDownTime(dragDropDownTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DriveTypesWithVolumeName(driveTypesWithVolumeName);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DropDownKey(dropDownKey);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HandleOLEDragDrop(handleOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IconVisibility(iconVisibility);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IMEMode(IMEMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemHeight(itemHeight);
	if(FAILED(hr)) {
		return hr;
	}
	/*hr = put_ListBackColor(listBackColor);
	if(FAILED(hr)) {
		return hr;
	}*/
	hr = put_ListDragScrollTimeBase(listDragScrollTimeBase);
	if(FAILED(hr)) {
		return hr;
	}
	/*hr = put_ListForeColor(listForeColor);
	if(FAILED(hr)) {
		return hr;
	}*/
	hr = put_ListHeight(listHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ListWidth(listWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OLEDragImageStyle(oleDragImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SelectionFieldHeight(selectionFieldHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TextEndEllipsis(textEndEllipsis);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemImageList(useSystemImageList);
	if(FAILED(hr)) {
		return hr;
	}
	BOOL b = flags.dontRecreate;
	flags.dontRecreate = TRUE;
	hr = put_DoOEMConversion(doOEMConversion);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ListAlwaysShowVerticalScrollBar(listAlwaysShowVerticalScrollBar);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	flags.dontRecreate = b;
	RecreateControlWindow();

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP DriveComboBox::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x10101010/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0102;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x02020202/*4x 0x02 (-> DriveComboBox)*/;
	if(FAILED(hr = pStream->Write(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.caseSensitiveItemSearching);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	VARTYPE vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.cueBanner.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.displayedDriveTypes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.displayNameStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayOverlayImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.doOEMConversion);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.dragDropDownTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.dropDownKey;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.handleOLEDragDrop;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.iconVisibility;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.IMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.setItemHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.listAlwaysShowVerticalScrollBar);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	/*propertyValue.lVal = properties.listBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}*/
	propertyValue.lVal = properties.listDragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	/*propertyValue.lVal = properties.listForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}*/
	propertyValue.lVal = properties.listHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.listWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.oleDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.setSelectionFieldHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.textEndEllipsis);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.useSystemImageList;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0101 starts here
	propertyValue.lVal = properties.driveTypesWithVolumeName;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND DriveComboBox::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_USEREX_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<DriveComboBox>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT DriveComboBox::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("DriveComboBox ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void DriveComboBox::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP DriveComboBox::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<DriveComboBox>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	return hr;
}

STDMETHODIMP DriveComboBox::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL DriveComboBox::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		if(pMessage->wParam == VK_TAB) {
			if(dialogCode & DLGC_WANTTAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
		}
	}
	return CComControl<DriveComboBox>::PreTranslateAccelerator(pMessage, hReturnValue);
}

HIMAGELIST DriveComboBox::CreateLegacyDragImage(int itemIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	/********************************************************************************************************
	 * Known problems:                                                                                      *
	 * - We use hardcoded margins.                                                                          *
	 ********************************************************************************************************/

	HIMAGELIST hSourceImageList = NULL;
	if(!(SendMessage(CBEM_GETEXSTYLE, 0, 0) & CBES_EX_NOEDITIMAGEINDENT)) {
		hSourceImageList = reinterpret_cast<HIMAGELIST>(SendMessage(CBEM_GETIMAGELIST, 0, 0));
	}
	SIZE imageSize = {0};
	if(hSourceImageList) {
		ImageList_GetIconSize(hSourceImageList, reinterpret_cast<PINT>(&imageSize.cx), reinterpret_cast<PINT>(&imageSize.cy));
	}

	// retrieve item details
	int selectedItemIndex = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
	BOOL itemIsSelected = (itemIndex == -1) || (itemIndex == selectedItemIndex);
	COMBOBOXEXITEM item = {0};
	item.iItem = (itemIsSelected ? selectedItemIndex : itemIndex);
	item.mask = CBEIF_IMAGE | CBEIF_SELECTEDIMAGE | CBEIF_OVERLAY | CBEIF_INDENT;
	SendMessage(CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item));
	if(itemIndex == -1 && (GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
		if(selectedItemIndex == CB_ERR) {
			return NULL;
		}
		itemIndex = selectedItemIndex;
	}
	int itemTextLength = static_cast<int>(SendMessage(CB_GETLBTEXTLEN, itemIndex, 0));
	if(itemTextLength == CB_ERR) {
		return NULL;
	}
	LPTSTR pItemText = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (itemTextLength + 1) * sizeof(TCHAR)));
	if(!pItemText) {
		return NULL;
	}
	SendMessage(CB_GETLBTEXT, itemIndex, reinterpret_cast<LPARAM>(pItemText));

	// retrieve window details
	BOOL hasFocus = (IsChild(GetFocus()));
	DWORD style = GetExStyle();
	DWORD textDrawStyle = DT_EDITCONTROL | DT_NOPREFIX | DT_SINGLELINE;
	if(style & WS_EX_RTLREADING) {
		textDrawStyle |= DT_RTLREADING;
	}
	BOOL layoutRTL = ((style & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
	BOOL itemIsFocused = hasFocus && itemIsSelected;
	int iconToDraw = (itemIsFocused && (selectedItemIndex != -1) ? item.iSelectedImage : item.iImage);

	// create the DCs we'll draw into
	HDC hCompatibleDC = GetDC();
	CDC memoryDC;
	memoryDC.CreateCompatibleDC(hCompatibleDC);
	CDC maskMemoryDC;
	maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

	CFontHandle font = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
	HFONT hPreviousFont = NULL;
	if(!font.IsNull()) {
		hPreviousFont = memoryDC.SelectFont(font);
	}

	// prepare themed item drawing
	CTheme themingEngine;
	BOOL themedListItems = FALSE;
	int themeState = 0/*LBPSI_NORMAL*/;
	if(itemIsSelected) {
		if(hasFocus) {
			themeState = LBPSI_SELECTED;
		} else {
			themeState = LBPSI_SELECTEDNOTFOCUS;
		}
	}
	if(flags.usingThemes) {
		themingEngine.OpenThemeData(*this, VSCLASS_LISTBOX);
		/* We use LBPSI_SELECTED here, because it's more likely this one is defined and we don't want a mixture
		   of themed and non-themed items in drag images. What we're doing with LBPSI_NORMAL should work
		   regardless whether LBPSI_NORMAL is defined. */
		themedListItems = themingEngine.IsThemePartDefined(LBCP_ITEM, LBPSI_SELECTED/*themeState*/);
	}

	// calculate the bounding rectangles of the various item parts
	CRect itemBoundingRect;
	CRect selectionBoundingRect;
	CRect labelBoundingRect;
	CRect iconBoundingRect;

	if(itemIsSelected) {
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(containedComboBox, &controlInfo);
		containedComboBox.MapWindowPoints(*this, &controlInfo.rcItem);
		selectionBoundingRect = controlInfo.rcItem;
	} else if(containedListBox.IsWindow()) {
		containedListBox.SendMessage(LB_GETITEMRECT, itemIndex, reinterpret_cast<LPARAM>(&selectionBoundingRect));
		containedListBox.MapWindowPoints(*this, &selectionBoundingRect);
	} else {
		ATLASSERT(FALSE && "TODO: Retrieve item rectangle for non-selected item");
	}

	if(!itemIsSelected) {
		selectionBoundingRect.OffsetRect(item.iIndent * 10, 0);
	}
	if(hSourceImageList) {
		iconBoundingRect = selectionBoundingRect;
		iconBoundingRect.right = iconBoundingRect.left + imageSize.cx;
	}
	if(hSourceImageList) {
		selectionBoundingRect.left = iconBoundingRect.right + 4;
	}
	if(themedListItems) {
		CT2W converter(pItemText);
		LPWSTR pLabelText = converter;
		themingEngine.DrawThemeText(memoryDC, LBCP_ITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle | DT_TOP | DT_CALCRECT, 0, &selectionBoundingRect);
	} else {
		memoryDC.DrawText(pItemText, itemTextLength, &selectionBoundingRect, textDrawStyle | DT_TOP | DT_CALCRECT);
	}
	if(hSourceImageList) {
		selectionBoundingRect.OffsetRect(0, iconBoundingRect.bottom - selectionBoundingRect.bottom - 1);
	}
	if(!layoutRTL) {
		selectionBoundingRect.right++;
	}
	labelBoundingRect = selectionBoundingRect;
	if(!themedListItems && hasFocus && itemIsFocused && (SendMessage(WM_QUERYUISTATE, 0, 0) & UISF_HIDEFOCUS) == 0) {
		if(layoutRTL) {
			selectionBoundingRect.right++;
		} else {
			selectionBoundingRect.left--;
		}
	}
	itemBoundingRect = selectionBoundingRect;
	// center the label rectangle vertically
	int cy = labelBoundingRect.Height();
	labelBoundingRect.top = itemBoundingRect.top + (itemBoundingRect.Height() - cy) / 2;
	labelBoundingRect.bottom = labelBoundingRect.top + cy;
	itemBoundingRect.UnionRect(&itemBoundingRect, &labelBoundingRect);
	if(hSourceImageList) {
		itemBoundingRect.UnionRect(&itemBoundingRect, &iconBoundingRect);
	}
	if(pBoundingRectangle) {
		*pBoundingRectangle = itemBoundingRect;
	}

	// calculate drag image size and upper-left corner
	SIZE dragImageSize = {0};
	if(pUpperLeftPoint) {
		pUpperLeftPoint->x = itemBoundingRect.left;
		pUpperLeftPoint->y = itemBoundingRect.top;
	}
	dragImageSize.cx = itemBoundingRect.Width();
	dragImageSize.cy = itemBoundingRect.Height();

	// offset RECTs
	SIZE offset = {0};
	offset.cx = itemBoundingRect.left;
	offset.cy = itemBoundingRect.top;
	labelBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	iconBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	selectionBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	itemBoundingRect.OffsetRect(-offset.cx, -offset.cy);

	// setup the DCs we'll draw into
	if(itemIsSelected) {
		memoryDC.SetBkColor(GetSysColor((hasFocus ? COLOR_HIGHLIGHT : COLOR_BTNFACE)));
		memoryDC.SetTextColor(GetSysColor((hasFocus ? COLOR_HIGHLIGHTTEXT : COLOR_BTNTEXT)));
	} else {
		memoryDC.SetBkColor(GetSysColor(COLOR_WINDOW));
		memoryDC.SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
		memoryDC.SetBkMode(TRANSPARENT);
	}

	// create drag image bitmap
	/* NOTE: We prefer creating 32bpp drag images, because this improves performance of
	         ListBoxItemContainer::CreateDragImage(). */
	BOOL doAlphaChannelProcessing = RunTimeHelper::IsCommCtrl6();
	BITMAPINFO bitmapInfo = {0};
	if(doAlphaChannelProcessing) {
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = dragImageSize.cx;
		bitmapInfo.bmiHeader.biHeight = -dragImageSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
	}
	CBitmap dragImage;
	LPRGBQUAD pDragImageBits = NULL;
	if(doAlphaChannelProcessing) {
		dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
	} else {
		dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageSize.cx, dragImageSize.cy);
	}
	HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
	CBitmap dragImageMask;
	dragImageMask.CreateBitmap(dragImageSize.cx, dragImageSize.cy, 1, 1, NULL);
	HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);

	// initialize the bitmap
	if(themedListItems) {
		// we need a transparent background
		LPRGBQUAD pPixel = pDragImageBits;
		for(int y = 0; y < dragImageSize.cy; ++y) {
			for(int x = 0; x < dragImageSize.cx; ++x, ++pPixel) {
				pPixel->rgbRed = 0xFF;
				pPixel->rgbGreen = 0xFF;
				pPixel->rgbBlue = 0xFF;
				pPixel->rgbReserved = 0x00;
			}
		}
	} else {
		memoryDC.FillRect(&itemBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
	}
	maskMemoryDC.FillRect(&itemBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

	// draw the selection area's background
	if(itemIsSelected) {
		if(themedListItems) {
			themingEngine.DrawThemeBackground(memoryDC, LBCP_ITEM, themeState, &selectionBoundingRect, NULL);
		} else {
			memoryDC.FillRect(&selectionBoundingRect, (hasFocus ? COLOR_HIGHLIGHT : COLOR_BTNFACE));
		}
		maskMemoryDC.FillRect(&selectionBoundingRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
	}

	if(!(SendMessage(CBEM_GETEXSTYLE, 0, 0) & CBES_EX_NOEDITIMAGE)) {
		// draw the icon
		if(hSourceImageList) {
			//ImageList_DrawEx(hSourceImageList, iconToDraw, memoryDC, iconBoundingRect.left, iconBoundingRect.top, imageSize.cx, imageSize.cy, CLR_NONE, CLR_NONE, (itemIsSelected ? ILD_SELECTED : ILD_NORMAL) | INDEXTOOVERLAYMASK(item.iOverlay));
			ImageList_DrawEx(hSourceImageList, iconToDraw, memoryDC, iconBoundingRect.left, iconBoundingRect.top, imageSize.cx, imageSize.cy, CLR_NONE, CLR_NONE, ILD_NORMAL | INDEXTOOVERLAYMASK(item.iOverlay));
			ImageList_Draw(hSourceImageList, iconToDraw, maskMemoryDC, iconBoundingRect.left, iconBoundingRect.top, ILD_MASK | INDEXTOOVERLAYMASK(item.iOverlay));
		}
	}

	if(itemTextLength > 0) {
		// draw the text
		CRect rc = labelBoundingRect;
		if(themedListItems) {
			CT2W converter(pItemText);
			LPWSTR pLabelText = converter;
			themingEngine.DrawThemeText(memoryDC, LBCP_ITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle | DT_VCENTER, 0, &rc);
		} else {
			memoryDC.DrawText(pItemText, itemTextLength, &rc, textDrawStyle | DT_VCENTER);
		}
		if(!itemIsSelected) {
			COLORREF bkColor = memoryDC.GetBkColor();
			for(int y = rc.top; y <= rc.bottom; ++y) {
				for(int x = rc.left; x <= rc.right; ++x) {
					if(memoryDC.GetPixel(x, y) != bkColor) {
						maskMemoryDC.SetPixelV(x, y, 0x00000000);
					}
				}
			}
		}
	}

	if(!themedListItems) {
		// draw the focus rectangle
		if(hasFocus && itemIsFocused) {
			if((SendMessage(WM_QUERYUISTATE, 0, 0) & UISF_HIDEFOCUS) == 0) {
				RECT focusRect = selectionBoundingRect;
				memoryDC.DrawFocusRect(&focusRect);
				maskMemoryDC.FrameRect(&focusRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
			}
		}
	}

	if(doAlphaChannelProcessing) {
		// correct the alpha channel
		LPRGBQUAD pPixel = pDragImageBits;
		POINT pt;
		for(pt.y = 0; pt.y < dragImageSize.cy; ++pt.y) {
			for(pt.x = 0; pt.x < dragImageSize.cx; ++pt.x, ++pPixel) {
				if(layoutRTL) {
					// we're working on raw data, so we've to handle WS_EX_LAYOUTRTL ourselves
					POINT pt2 = pt;
					pt2.x = dragImageSize.cx - pt.x - 1;
					if(maskMemoryDC.GetPixel(pt2.x, pt2.y) == 0x00000000) {
						if(themedListItems) {
							if(itemIsSelected) {
								if((pPixel->rgbReserved != 0x00) || labelBoundingRect.PtInRect(pt2)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else if(labelBoundingRect.PtInRect(pt2)) {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						} else {
							// items are not themed
							if(itemIsSelected) {
								if((pPixel->rgbReserved == 0x00) || selectionBoundingRect.PtInRect(pt2)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						}
					}
				} else {
					// layout is left to right
					if(maskMemoryDC.GetPixel(pt.x, pt.y) == 0x00000000) {
						if(themedListItems) {
							if(itemIsSelected) {
								if((pPixel->rgbReserved != 0x00) || labelBoundingRect.PtInRect(pt)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else if(labelBoundingRect.PtInRect(pt)) {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						} else {
							// items are not themed
							if(itemIsSelected) {
								if((pPixel->rgbReserved == 0x00) || selectionBoundingRect.PtInRect(pt)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						}
					}
				}
			}
		}
	}

	memoryDC.SelectBitmap(hPreviousBitmap);
	maskMemoryDC.SelectBitmap(hPreviousBitmapMask);
	if(hPreviousFont) {
		memoryDC.SelectFont(hPreviousFont);
	}

	// create the imagelist
	HIMAGELIST hDragImageList = ImageList_Create(dragImageSize.cx, dragImageSize.cy, (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
	ImageList_SetBkColor(hDragImageList, CLR_NONE);
	ImageList_Add(hDragImageList, dragImage, dragImageMask);

	HeapFree(GetProcessHeap(), 0, pItemText);
	ReleaseDC(hCompatibleDC);

	return hDragImageList;
}

BOOL DriveComboBox::CreateLegacyOLEDragImage(IDriveComboBoxItemContainer* pItems, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pItems);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	// use a normal legacy drag image as base
	OLE_HANDLE h = NULL;
	OLE_XPOS_PIXELS xUpperLeft = 0;
	OLE_YPOS_PIXELS yUpperLeft = 0;
	pItems->CreateDragImage(&xUpperLeft, &yUpperLeft, &h);
	if(h) {
		HIMAGELIST hImageList = static_cast<HIMAGELIST>(LongToHandle(h));

		// retrieve the drag image's size
		int bitmapHeight;
		int bitmapWidth;
		ImageList_GetIconSize(hImageList, &bitmapWidth, &bitmapHeight);
		pDragImage->sizeDragImage.cx = bitmapWidth;
		pDragImage->sizeDragImage.cy = bitmapHeight;

		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		pDragImage->hbmpDragImage = NULL;

		if(RunTimeHelper::IsCommCtrl6()) {
			// handle alpha channel
			IImageList* pImgLst = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
				HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
				if(pfnHIMAGELIST_QueryInterface) {
					pfnHIMAGELIST_QueryInterface(hImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
				}
				FreeLibrary(hMod);
			}
			if(!pImgLst) {
				pImgLst = reinterpret_cast<IImageList*>(hImageList);
				pImgLst->AddRef();
			}
			ATLASSUME(pImgLst);

			DWORD imageFlags = 0;
			pImgLst->GetItemFlags(0, &imageFlags);
			if(imageFlags & ILIF_ALPHA) {
				// the drag image makes use of the alpha channel
				IMAGEINFO imageInfo = {0};
				ImageList_GetImageInfo(hImageList, 0, &imageInfo);

				// fetch raw data
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
				bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pSourceBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * sizeof(RGBQUAD)));
				GetDIBits(memoryDC, imageInfo.hbmImage, 0, pDragImage->sizeDragImage.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
				// create target bitmap
				LPRGBQUAD pDragImageBits = NULL;
				pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
				pDragImage->crColorKey = 0xFFFFFFFF;

				// transfer raw data
				CopyMemory(pDragImageBits, pSourceBits, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * 4);

				// clean up
				HeapFree(GetProcessHeap(), 0, pSourceBits);
				DeleteObject(imageInfo.hbmImage);
				DeleteObject(imageInfo.hbmMask);
			}
			pImgLst->Release();
		}

		if(!pDragImage->hbmpDragImage) {
			// fallback mode
			memoryDC.SetBkMode(TRANSPARENT);

			// create target bitmap
			HDC hCompatibleDC = ::GetDC(NULL);
			pDragImage->hbmpDragImage = CreateCompatibleBitmap(hCompatibleDC, bitmapWidth, bitmapHeight);
			::ReleaseDC(NULL, hCompatibleDC);
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(pDragImage->hbmpDragImage);

			// draw target bitmap
			pDragImage->crColorKey = RGB(0xF4, 0x00, 0x00);
			CBrush backroundBrush;
			backroundBrush.CreateSolidBrush(pDragImage->crColorKey);
			memoryDC.FillRect(CRect(0, 0, bitmapWidth, bitmapHeight), backroundBrush);
			ImageList_Draw(hImageList, 0, memoryDC, 0, 0, ILD_NORMAL);

			// clean up
			memoryDC.SelectBitmap(hPreviousBitmap);
		}

		ImageList_Destroy(hImageList);

		if(pDragImage->hbmpDragImage) {
			// retrieve the offset
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			if(GetExStyle() & WS_EX_LAYOUTRTL) {
				pDragImage->ptOffset.x = xUpperLeft + pDragImage->sizeDragImage.cx - mousePosition.x;
			} else {
				pDragImage->ptOffset.x = mousePosition.x - xUpperLeft;
			}
			pDragImage->ptOffset.y = mousePosition.y - yUpperLeft;

			succeeded = TRUE;
		}
	}

	return succeeded;
}

BOOL DriveComboBox::CreateVistaOLEDragImage(IDriveComboBoxItemContainer* pItems, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pItems);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	CTheme themingEngine;
	themingEngine.OpenThemeData(NULL, VSCLASS_DRAGDROP);
	if(themingEngine.IsThemeNull()) {
		// FIXME: What should we do here?
		ATLASSERT(FALSE && "Current theme does not define the \"DragDrop\" class.");
	} else {
		// retrieve the drag image's size
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();

		themingEngine.GetThemePartSize(memoryDC, DD_IMAGEBG, 1, NULL, TS_TRUE, &pDragImage->sizeDragImage);
		MARGINS margins = {0};
		themingEngine.GetThemeMargins(memoryDC, DD_IMAGEBG, 1, TMT_CONTENTMARGINS, NULL, &margins);
		pDragImage->sizeDragImage.cx -= margins.cxLeftWidth + margins.cxRightWidth;
		pDragImage->sizeDragImage.cy -= margins.cyTopHeight + margins.cyBottomHeight;
	}

	ATLASSERT(pDragImage->sizeDragImage.cx > 0);
	ATLASSERT(pDragImage->sizeDragImage.cy > 0);

	// create target bitmap
	BITMAPINFO bitmapInfo = {0};
	bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
	bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
	bitmapInfo.bmiHeader.biPlanes = 1;
	bitmapInfo.bmiHeader.biBitCount = 32;
	bitmapInfo.bmiHeader.biCompression = BI_RGB;
	LPRGBQUAD pDragImageBits = NULL;
	pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);

	HIMAGELIST hSourceImageList = properties.hHighResImageList;
	if(!hSourceImageList) {
		hSourceImageList = reinterpret_cast<HIMAGELIST>(SendMessage(CBEM_GETIMAGELIST, 0, 0));
	}
	if(!hSourceImageList) {
		// report success, although we've an empty drag image
		return TRUE;
	}

	IImageList2* pImgLst = NULL;
	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
		HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
		if(pfnHIMAGELIST_QueryInterface) {
			pfnHIMAGELIST_QueryInterface(hSourceImageList, IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
		}
		FreeLibrary(hMod);
	}
	if(!pImgLst) {
		IImageList* p = reinterpret_cast<IImageList*>(hSourceImageList);
		p->QueryInterface(IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
	}
	ATLASSUME(pImgLst);

	if(pImgLst) {
		LONG numberOfItems = 0;
		pItems->Count(&numberOfItems);
		ATLASSERT(numberOfItems > 0);
		// don't display more than 5 (10) thumbnails
		numberOfItems = min(numberOfItems, (hSourceImageList == properties.hHighResImageList ? 5 : 10));

		CComPtr<IUnknown> pUnknownEnum = NULL;
		pItems->get__NewEnum(&pUnknownEnum);
		CComQIPtr<IEnumVARIANT> pEnum = pUnknownEnum;
		ATLASSUME(pEnum);
		if(pEnum) {
			int cx = 0;
			int cy = 0;
			pImgLst->GetIconSize(&cx, &cy);
			SIZE thumbnailSize;
			thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfItems - 1);
			if(thumbnailSize.cy < 8) {
				// don't get smaller than 8x8 thumbnails
				numberOfItems = (pDragImage->sizeDragImage.cy - 8) / 3 + 1;
				thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfItems - 1);
			}
			thumbnailSize.cx = thumbnailSize.cy;
			int thumbnailBufferSize = thumbnailSize.cx * thumbnailSize.cy * sizeof(RGBQUAD);
			LPRGBQUAD pThumbnailBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, thumbnailBufferSize));
			ATLASSERT(pThumbnailBits);
			if(pThumbnailBits) {
				// iterate over the dragged items
				VARIANT v;
				int i = 0;
				CComPtr<IDriveComboBoxItem> pItem = NULL;
				while(pEnum->Next(1, &v, NULL) == S_OK) {
					if(v.vt == VT_DISPATCH) {
						v.pdispVal->QueryInterface(IID_IDriveComboBoxItem, reinterpret_cast<LPVOID*>(&pItem));
						ATLASSUME(pItem);
						if(pItem) {
							// get the item's icon
							LONG icon = 0;
							pItem->get_IconIndex(&icon);
							LONG overlay = 0;
							pItem->get_OverlayIndex(&overlay);

							pImgLst->ForceImagePresent(icon, ILFIP_ALWAYS);
							HICON hIcon = NULL;
							pImgLst->GetIcon(icon, ILD_TRANSPARENT, &hIcon);
							ATLASSERT(hIcon);
							if(hIcon) {
								// finally create the thumbnail
								ZeroMemory(pThumbnailBits, thumbnailBufferSize);
								HRESULT hr = CreateThumbnail(hIcon, thumbnailSize, pThumbnailBits, TRUE);
								DestroyIcon(hIcon);
								if(FAILED(hr)) {
									pItem = NULL;
									VariantClear(&v);
									break;
								}

								// add the thumbail to the drag image keeping the alpha channel intact
								if(i == 0) {
									LPRGBQUAD pDragImagePixel = pDragImageBits;
									LPRGBQUAD pThumbnailPixel = pThumbnailBits;
									for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx, pThumbnailPixel += thumbnailSize.cx) {
										CopyMemory(pDragImagePixel, pThumbnailPixel, thumbnailSize.cx * sizeof(RGBQUAD));
									}
								} else {
									LPRGBQUAD pDragImagePixel = pDragImageBits;
									LPRGBQUAD pThumbnailPixel = pThumbnailBits;
									pDragImagePixel += 3 * i * pDragImage->sizeDragImage.cx;
									for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx) {
										LPRGBQUAD p = pDragImagePixel + 2 * i;
										for(int x = 0; x < thumbnailSize.cx; ++x, ++p, ++pThumbnailPixel) {
											// merge the pixels
											p->rgbRed = pThumbnailPixel->rgbRed * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbRed / 0xFF;
											p->rgbGreen = pThumbnailPixel->rgbGreen * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbGreen / 0xFF;
											p->rgbBlue = pThumbnailPixel->rgbBlue * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbBlue / 0xFF;
											p->rgbReserved = pThumbnailPixel->rgbReserved + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbReserved / 0xFF;
										}
									}
								}
							}

							++i;
							pItem = NULL;
							if(i == numberOfItems) {
								VariantClear(&v);
								break;
							}
						}
					}
					VariantClear(&v);
				}
				HeapFree(GetProcessHeap(), 0, pThumbnailBits);
				succeeded = TRUE;
			}
		}

		pImgLst->Release();
	}

	return succeeded;
}

HWND DriveComboBox::GethWndShellUIParentWindow(void)
{
	return properties.hWndShellUIParentWindow;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP DriveComboBox::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	HWND hWndToUse = NULL;
	if(WindowFromPoint(buffer) == containedListBox) {
		hWndToUse = containedListBox;
		Raise_ListOLEDragEnter(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);
	} else {
		hWndToUse = *this;
		Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);
	}

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(hWndToUse, pDataObject, &buffer, *pEffect);
		if(dragDropStatus.useItemCountLabelHack) {
			dragDropStatus.pDropTargetHelper->DragLeave();
			dragDropStatus.pDropTargetHelper->DragEnter(hWndToUse, pDataObject, &buffer, *pEffect);
			dragDropStatus.useItemCountLabelHack = FALSE;
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::DragLeave(void)
{
	BOOL callDropTargetHelper = TRUE;
	if(dragDropStatus.isOverListBox) {
		Raise_ListOLEDragLeave(FALSE, &callDropTargetHelper);
	} else {
		Raise_OLEDragLeave(FALSE, &callDropTargetHelper);
	}
	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragLeave();
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	HWND hWnd = WindowFromPoint(buffer);
	if(hWnd == containedListBox) {
		if(dragDropStatus.isOverListBox) {
			Raise_ListOLEDragMouseMove(pEffect, keyState, mousePosition, &callDropTargetHelper);
		} else {
			// we've left the combo box and entered the list box
			Raise_OLEDragLeave(TRUE, &callDropTargetHelper);
			if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragLeave();
			}

			DROPDESCRIPTION newDropDescription;
			ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
			IDataObject_GetDropDescription(pDataObject, newDropDescription);
			if(memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
				InvalidateDragWindow(pDataObject);
				oldDropDescription = newDropDescription;
			}

			Raise_ListOLEDragEnter(NULL, pEffect, keyState, mousePosition, &callDropTargetHelper);
			if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragEnter(containedListBox, pDataObject, &buffer, *pEffect);
			}
		}
	} else {
		if(dragDropStatus.isOverListBox) {
			// we've left the list box and entered the combo box
			Raise_ListOLEDragLeave(TRUE, &callDropTargetHelper);
			if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragLeave();
			}

			DROPDESCRIPTION newDropDescription;
			ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
			IDataObject_GetDropDescription(pDataObject, newDropDescription);
			if(memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
				InvalidateDragWindow(pDataObject);
				oldDropDescription = newDropDescription;
			}

			Raise_OLEDragEnter(NULL, pEffect, keyState, mousePosition, &callDropTargetHelper);
			if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			}
		} else {
			Raise_OLEDragMouseMove(pEffect, keyState, mousePosition, &callDropTargetHelper);
		}
	}
	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	BOOL callDropTargetHelper = TRUE;
	if(WindowFromPoint(buffer) == containedListBox) {
		Raise_ListOLEDragDrop(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);
	} else {
		Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);
	}
	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSource
STDMETHODIMP DriveComboBox::GiveFeedback(DWORD effect)
{
	VARIANT_BOOL useDefaultCursors = VARIANT_TRUE;
	//if(flags.usingThemes && RunTimeHelper::IsVista()) {
		ATLASSUME(dragDropStatus.pSourceDataObject);

		BOOL isShowingLayered = FALSE;
		FORMATETC format = {0};
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("IsShowingLayered")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		STGMEDIUM medium = {0};
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				isShowingLayered = *static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		BOOL useDropDescriptionHack = FALSE;
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				useDropDescriptionHack = *static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}

		if(isShowingLayered && properties.oleDragImageStyle != odistClassic) {
			SetCursor(static_cast<HCURSOR>(LoadImage(NULL, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED)));
			useDefaultCursors = VARIANT_FALSE;
		}
		if(useDropDescriptionHack) {
			// this will make drop descriptions work
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragWindow")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
				if(medium.hGlobal) {
					// WM_USER + 1 (with wParam = 0 and lParam = 0) hides the drag image
					#define WM_SETDROPEFFECT				WM_USER + 2     // (wParam = DCID_*, lParam = 0)
					#define DDWM_UPDATEWINDOW				WM_USER + 3     // (wParam = 0, lParam = 0)
					typedef enum DROPEFFECTS
					{
						DCID_NULL = 0,
						DCID_NO = 1,
						DCID_MOVE = 2,
						DCID_COPY = 3,
						DCID_LINK = 4,
						DCID_MAX = 5
					} DROPEFFECTS;

					HWND hWndDragWindow = *static_cast<HWND*>(GlobalLock(medium.hGlobal));
					GlobalUnlock(medium.hGlobal);

					DROPEFFECTS dropEffect = DCID_NULL;
					switch(effect) {
						case DROPEFFECT_NONE:
							dropEffect = DCID_NO;
							break;
						case DROPEFFECT_COPY:
							dropEffect = DCID_COPY;
							break;
						case DROPEFFECT_MOVE:
							dropEffect = DCID_MOVE;
							break;
						case DROPEFFECT_LINK:
							dropEffect = DCID_LINK;
							break;
					}
					if(::IsWindow(hWndDragWindow)) {
						::PostMessage(hWndDragWindow, WM_SETDROPEFFECT, dropEffect, 0);
					}
				}
				ReleaseStgMedium(&medium);
			}
		}
	//}

	Raise_OLEGiveFeedback(effect, &useDefaultCursors);
	return (useDefaultCursors == VARIANT_FALSE ? S_OK : DRAGDROP_S_USEDEFAULTCURSORS);
}

STDMETHODIMP DriveComboBox::QueryContinueDrag(BOOL pressedEscape, DWORD keyState)
{
	HRESULT actionToContinueWith = S_OK;
	if(pressedEscape) {
		actionToContinueWith = DRAGDROP_S_CANCEL;
	} else if(!(keyState & dragDropStatus.draggingMouseButton)) {
		actionToContinueWith = DRAGDROP_S_DROP;
	}
	Raise_OLEQueryContinueDrag(pressedEscape, keyState, &actionToContinueWith);
	return actionToContinueWith;
}
// implementation of IDropSource
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSourceNotify
STDMETHODIMP DriveComboBox::DragEnterTarget(HWND hWndTarget)
{
	Raise_OLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	return S_OK;
}

STDMETHODIMP DriveComboBox::DragLeaveTarget(void)
{
	Raise_OLEDragLeavePotentialTarget();
	return S_OK;
}
// implementation of IDropSourceNotify
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP DriveComboBox::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Accessibility:
			*pName = GetResString(IDPC_ACCESSIBILITY).Detach();
			return S_OK;
			break;
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_DCB_DROPDOWNBUTTONOBJECTSTATE:
			*pCategory = PROPCAT_Accessibility;
			return S_OK;
			break;
		/*case DISPID_DCB_APPEARANCE:
		case DISPID_DCB_BORDERSTYLE:*/
		case DISPID_DCB_CUEBANNER:
		case DISPID_DCB_DISPLAYNAMESTYLE:
		case DISPID_DCB_DISPLAYOVERLAYIMAGES:
		case DISPID_DCB_HWNDSHELLUIPARENTWINDOW:
		case DISPID_DCB_ICONVISIBILITY:
		case DISPID_DCB_ISDROPPEDDOWN:
		case DISPID_DCB_ITEMHEIGHT:
		case DISPID_DCB_MOUSEICON:
		case DISPID_DCB_MOUSEPOINTER:
		case DISPID_DCB_SELECTIONFIELDHEIGHT:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_DCB_CASESENSITIVEITEMSEARCHING:
		case DISPID_DCB_DISABLEDEVENTS:
		case DISPID_DCB_DISPLAYEDDRIVETYPES:
		case DISPID_DCB_DONTREDRAW:
		case DISPID_DCB_DOOEMCONVERSION:
		case DISPID_DCB_DRIVETYPESWITHVOLUMENAME:
		case DISPID_DCB_DROPDOWNKEY:
		case DISPID_DCB_HOVERTIME:
		case DISPID_DCB_IMEMODE:
		case DISPID_DCB_LISTALWAYSSHOWVERTICALSCROLLBAR:
		case DISPID_DCB_LISTHEIGHT:
		//case DISPID_DCB_LISTSCROLLABLEWIDTH:
		case DISPID_DCB_LISTWIDTH:
		//case DISPID_DCB_MINVISIBLEITEMS:
		case DISPID_DCB_PROCESSCONTEXTMENUKEYS:
		case DISPID_DCB_RIGHTTOLEFT:
		case DISPID_DCB_TEXTENDELLIPSIS:
		case DISPID_DCB_USESYSTEMIMAGELIST:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_DCB_BACKCOLOR:
		case DISPID_DCB_FORECOLOR:
		/*case DISPID_DCB_LISTBACKCOLOR:
		case DISPID_DCB_LISTFORECOLOR:*/
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_DCB_APPID:
		case DISPID_DCB_APPNAME:
		case DISPID_DCB_APPSHORTNAME:
		case DISPID_DCB_BUILD:
		case DISPID_DCB_CHARSET:
		case DISPID_DCB_HIMAGELIST:
		case DISPID_DCB_HWND:
		case DISPID_DCB_HWNDCOMBOBOX:
		case DISPID_DCB_HWNDLISTBOX:
		case DISPID_DCB_ISRELEASE:
		case DISPID_DCB_PROGRAMMER:
		case DISPID_DCB_TESTER:
		case DISPID_DCB_TEXT:
		case DISPID_DCB_TEXTLENGTH:
		case DISPID_DCB_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_DCB_DRAGDROPDOWNTIME:
		case DISPID_DCB_HANDLEOLEDRAGDROP:
		case DISPID_DCB_LISTDRAGSCROLLTIMEBASE:
		case DISPID_DCB_OLEDRAGIMAGESTYLE:
		case DISPID_DCB_REGISTERFOROLEDRAGDROP:
		case DISPID_DCB_SHOWDRAGIMAGE:
		case DISPID_DCB_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_DCB_FONT:
		case DISPID_DCB_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_DCB_COMBOITEMS:
		case DISPID_DCB_FIRSTVISIBLEITEM:
		//case DISPID_DCB_INTEGRALHEIGHT:
		case DISPID_DCB_SELECTEDITEM:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_DCB_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString DriveComboBox::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString DriveComboBox::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString DriveComboBox::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ComboListBoxControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString DriveComboBox::GetSpecialThanks(void)
{
	return TEXT("Wine Headquarters");
}

CAtlString DriveComboBox::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString DriveComboBox::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IMouseHookHandler
LRESULT DriveComboBox::HandleMessage(int code, WPARAM wParam, LPARAM lParam)
{
	ATLASSERT(mouseStatus_ComboBox.hHook);

	if(wParam == WM_RBUTTONUP) {
		UnhookWindowsHookEx(mouseStatus_ComboBox.hHook);
		mouseStatus_ComboBox.hHook = NULL;
		HANDLE hMem = RemoveProp(containedComboBox, HOOKPROPIDENTIFIER);
		if(hMem) {
			GlobalFree(hMem);
		}

		LPMOUSEHOOKSTRUCT pDetails = reinterpret_cast<LPMOUSEHOOKSTRUCT>(lParam);
		ATLASSERT_POINTER(pDetails, MOUSEHOOKSTRUCT);

		if(pDetails->wHitTestCode == HTCLIENT) {
			POINT mousePosition = pDetails->pt;
			containedComboBox.ScreenToClient(&mousePosition);
			BOOL dummy = TRUE;
			OnRButtonUp(1, 0, GetButtonStateBitField(), MAKELPARAM(mousePosition.x, mousePosition.y), dummy);
			dummy = TRUE;
			OnContextMenu(1, 0, reinterpret_cast<WPARAM>(static_cast<HWND>(containedComboBox)), MAKELPARAM(pDetails->pt.x, pDetails->pt.y), dummy);
		}
	}

	return CallNextHookEx(mouseStatus_ComboBox.hHook, code, wParam, lParam);
}
// implementation of IMouseHookHandler
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP DriveComboBox::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		/*case DISPID_DCB_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_DCB_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;*/
		case DISPID_DCB_CUEBANNER:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_DCB_DISPLAYNAMESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.displayNameStyle), description);
			break;
		case DISPID_DCB_DROPDOWNKEY:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.dropDownKey), description);
			break;
		case DISPID_DCB_HANDLEOLEDRAGDROP:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.handleOLEDragDrop), description);
			break;
		case DISPID_DCB_ICONVISIBILITY:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.iconVisibility), description);
			break;
		case DISPID_DCB_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_DCB_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_DCB_OLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.oleDragImageStyle), description);
			break;
		case DISPID_DCB_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_DCB_USESYSTEMIMAGELIST:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.useSystemImageList), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<DriveComboBox>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP DriveComboBox::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		//case DISPID_DCB_BORDERSTYLE:
		case DISPID_DCB_DROPDOWNKEY:
		case DISPID_DCB_OLEDRAGIMAGESTYLE:
			c = 2;
			break;
		//case DISPID_DCB_APPEARANCE:
		case DISPID_DCB_ICONVISIBILITY:
			c = 3;
			break;
		case DISPID_DCB_HANDLEOLEDRAGDROP:
		case DISPID_DCB_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_DCB_DISPLAYNAMESTYLE:
		case DISPID_DCB_USESYSTEMIMAGELIST:
			c = 5;
			break;
		case DISPID_DCB_IMEMODE:
			c = 12;
			break;
		case DISPID_DCB_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<DriveComboBox>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_DCB_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		/*case DISPID_DCB_APPEARANCE:
		case DISPID_DCB_BORDERSTYLE:*/
		case DISPID_DCB_DISPLAYNAMESTYLE:
		case DISPID_DCB_DROPDOWNKEY:
		case DISPID_DCB_HANDLEOLEDRAGDROP:
		case DISPID_DCB_ICONVISIBILITY:
		case DISPID_DCB_IMEMODE:
		case DISPID_DCB_MOUSEPOINTER:
		case DISPID_DCB_OLEDRAGIMAGESTYLE:
		case DISPID_DCB_RIGHTTOLEFT:
		case DISPID_DCB_USESYSTEMIMAGELIST:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<DriveComboBox>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_DCB_CUEBANNER:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<DriveComboBox>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT DriveComboBox::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		/*case DISPID_DCB_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;*/
		case DISPID_DCB_DISPLAYNAMESTYLE:
			switch(cookie) {
				case dnsDriveLetter:
					description = GetResStringWithNumber(IDP_DISPLAYNAMESTYLEDRIVELETTER, dnsDriveLetter);
					break;
				case dnsDriveLetterWithColon:
					description = GetResStringWithNumber(IDP_DISPLAYNAMESTYLEDRIVELETTERCOLON, dnsDriveLetterWithColon);
					break;
				case dnsVolumeName:
					description = GetResStringWithNumber(IDP_DISPLAYNAMESTYLEVOLUMENAME, dnsVolumeName);
					break;
				case dnsVB6Style:
					description = GetResStringWithNumber(IDP_DISPLAYNAMESTYLEVB6STYLE, dnsVB6Style);
					break;
				case dnsShellStyle:
					description = GetResStringWithNumber(IDP_DISPLAYNAMESTYLESHELLSTYLE, dnsShellStyle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_DROPDOWNKEY:
			switch(cookie) {
				case ddkF4:
					description = GetResStringWithNumber(IDP_DROPDOWNKEYF4, ddkF4);
					break;
				case ddkDownArrow:
					description = GetResStringWithNumber(IDP_DROPDOWNKEYDOWNARROW, ddkDownArrow);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_HANDLEOLEDRAGDROP:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPNONE, 0);
					break;
				case hoddSourcePart:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPSOURCEPART, hoddSourcePart);
					break;
				case hoddTargetPart:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPTARGETPART, hoddTargetPart);
					break;
				case hoddSourcePart | hoddTargetPart:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPBOTHPARTS, hoddSourcePart | hoddTargetPart);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_ICONVISIBILITY:
			switch(cookie) {
				case ivVisible:
					description = GetResStringWithNumber(IDP_ICONVISIBILITYVISIBLE, ivVisible);
					break;
				case ivHiddenButIndent:
					description = GetResStringWithNumber(IDP_ICONVISIBILITYHIDDENBUTINDENT, ivHiddenButIndent);
					break;
				case ivHiddenDontIndent:
					description = GetResStringWithNumber(IDP_ICONVISIBILITYHIDDENDONTINDENT, ivHiddenDontIndent);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_IMEMODE:
			switch(cookie) {
				case imeInherit:
					description = GetResStringWithNumber(IDP_IMEMODEINHERIT, imeInherit);
					break;
				case imeNoControl:
					description = GetResStringWithNumber(IDP_IMEMODENOCONTROL, imeNoControl);
					break;
				case imeOn:
					description = GetResStringWithNumber(IDP_IMEMODEON, imeOn);
					break;
				case imeOff:
					description = GetResStringWithNumber(IDP_IMEMODEOFF, imeOff);
					break;
				case imeDisable:
					description = GetResStringWithNumber(IDP_IMEMODEDISABLE, imeDisable);
					break;
				case imeHiragana:
					description = GetResStringWithNumber(IDP_IMEMODEHIRAGANA, imeHiragana);
					break;
				case imeKatakana:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANA, imeKatakana);
					break;
				case imeKatakanaHalf:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANAHALF, imeKatakanaHalf);
					break;
				case imeAlphaFull:
					description = GetResStringWithNumber(IDP_IMEMODEALPHAFULL, imeAlphaFull);
					break;
				case imeAlpha:
					description = GetResStringWithNumber(IDP_IMEMODEALPHA, imeAlpha);
					break;
				case imeHangulFull:
					description = GetResStringWithNumber(IDP_IMEMODEHANGULFULL, imeHangulFull);
					break;
				case imeHangul:
					description = GetResStringWithNumber(IDP_IMEMODEHANGUL, imeHangul);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_OLEDRAGIMAGESTYLE:
			switch(cookie) {
				case odistClassic:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLECLASSIC, odistClassic);
					break;
				case odistAeroIfAvailable:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLEAERO, odistAeroIfAvailable);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DCB_USESYSTEMIMAGELIST:
			switch(cookie) {
				case usilNone:
					description = GetResStringWithNumber(IDP_USESYSTEMIMAGELISTNONE, usilNone);
					break;
				case usilSmallImageList:
					description = GetResStringWithNumber(IDP_USESYSTEMIMAGELISTSMALL, usilSmallImageList);
					break;
				case usilLargeImageList:
					description = GetResStringWithNumber(IDP_USESYSTEMIMAGELISTLARGE, usilLargeImageList);
					break;
				case usilExtraLargeImageList:
					description = GetResStringWithNumber(IDP_USESYSTEMIMAGELISTEXTRALARGE, usilExtraLargeImageList);
					break;
				case usilJumboImageList:
					description = GetResStringWithNumber(IDP_USESYSTEMIMAGELISTJUMBO, usilJumboImageList);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP DriveComboBox::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 6;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_DriveTypeProperties;
		pPropertyPages->pElems[3] = CLSID_StockColorPage;
		pPropertyPages->pElems[4] = CLSID_StockFontPage;
		pPropertyPages->pElems[5] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP DriveComboBox::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<DriveComboBox>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP DriveComboBox::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP DriveComboBox::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT DriveComboBox::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT DriveComboBox::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_DCB_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT DriveComboBox::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


/*STDMETHODIMP DriveComboBox::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_APPEARANCE);
	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			SendMessage(WM_WINDOWPOSCHANGING, 0, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_DCB_APPEARANCE);
	}
	return S_OK;
}*/

STDMETHODIMP DriveComboBox::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 16;
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ComboListBoxControls");
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"CBLCtls");
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DCB_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(containedComboBox.IsWindow()) {
			containedComboBox.Invalidate();
		}
		FireViewChange();
		FireOnChanged(DISPID_DCB_BACKCOLOR);
	}
	return S_OK;
}

/*STDMETHODIMP DriveComboBox::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			SendMessage(WM_WINDOWPOSCHANGING, 0, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_DCB_BORDERSTYLE);
	}
	return S_OK;
}*/

STDMETHODIMP DriveComboBox::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_CaseSensitiveItemSearching(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.caseSensitiveItemSearching = ((SendMessage(CBEM_GETEXTENDEDSTYLE, 0, 0) & CBES_EX_CASESENSITIVE) == CBES_EX_CASESENSITIVE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveItemSearching);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_CaseSensitiveItemSearching(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_CASESENSITIVEITEMSEARCHING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.caseSensitiveItemSearching != b) {
		properties.caseSensitiveItemSearching = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.caseSensitiveItemSearching) {
				SendMessage(CBEM_SETEXTENDEDSTYLE, CBES_EX_CASESENSITIVE, CBES_EX_CASESENSITIVE);
			} else {
				SendMessage(CBEM_SETEXTENDEDSTYLE, CBES_EX_CASESENSITIVE, 0);
			}
		}
		FireOnChanged(DISPID_DCB_CASESENSITIVEITEMSEARCHING);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_ComboItems(IDriveComboBoxItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IDriveComboBoxItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	ClassFactory::InitDriveComboItems(this, IID_IDriveComboBoxItems, reinterpret_cast<LPUNKNOWN*>(ppItems));
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_CueBanner(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedComboBox.IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		WCHAR pBuffer[1025];
		// NOTE: CB_GETCUEBANNER isn't forwarded by ComboBoxEx32
		if(containedComboBox.SendMessage(CB_GETCUEBANNER, reinterpret_cast<WPARAM>(pBuffer), 1024)) {
			properties.cueBanner = pBuffer;
		}
	}

	*pValue = properties.cueBanner.Copy();
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_CueBanner(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_DCB_CUEBANNER);
	if(properties.cueBanner != newValue) {
		properties.cueBanner.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(RunTimeHelper::IsCommCtrl6()) {
			// NOTE: CB_SETCUEBANNER isn't forwarded by ComboBoxEx32
			if(containedComboBox.IsWindow()) {
				containedComboBox.SendMessage(CB_SETCUEBANNER, 0, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
			}
		}
		FireOnChanged(DISPID_DCB_CUEBANNER);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deMouseEvents) != (newValue & deMouseEvents)) {
			if(IsWindow()) {
				if(newValue & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					trackingOptions.hwndTrack = *this;
					TrackMouseEvent(&trackingOptions);
					if(containedComboBox.IsWindow()) {
						trackingOptions.hwndTrack = containedComboBox;
						TrackMouseEvent(&trackingOptions);
					}
				}
			}
		}
		if((properties.disabledEvents & deListBoxMouseEvents) != (newValue & deListBoxMouseEvents)) {
			if(containedListBox.IsWindow()) {
				if(newValue & deListBoxMouseEvents) {
					// nothing to do
				} else {
					itemUnderMouse = -1;
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DisplayedDriveTypes(DriveTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DriveTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.displayedDriveTypes;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DisplayedDriveTypes(DriveTypeConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DISPLAYEDDRIVETYPES);
	if(properties.displayedDriveTypes != newValue) {
		properties.displayedDriveTypes = newValue;
		ReloadDrives();
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_DISPLAYEDDRIVETYPES);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DisplayNameStyle(DisplayNameStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisplayNameStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.displayNameStyle;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DisplayNameStyle(DisplayNameStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DISPLAYNAMESTYLE);
	if(properties.displayNameStyle != newValue) {
		properties.displayNameStyle = newValue;
		ReloadDrives();
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_DISPLAYNAMESTYLE);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DisplayOverlayImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayOverlayImages);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DisplayOverlayImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DISPLAYOVERLAYIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayOverlayImages != b) {
		properties.displayOverlayImages = b;
		ReloadDrives();
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_DISPLAYOVERLAYIMAGES);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_DCB_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DoOEMConversion(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.doOEMConversion = ((GetStyle() & CBS_OEMCONVERT) == CBS_OEMCONVERT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.doOEMConversion);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DoOEMConversion(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DOOEMCONVERSION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.doOEMConversion != b) {
		properties.doOEMConversion = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_DCB_DOOEMCONVERSION);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DragDropDownTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragDropDownTime;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DragDropDownTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DRAGDROPDOWNTIME);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragDropDownTime != newValue) {
		properties.dragDropDownTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_DRAGDROPDOWNTIME);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DriveTypesWithVolumeName(DriveTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DriveTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.driveTypesWithVolumeName;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DriveTypesWithVolumeName(DriveTypeConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DRIVETYPESWITHVOLUMENAME);
	if(properties.driveTypesWithVolumeName != newValue) {
		properties.driveTypesWithVolumeName = newValue;
		ReloadDrives();
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_DRIVETYPESWITHVOLUMENAME);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_DropDownButtonObjectState(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedComboBox.IsWindow()) {
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(containedComboBox, &controlInfo);
		*pValue = controlInfo.stateButton;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::get_DropDownKey(DropDownKeyConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DropDownKeyConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.dropDownKey = (SendMessage(CB_GETEXTENDEDUI, 0, 0) ? ddkDownArrow : ddkF4);
	}

	*pValue = properties.dropDownKey;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_DropDownKey(DropDownKeyConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_DROPDOWNKEY);
	if(properties.dropDownKey != newValue) {
		properties.dropDownKey = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.dropDownKey) {
				case ddkF4:
					SendMessage(CB_SETEXTENDEDUI, FALSE, 0);
					break;
				case ddkDownArrow:
					SendMessage(CB_SETEXTENDEDUI, TRUE, 0);
					break;
			}
		}
		FireOnChanged(DISPID_DCB_DROPDOWNKEY);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_DCB_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_FirstVisibleItem(IDriveComboBoxItem** ppFirstItem)
{
	ATLASSERT_POINTER(ppFirstItem, IDriveComboBoxItem*);
	if(!ppFirstItem) {
		return E_POINTER;
	}

	if(containedComboBox.IsWindow()) {
		// NOTE: CB_GETTOPINDEX isn't forwarded by ComboBoxEx32
		ClassFactory::InitDriveComboItem(static_cast<int>(containedComboBox.SendMessage(CB_GETTOPINDEX, 0, 0)), this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFirstItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::putref_FirstVisibleItem(IDriveComboBoxItem* pNewFirstItem)
{
	PUTPROPPROLOG(DISPID_DCB_FIRSTVISIBLEITEM);
	HRESULT hr = E_FAIL;

	int newFirstItem = -1;
	if(pNewFirstItem) {
		LONG l = -1;
		pNewFirstItem->get_Index(&l);
		newFirstItem = l;
		// TODO: Shouldn't we AddRef' pNewFirstItem?
	}

	if(containedComboBox.IsWindow()) {
		// NOTE: CB_SETTOPINDEX isn't forwarded by ComboBoxEx32
		containedComboBox.SendMessage(CB_SETTOPINDEX, newFirstItem, 0);
		hr = S_OK;
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_FIRSTVISIBLEITEM);
	FireViewChange();
	return hr;
}

STDMETHODIMP DriveComboBox::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_DCB_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_FONT);
	return S_OK;
}

STDMETHODIMP DriveComboBox::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_DCB_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_FONT);
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DCB_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		if(containedComboBox.IsWindow()) {
			containedComboBox.Invalidate();
		}
		FireViewChange();
		FireOnChanged(DISPID_DCB_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_HandleOLEDragDrop(HandleOLEDragDropConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HandleOLEDragDropConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.handleOLEDragDrop;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_HandleOLEDragDrop(HandleOLEDragDropConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_HANDLEOLEDRAGDROP);
	if(properties.handleOLEDragDrop != newValue) {
		properties.handleOLEDragDrop = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_HANDLEOLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = NULL;
	switch(imageList) {
		case ilItems:
			if(IsWindow()) {
				*pValue = HandleToLong(reinterpret_cast<HIMAGELIST>(SendMessage(CBEM_GETIMAGELIST, 0, 0)));
			}
			break;
		case ilHighResolution:
			*pValue = HandleToLong(properties.hHighResImageList);
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_DCB_HIMAGELIST);
	BOOL fireViewChange = TRUE;
	switch(imageList) {
		case ilItems:
			if(IsWindow()) {
				SendMessage(CBEM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(LongToHandle(newValue)));
			}
			break;
		case ilHighResolution:
			properties.hHighResImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
			fireViewChange = FALSE;
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	FireOnChanged(DISPID_DCB_HIMAGELIST);
	if(fireViewChange) {
		FireViewChange();
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_DCB_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_hWndComboBox(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		//*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(CBEM_GETCOMBOCONTROL, 0, 0)));
		*pValue = HandleToLong(containedComboBox.m_hWnd);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::get_hWndListBox(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(containedListBox.m_hWnd);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_hWndShellUIParentWindow(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.hWndShellUIParentWindow);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_hWndShellUIParentWindow(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_DCB_HWNDSHELLUIPARENTWINDOW);
	if(properties.hWndShellUIParentWindow != static_cast<HWND>(LongToHandle(newValue))) {
		properties.hWndShellUIParentWindow = static_cast<HWND>(LongToHandle(newValue));
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_HWNDSHELLUIPARENTWINDOW);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_IconVisibility(IconVisibilityConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IconVisibilityConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(SendMessage(CBEM_GETEXTENDEDSTYLE, 0, 0) & (CBES_EX_NOEDITIMAGE | CBES_EX_NOEDITIMAGEINDENT)) {
			case CBES_EX_NOEDITIMAGE:
				properties.iconVisibility = ivHiddenButIndent;
				break;
			case CBES_EX_NOEDITIMAGEINDENT:
				properties.iconVisibility = ivHiddenDontIndent;
				break;
			default:
				properties.iconVisibility = ivVisible;
				break;
		}
	}

	*pValue = properties.iconVisibility;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_IconVisibility(IconVisibilityConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_ICONVISIBILITY);
	if(properties.iconVisibility != newValue) {
		properties.iconVisibility = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.iconVisibility) {
				case ivVisible:
					SendMessage(CBEM_SETEXTENDEDSTYLE, CBES_EX_NOEDITIMAGE | CBES_EX_NOEDITIMAGEINDENT, 0);
					break;
				case ivHiddenButIndent:
					SendMessage(CBEM_SETEXTENDEDSTYLE, CBES_EX_NOEDITIMAGE | CBES_EX_NOEDITIMAGEINDENT, CBES_EX_NOEDITIMAGE);
					break;
				case ivHiddenDontIndent:
					SendMessage(CBEM_SETEXTENDEDSTYLE, CBES_EX_NOEDITIMAGE | CBES_EX_NOEDITIMAGEINDENT, CBES_EX_NOEDITIMAGEINDENT);
					break;
			}
		}
		FireViewChange();
		FireOnChanged(DISPID_DCB_ICONVISIBILITY);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_IMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.IMEMode;
	} else {
		if((GetFocus() == *this) && (GetEffectiveIMEMode() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(*this);
			if((ime != imeInherit) && (properties.IMEMode != imeInherit)) {
				properties.IMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode();
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			HWND h = GetFocus();
			if(h == *this || h == containedComboBox || h == containedListBox) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(h, ime);
				}
			}
		}
		FireOnChanged(DISPID_DCB_IMEMODE);
	}
	return S_OK;
}

/*STDMETHODIMP DriveComboBox::get_IntegralHeight(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedComboBox.IsWindow()) {
		properties.integralHeight = !(containedComboBox.GetStyle() & CBS_NOINTEGRALHEIGHT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.integralHeight);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_IntegralHeight(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_INTEGRALHEIGHT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.integralHeight != b) {
		properties.integralHeight = b;
		SetDirty(TRUE);

		if(containedComboBox.IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_DCB_INTEGRALHEIGHT);
	}
	return S_OK;
}*/

STDMETHODIMP DriveComboBox::get_IsDroppedDown(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = BOOL2VARIANTBOOL(SendMessage(CB_GETDROPPEDSTATE, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_ItemHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.itemHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(CB_GETITEMHEIGHT, 0, 0));
	}

	*pValue = properties.itemHeight;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ItemHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_DCB_ITEMHEIGHT);
	if(properties.itemHeight != newValue) {
		properties.itemHeight = newValue;
		properties.setItemHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.itemHeight == -1) {
				ApplyFont();
				/*Called in OnSetFont: if(properties.setSelectionFieldHeight != -1) {
					SendMessage(CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), properties.setSelectionFieldHeight);
				}*/
			} else {
				SendMessage(CB_SETITEMHEIGHT, 0, properties.itemHeight);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_DCB_ITEMHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_ListAlwaysShowVerticalScrollBar(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.listAlwaysShowVerticalScrollBar = ((GetStyle() & CBS_DISABLENOSCROLL) == CBS_DISABLENOSCROLL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.listAlwaysShowVerticalScrollBar);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ListAlwaysShowVerticalScrollBar(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_LISTALWAYSSHOWVERTICALSCROLLBAR);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.listAlwaysShowVerticalScrollBar != b) {
		properties.listAlwaysShowVerticalScrollBar = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_DCB_LISTALWAYSSHOWVERTICALSCROLLBAR);
	}
	return S_OK;
}

/*STDMETHODIMP DriveComboBox::get_ListBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.listBackColor;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ListBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DCB_LISTBACKCOLOR);
	if(properties.listBackColor != newValue) {
		properties.listBackColor = newValue;
		SetDirty(TRUE);

		if(hListBackColorBrush) {
			DeleteObject(hListBackColorBrush);
			hListBackColorBrush = NULL;
		}
		if(!(properties.listBackColor & 0x80000000)) {
			hListBackColorBrush = CreateSolidBrush(OLECOLOR2COLORREF(properties.listBackColor));
		}
		if(containedListBox.IsWindow()) {
			containedListBox.Invalidate();
		}
		FireViewChange();
		FireOnChanged(DISPID_DCB_LISTBACKCOLOR);
	}
	return S_OK;
}*/

STDMETHODIMP DriveComboBox::get_ListDragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.listDragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ListDragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_DCB_LISTDRAGSCROLLTIMEBASE);

	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.listDragScrollTimeBase != newValue) {
		properties.listDragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_LISTDRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

/*STDMETHODIMP DriveComboBox::get_ListForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.listForeColor;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ListForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DCB_LISTFORECOLOR);
	if(properties.listForeColor != newValue) {
		properties.listForeColor = newValue;
		SetDirty(TRUE);

		if(containedListBox.IsWindow()) {
			containedListBox.Invalidate();
		}
		FireViewChange();
		FireOnChanged(DISPID_DCB_LISTFORECOLOR);
	}
	return S_OK;
}*/

STDMETHODIMP DriveComboBox::get_ListHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedListBox.IsWindow()) {
		CRect rc;
		containedListBox.GetWindowRect(&rc);
		properties.listHeight = rc.Height();
	}

	*pValue = properties.listHeight;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ListHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_DCB_LISTHEIGHT);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.listHeight != newValue) {
		properties.listHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			CRect rc;
			containedComboBox.GetWindowRect(&rc);
			if(properties.listHeight == -1) {
				// make room for 8 items
				int itemHeight = properties.itemHeight;
				if(itemHeight == -1) {
					itemHeight = static_cast<int>(SendMessage(CB_GETITEMHEIGHT, 0, 0));
				}
				rc.bottom += 8 * itemHeight + 2 * GetSystemMetrics(SM_CYBORDER);
			} else {
				rc.bottom += properties.listHeight;
			}
			ScreenToClient(&rc);
			containedComboBox.MoveWindow(&rc, FALSE);
		}
		FireOnChanged(DISPID_DCB_LISTHEIGHT);
	}
	return S_OK;
}

/*STDMETHODIMP DriveComboBox::get_ListScrollableWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedComboBox.IsWindow()) {
		// NOTE: CB_GETHORIZONTALEXTENT isn't forwarded by ComboBoxEx32
		properties.listScrollableWidth = static_cast<LONG>(containedComboBox.SendMessage(CB_GETHORIZONTALEXTENT, 0, 0));
	}

	*pValue = properties.listScrollableWidth;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ListScrollableWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_DCB_LISTSCROLLABLEWIDTH);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.listScrollableWidth != newValue) {
		properties.listScrollableWidth = newValue;
		SetDirty(TRUE);

		if(containedComboBox.IsWindow()) {
			// NOTE: CB_SETHORIZONTALEXTENT isn't forwarded by ComboBoxEx32
			containedComboBox.SendMessage(CB_SETHORIZONTALEXTENT, properties.listScrollableWidth, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_DCB_LISTSCROLLABLEWIDTH);
	}
	return S_OK;
}*/

STDMETHODIMP DriveComboBox::get_ListWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedComboBox.IsWindow()) {
		// NOTE: CB_GETDROPPEDWIDTH isn't forwarded by ComboBoxEx32
		properties.listWidth = static_cast<LONG>(containedComboBox.SendMessage(CB_GETDROPPEDWIDTH, 0, 0));
	}

	*pValue = properties.listWidth;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ListWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_DCB_LISTWIDTH);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.listWidth != newValue) {
		properties.listWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(CB_SETDROPPEDWIDTH, properties.listWidth, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_DCB_LISTWIDTH);
	}
	return S_OK;
}

/*STDMETHODIMP DriveComboBox::get_MinVisibleItems(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && RunTimeHelper::IsCommCtrl6()) {
		properties.minVisibleItems = static_cast<LONG>(SendMessage(CB_GETMINVISIBLE, 0, 0));
	}

	*pValue = properties.minVisibleItems;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_MinVisibleItems(LONG newValue)
{
	PUTPROPPROLOG(DISPID_DCB_MINVISIBLEITEMS);
	if(newValue < 1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.minVisibleItems != newValue) {
		properties.minVisibleItems = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(CB_SETMINVISIBLE, properties.minVisibleItems, 0);
		}
		FireOnChanged(DISPID_DCB_MINVISIBLEITEMS);
	}
	return S_OK;
}*/

STDMETHODIMP DriveComboBox::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_DCB_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP DriveComboBox::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_DCB_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.oleDragImageStyle;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_OLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_OLEDRAGIMAGESTYLE);
	if(properties.oleDragImageStyle != newValue) {
		properties.oleDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_OLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
				if(containedListBox.IsWindow()) {
					ATLVERIFY(RegisterDragDrop(containedListBox, static_cast<IDropTarget*>(this)) == S_OK);
				}
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
				if(containedListBox.IsWindow()) {
					ATLVERIFY(RevokeDragDrop(containedListBox) == S_OK);
				}
			}
		}
		FireOnChanged(DISPID_DCB_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		DWORD style = GetExStyle();
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_DCB_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_SelectedItem(IDriveComboBoxItem** ppSelectedItem)
{
	ATLASSERT_POINTER(ppSelectedItem, IDriveComboBoxItem*);
	if(!ppSelectedItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		int itemIndex = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		if(itemIndex != CB_ERR) {
			ClassFactory::InitDriveComboItem(itemIndex, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppSelectedItem));
			return S_OK;
		} else {
			*ppSelectedItem = NULL;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::putref_SelectedItem(IDriveComboBoxItem* pNewSelectedItem)
{
	PUTPROPPROLOG(DISPID_DCB_SELECTEDITEM);
	int newSelectedItem = -1;
	if(pNewSelectedItem) {
		LONG l = -1;
		pNewSelectedItem->get_Index(&l);
		newSelectedItem = l;
		// TODO: Shouldn't we AddRef' pNewSelectedItem?
	}

	if(IsWindow()) {
		SendMessage(CB_SETCURSEL, newSelectedItem, 0);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_SELECTEDITEM);
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_SelectionFieldHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.selectionFieldHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(CB_GETITEMHEIGHT, static_cast<WPARAM>(-1), 0));
	}

	*pValue = properties.selectionFieldHeight;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_SelectionFieldHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_DCB_SELECTIONFIELDHEIGHT);
	if(properties.selectionFieldHeight != newValue) {
		properties.selectionFieldHeight = newValue;
		properties.setSelectionFieldHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.selectionFieldHeight == -1) {
				ApplyFont();
				/*Called in OnSetFont: if(properties.setItemHeight != -1) {
					SendMessage(CB_SETITEMHEIGHT, 0, properties.setItemHeight);
				}*/
			} else {
				SendMessage(CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), properties.selectionFieldHeight);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_DCB_SELECTIONFIELDHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_SelectionFieldItem(IDriveComboBoxItem** ppSelectionFieldItem)
{
	ATLASSERT_POINTER(ppSelectionFieldItem, IDriveComboBoxItem*);
	if(!ppSelectionFieldItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		ClassFactory::InitDriveComboItem(-1, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppSelectionFieldItem), FALSE);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_SHOWDRAGIMAGE);
	if(!dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage(FALSE);
	} else {
		dragDropStatus.ShowDragImage(FALSE);
	}

	FireOnChanged(DISPID_DCB_SHOWDRAGIMAGE);
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if((GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
			properties.text = L"";
			CComPtr<IDriveComboBoxItem> pDCBoxItem;
			get_SelectionFieldItem(&pDCBoxItem);
			if(pDCBoxItem) {
				pDCBoxItem->get_Text(&properties.text);
			}
		} else {
			properties.text = L"";
			int textLength = static_cast<int>(SendMessage(CB_GETLBTEXTLEN, static_cast<WPARAM>(-1), 0));
			if(textLength != CB_ERR) {
				LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
				if(pBuffer) {
					ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
					SendMessage(CB_GETLBTEXT, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(pBuffer));
					properties.text = pBuffer;
					HeapFree(GetProcessHeap(), 0, pBuffer);
				}
			}
		}
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_DCB_TEXT);
	properties.text.AssignBSTR(newValue);
	if(IsWindow()) {
		SetWindowText(COLE2CT(properties.text));
		// TODO: Select by text for CBS_DROPDOWNLIST.
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_TEXT);
	SendOnDataChange();
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_TextEndEllipsis(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	
	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.textEndEllipsis = ((SendMessage(CBEM_GETEXTENDEDSTYLE, 0, 0) & CBES_EX_TEXTENDELLIPSIS) == CBES_EX_TEXTENDELLIPSIS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.textEndEllipsis);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_TextEndEllipsis(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_TEXTENDELLIPSIS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.textEndEllipsis != b) {
		properties.textEndEllipsis = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.textEndEllipsis) {
				SendMessage(CBEM_SETEXTENDEDSTYLE, CBES_EX_TEXTENDELLIPSIS, CBES_EX_TEXTENDELLIPSIS);
			} else {
				SendMessage(CBEM_SETEXTENDEDSTYLE, CBES_EX_TEXTENDELLIPSIS, 0);
			}
		}
		FireOnChanged(DISPID_DCB_TEXTENDELLIPSIS);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_TextLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if((GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
			CComBSTR text = L"";
			CComPtr<IDriveComboBoxItem> pDCBoxItem;
			get_SelectionFieldItem(&pDCBoxItem);
			if(pDCBoxItem) {
				pDCBoxItem->get_Text(&text);
				*pValue = text.Length();
			}
		} else {
			*pValue = static_cast<LONG>(SendMessage(CB_GETLBTEXTLEN, static_cast<WPARAM>(-1), 0));
		}
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DCB_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_DCB_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_UseSystemImageList(UseSystemImageListConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UseSystemImageListConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.useSystemImageList;
	return S_OK;
}

STDMETHODIMP DriveComboBox::put_UseSystemImageList(UseSystemImageListConstants newValue)
{
	PUTPROPPROLOG(DISPID_DCB_USESYSTEMIMAGELIST);
	if(properties.useSystemImageList != newValue) {
		properties.useSystemImageList = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(newValue == usilNone) {
				// switching from system image list to custom image list - clear the current image list
				SendMessage(CBEM_SETIMAGELIST, 0, NULL);
				properties.hHighResImageList = NULL;
			}
			SetSystemImageList();
		}
		FireOnChanged(DISPID_DCB_USESYSTEMIMAGELIST);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP DriveComboBox::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP DriveComboBox::CloseDropDownWindow(void)
{
	if(IsWindow()) {
		SendMessage(CB_SHOWDROPDOWN, FALSE, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::CreateItemContainer(VARIANT items/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IDriveComboBoxItemContainer** ppContainer/* = NULL*/)
{
	ATLASSERT_POINTER(ppContainer, IDriveComboBoxItemContainer*);
	if(!ppContainer) {
		return E_POINTER;
	}

	*ppContainer = NULL;
	CComObject<DriveComboBoxItemContainer>* pDCBItemContainerObj = NULL;
	CComObject<DriveComboBoxItemContainer>::CreateInstance(&pDCBItemContainerObj);
	pDCBItemContainerObj->AddRef();

	// clone all settings
	pDCBItemContainerObj->SetOwner(this);

	pDCBItemContainerObj->QueryInterface(__uuidof(IDriveComboBoxItemContainer), reinterpret_cast<LPVOID*>(ppContainer));
	pDCBItemContainerObj->Release();

	if(*ppContainer) {
		(*ppContainer)->Add(items);
		RegisterItemContainer(static_cast<IItemContainer*>(pDCBItemContainerObj));
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP DriveComboBox::FindItemByItemData(LONG itemData, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IDriveComboBoxItem** ppFoundItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppFoundItem, IDriveComboBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IDriveComboBoxItem> pItem = startAfterItem.pdispVal;
			if(pItem) {
				pItem->get_Index(&itemIndex);
			}
		} else {
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &startAfterItem, 0, VT_UINT))) {
				itemIndex = v.uintVal;
			}
		}
		if(itemIndex == -1) {
			return E_INVALIDARG;
		}
	}

	int numberOfItems = static_cast<int>(SendMessage(CB_GETCOUNT, 0, 0));
	for(int i = itemIndex + 1; i < numberOfItems; i++) {
		if(static_cast<LONG>(SendMessage(CB_GETITEMDATA, i, 0)) == itemData) {
			ClassFactory::InitDriveComboItem(i, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
			return S_OK;
		}
	}
	for(int i = 0; i <= itemIndex; i++) {
		if(static_cast<LONG>(SendMessage(CB_GETITEMDATA, i, 0)) == itemData) {
			ClassFactory::InitDriveComboItem(i, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
			return S_OK;
		}
	}
	*ppFoundItem = NULL;
	return S_OK;
}

STDMETHODIMP DriveComboBox::FindItemByPath(BSTR path, IDriveComboBoxItem** ppFoundItem)
{
	ATLASSERT_POINTER(ppFoundItem, IDriveComboBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	COLE2T converter(path);
	LPTSTR pSearchString = converter;
	int searchStringLength = lstrlen(pSearchString);
	LCID locale = GetThreadLocale();

	int c = static_cast<int>(SendMessage(CB_GETCOUNT, 0, 0));
	for(int i = 0; i < c; i++) {
		int driveIndex = GetDriveIndex(i);
		if(driveIndex >= 0) {
			TCHAR pDrive[4] = TEXT("A:\\");
			pDrive[0] += static_cast<TCHAR>(driveIndex);

			int result = CompareString(locale, NORM_IGNORECASE, pSearchString, searchStringLength, pDrive, 3);
			if(result == CSTR_EQUAL) {
				ClassFactory::InitDriveComboItem(i, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
				return S_OK;
			}
		}
	}
	*ppFoundItem = NULL;
	return S_OK;
}

STDMETHODIMP DriveComboBox::FindItemByText(BSTR searchString, VARIANT_BOOL exactMatch/* = VARIANT_TRUE*/, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IDriveComboBoxItem** ppFoundItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppFoundItem, IDriveComboBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IDriveComboBoxItem> pItem = startAfterItem.pdispVal;
			if(pItem) {
				pItem->get_Index(&itemIndex);
			}
		} else {
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &startAfterItem, 0, VT_UINT))) {
				itemIndex = v.uintVal;
			}
		}
		if(itemIndex == -1) {
			return E_INVALIDARG;
		}
	}

	COLE2T converter(searchString);
	LPTSTR pSearchString = converter;
	int foundItem = -1;
	if(exactMatch == VARIANT_FALSE) {
		// ComboBoxEx32 doesn't support CB_FINDSTRING
		HANDLE hHeap = GetProcessHeap();
		int searchStringLength = lstrlen(pSearchString);
		LCID locale = GetThreadLocale();
		BOOL caseSensitive = ((SendMessage(CBEM_GETEXTENDEDSTYLE, 0, 0) & CBES_EX_CASESENSITIVE) == CBES_EX_CASESENSITIVE);

		int numberOfItems = static_cast<int>(SendMessage(CB_GETCOUNT, 0, 0));
		for(int i = itemIndex + 1; i < numberOfItems; i++) {
			int textLength = static_cast<int>(SendMessage(CB_GETLBTEXTLEN, i, 0));
			if(textLength == CB_ERR) {
				return DISP_E_BADINDEX;
			}
			if(textLength < searchStringLength) {
				continue;
			}
			LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(hHeap, 0, (textLength + 1) * sizeof(TCHAR)));
			if(!pBuffer) {
				return E_OUTOFMEMORY;
			}
			ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
			SendMessage(CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(pBuffer));
			int result = 0;
			if(caseSensitive) {
				result = CompareString(locale, 0, pSearchString, searchStringLength, pBuffer, searchStringLength);
			} else {
				result = CompareString(locale, NORM_IGNORECASE, pSearchString, searchStringLength, pBuffer, searchStringLength);
			}
			HeapFree(hHeap, 0, pBuffer);

			if(result == CSTR_EQUAL) {
				foundItem = i;
				break;
			}
		}
		if(foundItem == -1) {
			for(int i = 0; i <= itemIndex; i++) {
				int textLength = static_cast<int>(SendMessage(CB_GETLBTEXTLEN, i, 0));
				if(textLength == CB_ERR) {
					return DISP_E_BADINDEX;
				}
				if(textLength < searchStringLength) {
					continue;
				}
				LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(hHeap, 0, (textLength + 1) * sizeof(TCHAR)));
				if(!pBuffer) {
					return E_OUTOFMEMORY;
				}
				ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
				SendMessage(CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(pBuffer));
				int result = 0;
				if(caseSensitive) {
					result = CompareString(locale, 0, pSearchString, searchStringLength, pBuffer, searchStringLength);
				} else {
					result = CompareString(locale, NORM_IGNORECASE, pSearchString, searchStringLength, pBuffer, searchStringLength);
				}
				HeapFree(hHeap, 0, pBuffer);

				if(result == CSTR_EQUAL) {
					foundItem = i;
					break;
				}
			}
		}
	} else {
		foundItem = static_cast<int>(SendMessage(CB_FINDSTRINGEXACT, itemIndex, reinterpret_cast<LPARAM>(pSearchString)));
	}

	ClassFactory::InitDriveComboItem(foundItem, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
	return S_OK;
}

STDMETHODIMP DriveComboBox::GetDropDownButtonRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(containedComboBox.IsWindow()) {
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(containedComboBox, &controlInfo);
		containedComboBox.MapWindowPoints(*this, &controlInfo.rcItem);
		if(pLeft) {
			*pLeft = controlInfo.rcButton.left;
		}
		if(pTop) {
			*pTop = controlInfo.rcButton.top;
		}
		if(pRight) {
			*pRight = controlInfo.rcButton.right;
		}
		if(pBottom) {
			*pBottom = controlInfo.rcButton.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::GetDroppedStateRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(IsWindow()) {
		RECT boundingRectangle = {0};
		SendMessage(CB_GETDROPPEDCONTROLRECT, 0, reinterpret_cast<LPARAM>(&boundingRectangle));
		if(pLeft) {
			*pLeft = boundingRectangle.left;
		}
		if(pTop) {
			*pTop = boundingRectangle.top;
		}
		if(pRight) {
			*pRight = boundingRectangle.right;
		}
		if(pBottom) {
			*pBottom = boundingRectangle.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::GetSelectionFieldRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(containedComboBox.IsWindow()) {
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(containedComboBox, &controlInfo);
		containedComboBox.MapWindowPoints(*this, &controlInfo.rcItem);
		if(pLeft) {
			*pLeft = controlInfo.rcItem.left;
		}
		if(pTop) {
			*pTop = controlInfo.rcItem.top;
		}
		if(pRight) {
			*pRight = controlInfo.rcItem.right;
		}
		if(pBottom) {
			*pBottom = controlInfo.rcItem.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails)
{
	ATLASSERT_POINTER(pHitTestDetails, HitTestConstants);
	if(!pHitTestDetails) {
		return E_POINTER;
	}

	if(IsWindow()) {
		POINT pt = {x, y};
		CRect rc;
		GetWindowRect(&rc);
		ScreenToClient(&rc);
		if(rc.PtInRect(pt)) {
			if(containedComboBox.IsWindow()) {
				containedComboBox.GetWindowRect(&rc);
				ScreenToClient(&rc);
				if(rc.PtInRect(pt)) {
					*pHitTestDetails = htComboBoxPortion;
				}
			}
		} else {
			*pHitTestDetails = static_cast<HitTestConstants>(0);
			if(pt.x < rc.left) {
				*pHitTestDetails = static_cast<HitTestConstants>(*pHitTestDetails | htToLeft);
			} else if(pt.x >= rc.right) {
				*pHitTestDetails = static_cast<HitTestConstants>(*pHitTestDetails | htToRight);
			}
			if(pt.y < rc.top) {
				*pHitTestDetails = static_cast<HitTestConstants>(*pHitTestDetails | htAbove);
			} else if(pt.y >= rc.bottom) {
				*pHitTestDetails = static_cast<HitTestConstants>(*pHitTestDetails | htBelow);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::ListHitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IDriveComboBoxItem** ppHitItem)
{
	ATLASSERT_POINTER(pHitTestDetails, HitTestConstants);
	ATLASSERT_POINTER(ppHitItem, IDriveComboBoxItem*);
	if(!pHitTestDetails || !ppHitItem) {
		return E_POINTER;
	}

	if(containedListBox.IsWindow()) {
		ClassFactory::InitDriveComboItem(ListBoxHitTest(x, y, pHitTestDetails), this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppHitItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::OLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, IDriveComboBoxItemContainer* pDraggedItems/* = NULL*/, LONG itemCountToDisplay/* = -1*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
{
	if(supportedEffects == odeNone) {
		// don't waste time
		return S_OK;
	}
	if(hWndToAskForDragImage == -1) {
		ATLASSUME(pDraggedItems);
		if(!pDraggedItems) {
			return E_INVALIDARG;
		}
	}
	ATLASSERT_POINTER(pDataObject, LONG);
	LONG dummy = NULL;
	if(!pDataObject) {
		pDataObject = &dummy;
	}
	ATLASSERT_POINTER(pPerformedEffects, OLEDropEffectConstants);
	OLEDropEffectConstants performedEffect = odeNone;
	if(!pPerformedEffects) {
		pPerformedEffects = &performedEffect;
	}

	HWND hWnd = NULL;
	if(hWndToAskForDragImage == -1) {
		hWnd = *this;
	} else {
		hWnd = static_cast<HWND>(LongToHandle(hWndToAskForDragImage));
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IDataObject> pDataObjectToUse = NULL;
	if(*pDataObject) {
		pDataObjectToUse = *reinterpret_cast<LPDATAOBJECT*>(pDataObject);

		CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->Attach(pDataObjectToUse);
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->Release();
	} else {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->SetOwner(this);
		if(itemCountToDisplay == -1) {
			if(pDraggedItems) {
				if(flags.usingThemes && RunTimeHelper::IsVista()) {
					pDraggedItems->Count(&pOLEDataObjectObj->properties.numberOfItemsToDisplay);
				}
			}
		} else if(itemCountToDisplay >= 0) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				pOLEDataObjectObj->properties.numberOfItemsToDisplay = itemCountToDisplay;
			}
		}
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&pDataObjectToUse));
		pOLEDataObjectObj->Release();
	}
	ATLASSUME(pDataObjectToUse);
	pDataObjectToUse->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
	CComQIPtr<IDropSource, &IID_IDropSource> pDragSource(this);

	if(pDraggedItems) {
		pDraggedItems->Clone(&dragDropStatus.pDraggedItems);
	}
	POINT mousePosition = {0};
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);

	if(properties.supportOLEDragImages) {
		IDragSourceHelper* pDragSourceHelper = NULL;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDragSourceHelper));
		if(pDragSourceHelper) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				IDragSourceHelper2* pDragSourceHelper2 = NULL;
				pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&pDragSourceHelper2));
				if(pDragSourceHelper2) {
					pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
					// this was the only place we actually use IDragSourceHelper2
					pDragSourceHelper->Release();
					pDragSourceHelper = static_cast<IDragSourceHelper*>(pDragSourceHelper2);
				}
			}

			HRESULT hr = pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
			if(FAILED(hr)) {
				/* This happens if full window dragging is deactivated. Actually, InitializeFromWindow() contains a
				   fallback mechanism for this case. This mechanism retrieves the passed window's class name and
				   builds the drag image using TVM_CREATEDRAGIMAGE if it's SysTreeView32, LVM_CREATEDRAGIMAGE if
				   it's SysListView32 and so on. Our class name is ExplorerListView[U|A], so we're doomed.
				   So how can we have drag images anyway? Well, we use a very ugly hack: We temporarily activate
				   full window dragging. */
				BOOL fullWindowDragging;
				SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &fullWindowDragging, 0);
				if(!fullWindowDragging) {
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, TRUE, NULL, 0);
					pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FALSE, NULL, 0);
				}
			}

			if(pDragSourceHelper) {
				pDragSourceHelper->Release();
			}
		}
	}

	if(IsLeftMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_LBUTTON;
	} else if(IsRightMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_RBUTTON;
	}
	if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
		dragDropStatus.useItemCountLabelHack = TRUE;
	}

	if(pOLEDataObject) {
		Raise_OLEStartDrag(pOLEDataObject);
	}
	HRESULT hr = DoDragDrop(pDataObjectToUse, pDragSource, supportedEffects, reinterpret_cast<LPDWORD>(pPerformedEffects));
	KillTimer(timers.ID_DRAGSCROLL);
	if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
		Raise_OLECompleteDrag(pOLEDataObject, *pPerformedEffects);
	}

	dragDropStatus.Reset();
	return S_OK;
}

STDMETHODIMP DriveComboBox::OpenDropDownWindow(void)
{
	if(IsWindow()) {
		SendMessage(CB_SHOWDROPDOWN, TRUE, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DriveComboBox::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		UpdateWindow();
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::ReloadDrives(void)
{
	if(IsWindow()) {
		SendMessage(CB_RESETCONTENT, 0, 0);

		BOOL displayIcons = (SendMessage(CBEM_GETIMAGELIST, 0, 0) != NULL && properties.useSystemImageList != usilNone);
		SHFILEINFO fileInfo = {0};

		CComPtr<IShellFolder> pMyComputerISF;
		if(properties.displayOverlayImages || properties.displayNameStyle == dnsShellStyle) {
			PIDLIST_ABSOLUTE pIDLMyComputer = GetMyComputerPIDL();
			if(pIDLMyComputer) {
				ATLVERIFY(SUCCEEDED(BindToPIDL(pIDLMyComputer, IID_PPV_ARGS(&pMyComputerISF))));
				ILFree(pIDLMyComputer);
			}
		}
		PIDLIST_ABSOLUTE pIDLDrive = NULL;

		BOOL displayDriveType[7];
		if(properties.displayedDriveTypes != dtAll) {
			displayDriveType[DRIVE_UNKNOWN] = (properties.displayedDriveTypes & dtUnknown);
			displayDriveType[DRIVE_NO_ROOT_DIR] = (properties.displayedDriveTypes & dtOther);
			displayDriveType[DRIVE_REMOVABLE] = (properties.displayedDriveTypes & dtRemovable);
			displayDriveType[DRIVE_FIXED] = (properties.displayedDriveTypes & dtFixed);
			displayDriveType[DRIVE_REMOTE] = (properties.displayedDriveTypes & dtRemote);
			displayDriveType[DRIVE_CDROM] = (properties.displayedDriveTypes & dtOptical);
			displayDriveType[DRIVE_RAMDISK] = (properties.displayedDriveTypes & dtRAMDisk);
		} else {
			for(int i = 0; i < 7; i++) {
				displayDriveType[i] = TRUE;
			}
		}
		BOOL retrieveVolumeName[7];
		if(properties.driveTypesWithVolumeName != dtAll) {
			retrieveVolumeName[DRIVE_UNKNOWN] = (properties.driveTypesWithVolumeName & dtUnknown);
			retrieveVolumeName[DRIVE_NO_ROOT_DIR] = (properties.driveTypesWithVolumeName & dtOther);
			retrieveVolumeName[DRIVE_REMOVABLE] = (properties.driveTypesWithVolumeName & dtRemovable);
			retrieveVolumeName[DRIVE_FIXED] = (properties.driveTypesWithVolumeName & dtFixed);
			retrieveVolumeName[DRIVE_REMOTE] = (properties.driveTypesWithVolumeName & dtRemote);
			retrieveVolumeName[DRIVE_CDROM] = (properties.driveTypesWithVolumeName & dtOptical);
			retrieveVolumeName[DRIVE_RAMDISK] = (properties.driveTypesWithVolumeName & dtRAMDisk);
		} else {
			for(int i = 0; i < 7; i++) {
				retrieveVolumeName[i] = TRUE;
			}
		}

		DWORD availableDrives = GetLogicalDrives();
		TCHAR pDrive[4] = TEXT("A:\\");
		COMBOBOXEXITEM insertionData = {0};
		insertionData.iItem = -1;
		TCHAR pDisplayName[2 * MAX_PATH];
		TCHAR pVolumeName[MAX_PATH + 1];
		for(int i = 0; i < 32; ++i) {
			BOOL displayDrive = (availableDrives & (1 << i));
			if(!displayDrive) {
				// GetLogicalDrives omits remote drives that are currently unavailable
				PIDLIST_ABSOLUTE pIDL = ILCreateFromPath(pDrive);
				if(pIDL) {
					displayDrive = TRUE;
					ILFree(pIDL);
				}
			}
			UINT driveType = UINT_MAX;
			if(displayDrive && properties.displayedDriveTypes != dtAll) {
				driveType = GetDriveType(pDrive);
				displayDrive = displayDrive && displayDriveType[driveType];
			}

			if(displayDrive) {
				insertionData.mask = CBEIF_LPARAM | CBEIF_TEXT | CBEIF_NOINTERCEPTION;
				insertionData.lParam = i + 1;

				ZeroMemory(pVolumeName, (MAX_PATH + 1) * sizeof(TCHAR));
				switch(properties.displayNameStyle) {
					case dnsDriveLetter:
						lstrcpyn(pDisplayName, pDrive, 2);
						break;
					case dnsDriveLetterWithColon:
						lstrcpyn(pDisplayName, pDrive, 3);
						break;
					case dnsVolumeName:
						if(driveType == UINT_MAX) {
							driveType = GetDriveType(pDrive);
						}
						if(retrieveVolumeName[driveType]) {
							GetVolumeInformation(pDrive, pVolumeName, MAX_PATH + 1, NULL, NULL, NULL, NULL, 0);
						}
						lstrcpyn(pDisplayName, pVolumeName, MAX_PATH + 1);
						break;
					case dnsVB6Style:
						lstrcpyn(pDisplayName, pDrive, 3);
						if(driveType == UINT_MAX) {
							driveType = GetDriveType(pDrive);
						}
						if(retrieveVolumeName[driveType]) {
							GetVolumeInformation(pDrive, pVolumeName, MAX_PATH + 1, NULL, NULL, NULL, NULL, 0);
						}
						if(lstrlen(pVolumeName) > 0) {
							StringCchCat(pDisplayName, 2 * MAX_PATH, TEXT(" ["));
							StringCchCat(pDisplayName, 2 * MAX_PATH, pVolumeName);
							StringCchCat(pDisplayName, 2 * MAX_PATH, TEXT("]"));
						}
						break;
					case dnsShellStyle:
						if(pMyComputerISF) {
							pIDLDrive = ILCreateFromPath(pDrive);
							if(pIDLDrive) {
								STRRET displayName;
								if(SUCCEEDED(pMyComputerISF->GetDisplayNameOf(ILFindLastID(pIDLDrive), SHGDN_NORMAL, &displayName))) {
									StrRetToBuf(&displayName, ILFindLastID(pIDLDrive), pDisplayName, 2 * MAX_PATH);
								}
							}
						}
						break;
				}
				insertionData.pszText = pDisplayName;

				if(displayIcons) {
					SHGetFileInfo(pDrive, FILE_ATTRIBUTE_NORMAL, &fileInfo, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHIL_SMALL);
					insertionData.iImage = fileInfo.iIcon;
					insertionData.mask |= CBEIF_IMAGE | CBEIF_SELECTEDIMAGE;
					insertionData.iSelectedImage = insertionData.iImage;
				}
				if(properties.displayOverlayImages) {
					if(!pIDLDrive) {
						pIDLDrive = ILCreateFromPath(pDrive);
					}
					if(pIDLDrive && pMyComputerISF) {
						CComQIPtr<IShellIconOverlay> pParentISIO = pMyComputerISF;
						insertionData.iOverlay = 0;
						pParentISIO->GetOverlayIndex(ILFindLastID(pIDLDrive), &insertionData.iOverlay);
						if(insertionData.iOverlay > 0) {
							insertionData.mask |= CBEIF_OVERLAY;
						}
					}
				}

				SendMessage(CBEM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));

				if(pIDLDrive) {
					ILFree(pIDLDrive);
					pIDLDrive = NULL;
				}
			}
			pDrive[0]++;
		}
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP DriveComboBox::SelectItemByItemData(LONG itemData, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IDriveComboBoxItem** ppFoundItem/* = NULL*/)
{
	// cannot use CB_SELECTSTRING, so use emulation
	HRESULT hr = FindItemByItemData(itemData, startAfterItem, ppFoundItem);
	if(SUCCEEDED(hr)) {
		hr = putref_SelectedItem(*ppFoundItem);
	}
	return hr;
}

STDMETHODIMP DriveComboBox::SelectItemByPath(BSTR path, IDriveComboBoxItem** ppFoundItem)
{
	HRESULT hr = FindItemByPath(path, ppFoundItem);
	if(SUCCEEDED(hr)) {
		hr = putref_SelectedItem(*ppFoundItem);
	}
	return hr;
}

STDMETHODIMP DriveComboBox::SelectItemByText(BSTR searchString, VARIANT_BOOL exactMatch/* = VARIANT_TRUE*/, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IDriveComboBoxItem** ppFoundItem/* = NULL*/)
{
	// cannot use CB_SELECTSTRING, so use emulation
	HRESULT hr = FindItemByText(searchString, exactMatch, startAfterItem, ppFoundItem);
	if(SUCCEEDED(hr)) {
		hr = putref_SelectedItem(*ppFoundItem);
	}
	return hr;
}


LRESULT DriveComboBox::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(*this);
	}
	return lr;
}

LRESULT DriveComboBox::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(*this);

	wasHandled = FALSE;
	return 0;
}

LRESULT DriveComboBox::OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if(ime != imeNoControl) {
		HWND h = GetFocus();
		if(h == *this || h == containedComboBox || h == containedListBox) {
			// we've the focus, so configure the IME
			SetCurrentIMEContextMode(h, ime);
		}
	}
	return lr;
}

LRESULT DriveComboBox::OnMouseWheel(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		switch(index) {
			case 1:
				Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam), htComboBoxPortion);
				break;
		}
	} else {
		switch(index) {
			case 1:
				if(!mouseStatus_ComboBox.enteredControl) {
					mouseStatus_ComboBox.EnterControl();
				}
				break;
		}
	}

	return 0;
}

LRESULT DriveComboBox::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT DriveComboBox::OnSetCursor(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	switch(index) {
		case 0:
			GetClientRect(&clientArea);
			ClientToScreen(&clientArea);
			break;
		case 1:
			containedComboBox.GetClientRect(&clientArea);
			containedComboBox.ClientToScreen(&clientArea);
			break;
	}
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		switch(index) {
			case 0:
				setCursor = (WindowFromPoint(mousePosition) == *this);
				break;
			case 1:
				setCursor = (WindowFromPoint(mousePosition) == containedComboBox);
				break;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT DriveComboBox::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<DriveComboBox>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<DriveComboBox>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT DriveComboBox::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_DCB_FONT) == S_FALSE) {
		return 0;
	}

	if(!IsInDesignMode()) {
		if(properties.setSelectionFieldHeight != -1) {
			properties.selectionFieldHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(CB_GETITEMHEIGHT, static_cast<WPARAM>(-1), 0));
		}
		if(properties.setItemHeight != -1) {
			properties.itemHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(CB_GETITEMHEIGHT, 0, 0));
		}
	}
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(!IsInDesignMode()) {
		if(properties.setItemHeight != -1) {
			SendMessage(CB_SETITEMHEIGHT, 0, properties.itemHeight);
		}
		if(properties.setSelectionFieldHeight != -1) {
			SendMessage(CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), properties.selectionFieldHeight);
		}
	}

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_DCB_FONT);
	}

	return lr;
}

LRESULT DriveComboBox::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT DriveComboBox::OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_DCB_TEXT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_TEXT);
	SendOnDataChange();
	return lr;
}

LRESULT DriveComboBox::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT DriveComboBox::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT DriveComboBox::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		case timers.ID_DRAGSCROLL:
			ListBoxAutoScroll();
			break;

		case timers.ID_DRAGDROPDOWN:
			KillTimer(timers.ID_DRAGDROPDOWN);
			dragDropStatus.dropDownTimerIsRunning = FALSE;
			OpenDropDownWindow();
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT DriveComboBox::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT DriveComboBox::OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL succeeded = FALSE;
	BOOL useVistaDragImage = FALSE;
	if(dragDropStatus.pDraggedItems) {
		if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
			succeeded = CreateVistaOLEDragImage(dragDropStatus.pDraggedItems, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
			useVistaDragImage = succeeded;
		}
		if(!succeeded) {
			// use a legacy drag image as fallback
			succeeded = CreateLegacyOLEDragImage(dragDropStatus.pDraggedItems, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
		}

		if(succeeded && RunTimeHelper::IsVista()) {
			FORMATETC format = {0};
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			STGMEDIUM medium = {0};
			medium.tymed = TYMED_HGLOBAL;
			medium.hGlobal = GlobalAlloc(GPTR, sizeof(BOOL));
			if(medium.hGlobal) {
				LPBOOL pUseVistaDragImage = static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				*pUseVistaDragImage = useVistaDragImage;
				GlobalUnlock(medium.hGlobal);

				dragDropStatus.pSourceDataObject->SetData(&format, &medium, TRUE);
			}
		}
	}

	wasHandled = succeeded;
	// TODO: Why do we have to return FALSE to have the set offset not ignored if a Vista drag image is used?
	return succeeded && !useVistaDragImage;
}

LRESULT DriveComboBox::OnDeleteItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	//TODO: toolTipStatus.InvalidateToolTipItem();

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IDriveComboBoxItem> pDCBoxItemItf = NULL;
	CComObject<DriveComboBoxItem>* pDCBoxItemObj = NULL;
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		CComObject<DriveComboBoxItem>::CreateInstance(&pDCBoxItemObj);
		pDCBoxItemObj->AddRef();
		pDCBoxItemObj->SetOwner(this);
		pDCBoxItemObj->Attach(static_cast<int>(wParam));
		pDCBoxItemObj->QueryInterface(IID_IDriveComboBoxItem, reinterpret_cast<LPVOID*>(&pDCBoxItemItf));
		pDCBoxItemObj->Release();
		Raise_RemovingItem(pDCBoxItemItf, &cancel);
	}

	LRESULT lr = CB_ERR;
	if(cancel == VARIANT_FALSE) {
		if(!(properties.disabledEvents & deMouseEvents)) {
			if(static_cast<int>(wParam) == itemUnderMouse) {
				// we're removing the item below the mouse cursor
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				containedListBox.ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
				ListBoxHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				Raise_ItemMouseLeave(pDCBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				itemUnderMouse = -1;
			}
		}

		if(!(properties.disabledEvents & deFreeItemData)) {
			Raise_FreeItemData(pDCBoxItemItf);
		}
		/*BOOL textChanged = FALSE;
		if(static_cast<int>(wParam) == currentSelectedItem) {
			if(!(properties.disabledEvents & deTextChangedEvents)) {
				CComBSTR selectedItemText;
				CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(currentSelectedItem, this);
				if(pDCBItem) {
					pDCBItem->get_Text(&selectedItemText);
					CComBSTR controlText;
					get_Text(&controlText);

					textChanged = (controlText == selectedItemText);
				}
			}
		}*/

		CComPtr<IVirtualDriveComboBoxItem> pVDCBoxItemItf = NULL;
		if(!(properties.disabledEvents & deItemDeletionEvents)) {
			CComObject<VirtualDriveComboBoxItem>* pVDCBoxItemObj = NULL;
			CComObject<VirtualDriveComboBoxItem>::CreateInstance(&pVDCBoxItemObj);
			pVDCBoxItemObj->AddRef();
			pVDCBoxItemObj->SetOwner(this);
			if(pDCBoxItemObj) {
				pDCBoxItemObj->SaveState(pVDCBoxItemObj);
				pVDCBoxItemObj->properties.driveIndex = GetDriveIndex(static_cast<int>(wParam));
			}

			pVDCBoxItemObj->QueryInterface(IID_IVirtualDriveComboBoxItem, reinterpret_cast<LPVOID*>(&pVDCBoxItemItf));
			pVDCBoxItemObj->Release();
		}
		lr = DefWindowProc(message, wParam, lParam);
		if(lr != CB_ERR) {
			LONG itemID = ItemIndexToID(static_cast<int>(wParam));
			RemoveItemFromItemContainers(itemID);
			#ifdef USE_STL
				if(itemID >= 0) {
					itemIDs.erase(itemIDs.begin() + static_cast<int>(wParam));
				}
				std::list<ItemData>::iterator iter2 = itemParams.begin();
				if(iter2 != itemParams.end()) {
					std::advance(iter2, static_cast<int>(wParam));
					if(iter2 != itemParams.end()) {
						itemParams.erase(iter2);
					}
				}
			#else
				if(itemID >= 0) {
					itemIDs.RemoveAt(static_cast<int>(wParam));
				}
				POSITION p = itemParams.FindIndex(static_cast<int>(wParam));
				if(p) {
					itemParams.RemoveAt(p);
				}
			#endif
			if(!(properties.disabledEvents & deItemDeletionEvents)) {
				Raise_RemovedItem(pVDCBoxItemItf);
			}
			/*if(static_cast<int>(wParam) == currentSelectedItem) {
				// NOTE: currentSelectedItem has already been deleted, so the event is not quite right
				Raise_SelectionChanged(currentSelectedItem, -1);
				if(textChanged && !(properties.disabledEvents & deTextChangedEvents)) {
					Raise_TextChanged();
				}
			}*/

			if(!(properties.disabledEvents & deMouseEvents)) {
				// maybe we have a new item below the mouse cursor now
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				containedListBox.ScreenToClient(&mousePosition);
				HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
				int newItemUnderMouse = ListBoxHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				if(newItemUnderMouse != itemUnderMouse) {
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					if(itemUnderMouse >= 0) {
						pDCBoxItemItf = ClassFactory::InitDriveComboItem(itemUnderMouse, this);
						Raise_ItemMouseLeave(pDCBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
					itemUnderMouse = newItemUnderMouse;
					if(itemUnderMouse >= 0) {
						pDCBoxItemItf = ClassFactory::InitDriveComboItem(itemUnderMouse, this);
						Raise_ItemMouseEnter(pDCBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}
	return lr;
}

LRESULT DriveComboBox::OnGetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL overwriteLParam = FALSE;

	PCOMBOBOXEXITEM pItemData = reinterpret_cast<PCOMBOBOXEXITEM>(lParam);
	#ifdef DEBUG
		if(message == CBEM_GETITEMA) {
			ATLASSERT_POINTER(pItemData, COMBOBOXEXITEMA);
		} else {
			ATLASSERT_POINTER(pItemData, COMBOBOXEXITEMW);
		}
	#endif
	if(!pItemData) {
		return DefWindowProc(message, wParam, lParam);
	}

	if(pItemData->mask & CBEIF_NOINTERCEPTION) {
		// ComboBoxEx32 shouldn't see this flag
		pItemData->mask &= ~CBEIF_NOINTERCEPTION;
	} else if(pItemData->mask == CBEIF_LPARAM) {
		// just look up in the 'itemParams' list
		#ifdef USE_STL
			std::list<ItemData>::iterator iter = itemParams.begin();
			if(iter != itemParams.end()) {
				std::advance(iter, pItemData->iItem);
				if(iter != itemParams.end()) {
					pItemData->lParam = iter->lParam;
					return TRUE;
				}
			}
		#else
			POSITION p = itemParams.FindIndex(pItemData->iItem);
			if(p) {
				pItemData->lParam = itemParams.GetAt(p).lParam;
				return TRUE;
			}
		#endif
		// item not found
		return FALSE;
	} else if(pItemData->mask & CBEIF_LPARAM) {
		overwriteLParam = TRUE;
	}

	// forward the message
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(overwriteLParam) {
		#ifdef USE_STL
			std::list<ItemData>::iterator iter = itemParams.begin();
			if(iter != itemParams.end()) {
				std::advance(iter, pItemData->iItem);
				if(iter != itemParams.end()) {
					pItemData->lParam = iter->lParam;
				}
			}
		#else
			POSITION p = itemParams.FindIndex(pItemData->iItem);
			if(p) {
				pItemData->lParam = itemParams.GetAt(p).lParam;
			}
		#endif
	}

	return lr;
}

LRESULT DriveComboBox::OnGetItemData(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	// just look up in the 'itemParams' list
	#ifdef USE_STL
		std::list<ItemData>::iterator iter = itemParams.begin();
		if(iter != itemParams.end()) {
			std::advance(iter, wParam);
			if(iter != itemParams.end()) {
				return iter->lParam;
			}
		}
	#else
		POSITION p = itemParams.FindIndex(wParam);
		if(p) {
			return itemParams.GetAt(p).lParam;
		}
	#endif
	// item not found
	return CB_ERR;
}

LRESULT DriveComboBox::OnInsertItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	//TODO: toolTipStatus.InvalidateToolTipItem();

	int insertedItem = CB_ERR;

	if(!(properties.disabledEvents & deItemInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;

		CComObject<VirtualDriveComboBoxItem>* pVDCBoxItemObj = NULL;
		CComPtr<IVirtualDriveComboBoxItem> pVDCBoxItemItf = NULL;
		CComObject<VirtualDriveComboBoxItem>::CreateInstance(&pVDCBoxItemObj);
		pVDCBoxItemObj->AddRef();
		pVDCBoxItemObj->SetOwner(this);

		#ifdef UNICODE
			BOOL mustConvert = (message == CBEM_INSERTITEMA);
		#else
			BOOL mustConvert = (message == CBEM_INSERTITEMW);
		#endif
		if(mustConvert) {
			#ifdef UNICODE
				PCOMBOBOXEXITEMA pItemData = reinterpret_cast<PCOMBOBOXEXITEMA>(lParam);
				COMBOBOXEXITEMW convertedItemData = {0};
				LPSTR p = NULL;
				if(pItemData->pszText != LPSTR_TEXTCALLBACKA) {
					p = pItemData->pszText;
				}
				CA2W converter(p);
				if(pItemData->pszText == LPSTR_TEXTCALLBACKA) {
					convertedItemData.pszText = LPSTR_TEXTCALLBACKW;
				} else {
					convertedItemData.pszText = converter;
				}
			#else
				PCOMBOBOXEXITEMW pItemData = reinterpret_cast<PCOMBOBOXEXITEMW>(lParam);
				COMBOBOXEXITEMA convertedItemData = {0};
				LPWSTR p = NULL;
				if(pItemData->pszText != LPSTR_TEXTCALLBACKW) {
					p = pItemData->pszText;
				}
				CW2A converter(p);
				if(pItemData->pszText == LPSTR_TEXTCALLBACKW) {
					convertedItemData.pszText = LPSTR_TEXTCALLBACKA;
				} else {
					convertedItemData.pszText = converter;
				}
			#endif
			convertedItemData.cchTextMax = pItemData->cchTextMax;
			convertedItemData.mask = pItemData->mask;
			convertedItemData.iItem = pItemData->iItem;
			convertedItemData.iImage = pItemData->iImage;
			convertedItemData.iSelectedImage = pItemData->iSelectedImage;
			convertedItemData.iOverlay = pItemData->iOverlay;
			convertedItemData.lParam = pItemData->lParam;
			convertedItemData.iIndent = pItemData->iIndent;
			pVDCBoxItemObj->LoadState(&convertedItemData);
		} else {
			pVDCBoxItemObj->LoadState(reinterpret_cast<PCOMBOBOXEXITEM>(lParam));
		}

		pVDCBoxItemObj->QueryInterface(IID_IVirtualDriveComboBoxItem, reinterpret_cast<LPVOID*>(&pVDCBoxItemItf));
		pVDCBoxItemObj->Release();

		Raise_InsertingItem(pVDCBoxItemItf, &cancel);
		pVDCBoxItemObj = NULL;

		if(cancel == VARIANT_FALSE) {
			// finally pass the message to the combo box
			PCOMBOBOXEXITEM pDetails = reinterpret_cast<PCOMBOBOXEXITEM>(lParam);
			ItemData itemData = {0};
			if(!(pDetails->mask & CBEIF_NOINTERCEPTION)) {
				if(pDetails->mask & CBEIF_LPARAM) {
					itemData.lParam = pDetails->lParam;
					pDetails->lParam = 0;
				}
			}
			pDetails->mask &= ~CBEIF_NOINTERCEPTION;

			LONG itemID = GetNewItemID();
			insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
			if(insertedItem >= 0) {
				#ifdef USE_STL
					if(insertedItem >= static_cast<int>(itemIDs.size())) {
						itemIDs.push_back(itemID);
					} else {
						itemIDs.insert(itemIDs.begin() + insertedItem, itemID);
					}
					std::list<ItemData>::iterator iter2 = itemParams.begin();
					if(iter2 != itemParams.end()) {
						std::advance(iter2, insertedItem);
					}
					itemParams.insert(iter2, itemData);
				#else
					if(insertedItem >= static_cast<int>(itemIDs.GetCount())) {
						itemIDs.Add(itemID);
					} else {
						itemIDs.InsertAt(insertedItem, itemID);
					}
					POSITION p = itemParams.FindIndex(insertedItem);
					if(p) {
						itemParams.InsertBefore(p, itemData);
					} else {
						itemParams.AddTail(itemData);
					}
				#endif
				CComPtr<IDriveComboBoxItem> pDCBoxItemItf = ClassFactory::InitDriveComboItem(insertedItem, this);
				if(pDCBoxItemItf) {
					Raise_InsertedItem(pDCBoxItemItf);
				}
			}
		}
	} else {
		// finally pass the message to the combo box
		PCOMBOBOXEXITEM pDetails = reinterpret_cast<PCOMBOBOXEXITEM>(lParam);
		ItemData itemData = {0};
		if(!(pDetails->mask & CBEIF_NOINTERCEPTION)) {
			if(pDetails->mask & CBEIF_LPARAM) {
				itemData.lParam = pDetails->lParam;
				pDetails->lParam = 0;
			}
		}
		pDetails->mask &= ~CBEIF_NOINTERCEPTION;

		LONG itemID = GetNewItemID();
		insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem >= 0) {
			#ifdef USE_STL
				if(insertedItem >= static_cast<int>(itemIDs.size())) {
					itemIDs.push_back(itemID);
				} else {
					itemIDs.insert(itemIDs.begin() + insertedItem, itemID);
				}
				std::list<ItemData>::iterator iter2 = itemParams.begin();
				if(iter2 != itemParams.end()) {
					std::advance(iter2, insertedItem);
				}
				itemParams.insert(iter2, itemData);
			#else
				if(insertedItem >= static_cast<int>(itemIDs.GetCount())) {
					itemIDs.Add(itemID);
				} else {
					itemIDs.InsertAt(insertedItem, itemID);
				}
				POSITION p = itemParams.FindIndex(insertedItem);
				if(p) {
					itemParams.InsertBefore(p, itemData);
				} else {
					itemParams.AddTail(itemData);
				}
			#endif
		}
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		if(insertedItem != CB_ERR) {
			// maybe we have a new item below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			containedListBox.ScreenToClient(&mousePosition);
			HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
			int newItemUnderMouse = ListBoxHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			if(newItemUnderMouse != itemUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(itemUnderMouse >= 0) {
					CComPtr<IDriveComboBoxItem> pDCBoxItem = ClassFactory::InitDriveComboItem(itemUnderMouse, this);
					Raise_ItemMouseLeave(pDCBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				if(itemUnderMouse >= 0) {
					CComPtr<IDriveComboBoxItem> pDCBoxItem = ClassFactory::InitDriveComboItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pDCBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return insertedItem;
}

LRESULT DriveComboBox::OnSetImageList(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_DCB_HIMAGELIST) == S_FALSE) {
		return 0;
	}

	if(!IsInDesignMode()) {
		if(properties.setSelectionFieldHeight != -1) {
			properties.selectionFieldHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(CB_GETITEMHEIGHT, static_cast<WPARAM>(-1), 0));
		}
		if(properties.setItemHeight != -1) {
			properties.itemHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(CB_GETITEMHEIGHT, 0, 0));
		}
	}
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(!IsInDesignMode()) {
		if(properties.setItemHeight != -1) {
			SendMessage(CB_SETITEMHEIGHT, 0, properties.itemHeight);
		}
		if(properties.setSelectionFieldHeight != -1) {
			SendMessage(CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), properties.selectionFieldHeight);
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_HIMAGELIST);
	return lr;
}

LRESULT DriveComboBox::OnSetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	PCOMBOBOXEXITEM pDetails = reinterpret_cast<PCOMBOBOXEXITEM>(lParam);
	#ifdef DEBUG
		if(message == CBEM_SETITEMA) {
			ATLASSERT_POINTER(pDetails, COMBOBOXEXITEMA);
		} else {
			ATLASSERT_POINTER(pDetails, COMBOBOXEXITEMW);
		}
	#endif
	if(!pDetails) {
		wasHandled = FALSE;
		return TRUE;
	}

	if(pDetails && pDetails->iItem == -1) {
		if(pDetails->mask & CBEIF_IMAGE) {
			properties.selectionFieldItemProperties.iImage = pDetails->iImage;
			properties.selectionFieldItemProperties.mask |= CBEIF_IMAGE;
		}
		if(pDetails->mask & CBEIF_SELECTEDIMAGE) {
			properties.selectionFieldItemProperties.iSelectedImage = pDetails->iSelectedImage;
			properties.selectionFieldItemProperties.mask |= CBEIF_SELECTEDIMAGE;
		}
		if(pDetails->mask & CBEIF_OVERLAY) {
			properties.selectionFieldItemProperties.iOverlay = pDetails->iOverlay;
			properties.selectionFieldItemProperties.mask |= CBEIF_OVERLAY;
		}

		UINT f = pDetails->mask & (CBEIF_IMAGE | CBEIF_SELECTEDIMAGE | CBEIF_OVERLAY);
		if(f) {
			// cannot set one without losing the other
			if(!(f & CBEIF_IMAGE)) {
				if(properties.selectionFieldItemProperties.mask & CBEIF_IMAGE) {
					pDetails->iImage = properties.selectionFieldItemProperties.iImage;
					pDetails->mask |= CBEIF_IMAGE;
				}
			}
			if(!(f & CBEIF_SELECTEDIMAGE)) {
				if(properties.selectionFieldItemProperties.mask & CBEIF_SELECTEDIMAGE) {
					pDetails->iSelectedImage = properties.selectionFieldItemProperties.iSelectedImage;
					pDetails->mask |= CBEIF_SELECTEDIMAGE;
				}
			}
			if(!(f & CBEIF_OVERLAY)) {
				if(properties.selectionFieldItemProperties.mask & CBEIF_OVERLAY) {
					pDetails->iOverlay = properties.selectionFieldItemProperties.iOverlay;
					pDetails->mask |= CBEIF_OVERLAY;
				}
			}
		}
	}

	if(pDetails->mask & CBEIF_NOINTERCEPTION) {
		// ComboBoxEx32 shouldn't see this flag
		pDetails->mask &= ~CBEIF_NOINTERCEPTION;
	} else if(pDetails->mask == CBEIF_LPARAM) {
		// simply update the 'itemParams' list
		#ifdef USE_STL
			std::list<ItemData>::iterator iter = itemParams.begin();
			if(iter != itemParams.end()) {
				std::advance(iter, pDetails->iItem);
				if(iter != itemParams.end()) {
					iter->lParam = pDetails->lParam;
					return TRUE;
				}
			}
		#else
			POSITION p = itemParams.FindIndex(pDetails->iItem);
			if(p) {
				itemParams.GetAt(p).lParam = pDetails->lParam;
				return TRUE;
			}
		#endif
		// item not found
		return FALSE;
	} else if(pDetails->mask & CBEIF_LPARAM) {
		#ifdef USE_STL
			std::list<ItemData>::iterator iter = itemParams.begin();
			if(iter != itemParams.end()) {
				std::advance(iter, pDetails->iItem);
				if(iter != itemParams.end()) {
					iter->lParam = pDetails->lParam;
					pDetails->mask &= ~CBEIF_LPARAM;
				}
			}
		#else
			POSITION p = itemParams.FindIndex(pDetails->iItem);
			if(p) {
				itemParams.GetAt(p).lParam = pDetails->lParam;
				pDetails->mask &= ~CBEIF_LPARAM;
			}
		#endif
	}

	// let ComboBoxEx32 handle the message
	return DefWindowProc(message, wParam, lParam);
}

LRESULT DriveComboBox::OnSetItemData(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	// simply update the 'itemParams' list
	#ifdef USE_STL
		std::list<ItemData>::iterator iter = itemParams.begin();
		if(iter != itemParams.end()) {
			std::advance(iter, wParam);
			if(iter != itemParams.end()) {
				iter->lParam = lParam;
				return 0;
			}
		}
	#else
		POSITION p = itemParams.FindIndex(wParam);
		if(p) {
			itemParams.GetAt(p).lParam = lParam;
			return 0;
		}
	#endif
	// item not found
	return CB_ERR;
}

LRESULT DriveComboBox::OnComboBoxDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_DestroyedComboBoxControlWindow(containedComboBox);
	return 0;
}

LRESULT DriveComboBox::OnComboBoxResetContent(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	//TODO: toolTipStatus.InvalidateToolTipItem();

	/* NOTE: If CB_DELETESTRING is called for the last item in the control, the control does *not* send
	         CB_RESETCONTENT to itself. */
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_RemovingItem(NULL, &cancel);
		if(cancel != VARIANT_FALSE) {
			return 0;
		}
	}
	if(!(properties.disabledEvents & deMouseEvents)) {
		if(itemUnderMouse != -1) {
			// we're removing the item below the mouse cursor
			CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(itemUnderMouse, this);
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			containedListBox.ScreenToClient(&mousePosition);
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
			ListBoxHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			Raise_ItemMouseLeave(pDCBItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
			itemUnderMouse = -1;
		}
	}
	if(!(properties.disabledEvents & deFreeItemData)) {
		Raise_FreeItemData(NULL);
	}
	/*BOOL textChanged = FALSE;
	if(currentSelectedItem != -1) {
		if(!(properties.disabledEvents & deTextChangedEvents)) {
			CComBSTR selectedItemText;
			CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(currentSelectedItem, this);
			if(pDCBItem) {
				pDCBItem->get_Text(&selectedItemText);
				CComBSTR controlText;
				get_Text(&controlText);

				textChanged = (controlText == selectedItemText);
			}
		}
	}*/
	LRESULT lr = containedComboBox.DefWindowProc(message, wParam, lParam);
	#ifdef USE_STL
		itemIDs.clear();
	#else
		itemIDs.RemoveAll();
	#endif
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		Raise_RemovedItem(NULL);
	}
	/*if(currentSelectedItem != -1) {
		// NOTE: currentSelectedItem has already been deleted, so the event is not quite right
		Raise_SelectionChanged(currentSelectedItem, -1);
		if(textChanged && !(properties.disabledEvents & deTextChangedEvents)) {
			Raise_TextChanged();
		}
	}*/

	RemoveItemFromItemContainers(-1);
	#ifdef USE_STL
		itemParams.clear();
	#else
		itemParams.RemoveAll();
	#endif
	return lr;
}

LRESULT DriveComboBox::OnComboBoxSetCueBanner(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_DCB_CUEBANNER) == S_FALSE) {
		return FALSE;
	}

	LRESULT lr = containedComboBox.DefWindowProc(message, wParam, lParam);
	if(lr) {
		properties.cueBanner = reinterpret_cast<LPWSTR>(lParam);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_CUEBANNER);
	SendOnDataChange();
	return lr;
}

LRESULT DriveComboBox::OnComboBoxSetCurSel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_DCB_SELECTEDITEM) == S_FALSE) {
		return FALSE;
	}

	LRESULT lr = containedComboBox.DefWindowProc(message, wParam, lParam);
	Raise_SelectionChanged(currentSelectedItem, wParam);

	SetDirty(TRUE);
	FireOnChanged(DISPID_DCB_SELECTEDITEM);
	SendOnDataChange();
	return lr;
}

LRESULT DriveComboBox::OnListBoxDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_DestroyedListBoxControlWindow(containedListBox);
	return 0;
}

LRESULT DriveComboBox::OnListBoxLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_ListMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	//if(properties.allowDragDrop) {
		// not useful
		/*HitTestConstants dummy = static_cast<HitTestConstants>(0);
		int itemIndex = ListBoxHitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
		if(itemIndex != -1) {
			dragDropStatus.candidate.itemIndex = itemIndex;
			dragDropStatus.candidate.button = MK_LBUTTON;
			dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
			dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);
			//containedListBox.SetCapture();
		}*/
	//}

	return 0;
}

LRESULT DriveComboBox::OnListBoxLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	//if(properties.allowDragDrop) {
		// not useful
		/*dragDropStatus.candidate.itemIndex = -1;*/
	//}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_ListMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DriveComboBox::OnListBoxMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_ListMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DriveComboBox::OnListBoxMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_ListMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DriveComboBox::OnListBoxMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	static int lastSelectedItem = -1;
	HitTestConstants dummy;
	int itemIndex = ListBoxHitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
	if(itemIndex != lastSelectedItem) {
		RECT rc = {0};
		if(containedListBox.SendMessage(LB_GETITEMRECT, lastSelectedItem, reinterpret_cast<LPARAM>(&rc)) != LB_ERR) {
			containedListBox.InvalidateRect(&rc);
		}
		if(containedListBox.SendMessage(LB_GETITEMRECT, itemIndex, reinterpret_cast<LPARAM>(&rc)) != LB_ERR) {
			containedListBox.InvalidateRect(&rc);
		}
		lastSelectedItem = itemIndex;
	}

	if(!(properties.disabledEvents & deListBoxMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_ListMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	//if(properties.allowDragDrop) {
		// not useful
		/*if(dragDropStatus.candidate.itemIndex != -1) {
			// calculate the rectangle, that the mouse cursor must leave to start dragging
			int clickRectWidth = GetSystemMetrics(SM_CXDRAG);
			if(clickRectWidth < 4) {
				clickRectWidth = 4;
			}
			int clickRectHeight = GetSystemMetrics(SM_CYDRAG);
			if(clickRectHeight < 4) {
				clickRectHeight = 4;
			}
			CRect rc(dragDropStatus.candidate.position.x - clickRectWidth, dragDropStatus.candidate.position.y - clickRectHeight, dragDropStatus.candidate.position.x + clickRectWidth, dragDropStatus.candidate.position.y + clickRectHeight);

			if(!rc.PtInRect(CPoint(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)))) {
				BOOL useUnicode = (GetParent().SendMessage(WM_NOTIFYFORMAT, reinterpret_cast<WPARAM>(m_hWnd), NF_QUERY) == NFR_UNICODE);
				// we don't pass a string, so only the struct size matters, not the format
				NMCBEDRAGBEGINW notificationDetails = {0};
				notificationDetails.hdr.hwndFrom = *this;
				notificationDetails.hdr.code = (useUnicode ? CBEN_DRAGBEGINW : CBEN_DRAGBEGINA);
				notificationDetails.iItemid = dragDropStatus.candidate.itemIndex;
				GetParent().SendMessage(WM_NOTIFY, 0, reinterpret_cast<LPARAM>(&notificationDetails));

				dragDropStatus.candidate.itemIndex = -1;
			}
		}*/
	//}

	return 0;
}

LRESULT DriveComboBox::OnListBoxMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deListBoxMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		Raise_ListMouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else {
		if(!mouseStatus_ListBox.enteredControl) {
			mouseStatus_ListBox.EnterControl();
		}
	}

	return 0;
}

LRESULT DriveComboBox::OnListBoxRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_ListMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	//if(properties.allowDragDrop) {
		// not useful
		/*HitTestConstants dummy = static_cast<HitTestConstants>(0);
		int itemIndex = ListBoxHitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
		if(itemIndex != -1) {
			dragDropStatus.candidate.itemIndex = itemIndex;
			dragDropStatus.candidate.button = MK_RBUTTON;
			dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
			dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);
		}*/
	//}

	return 0;
}

LRESULT DriveComboBox::OnListBoxRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	//if(properties.allowDragDrop) {
		// not useful
		/*dragDropStatus.candidate.itemIndex = -1;*/
	//}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_ListMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}


LRESULT DriveComboBox::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT DriveComboBox::OnContextMenu(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	switch(index) {
		case 1:
			Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion, &showDefaultMenu);
			break;
	}
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}

	return 0;
}

LRESULT DriveComboBox::OnKeyDown(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	switch(index) {
		case 1:
			return containedComboBox.DefWindowProc(message, wParam, lParam);
		default:
			return DefWindowProc(message, wParam, lParam);
	}
}

LRESULT DriveComboBox::OnKeyUp(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	switch(index) {
		case 1:
			return containedComboBox.DefWindowProc(message, wParam, lParam);
		default:
			return DefWindowProc(message, wParam, lParam);
	}
}

LRESULT DriveComboBox::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<DriveComboBox>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	return lr;
}

LRESULT DriveComboBox::OnLButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		switch(index) {
			case 1:
				Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
				break;
		}
	}

	return 0;
}

LRESULT DriveComboBox::OnLButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	switch(index) {
		case 1:
			Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	return 0;
}

LRESULT DriveComboBox::OnLButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	switch(index) {
		case 1:
			Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	return 0;
}

LRESULT DriveComboBox::OnMButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		switch(index) {
			case 1:
				Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
				break;
		}
	}

	return 0;
}

LRESULT DriveComboBox::OnMButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	switch(index) {
		case 1:
			Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	return 0;
}

LRESULT DriveComboBox::OnMButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	switch(index) {
		case 1:
			Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	return 0;
}

LRESULT DriveComboBox::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT DriveComboBox::OnMouseHover(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	switch(index) {
		case 1:
			Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	return 0;
}

LRESULT DriveComboBox::OnMouseLeave(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	switch(index) {
		case 1:
			Raise_MouseLeave(button, shift, mouseStatus_ComboBox.lastPosition.x, mouseStatus_ComboBox.lastPosition.y, htComboBoxPortion);
			break;
	}

	return 0;
}

LRESULT DriveComboBox::OnMouseMove(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		switch(index) {
			case 1:
				Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
				break;
		}
	} else {
		switch(index) {
			case 1:
				if(!mouseStatus_ComboBox.enteredControl) {
					mouseStatus_ComboBox.EnterControl();
				}
				break;
		}
	}

	return 0;
}

LRESULT DriveComboBox::OnRButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		switch(index) {
			case 1:
				Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
				break;
		}
	}

	return 0;
}

LRESULT DriveComboBox::OnRButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	switch(index) {
		case 1:
			Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	if(index == 1) {
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(containedComboBox, &controlInfo);
		POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(!PtInRect(&controlInfo.rcButton, pt)) {
			/* Install a mouse hook, because ComboBoxEx32 seems to eat WM_RBUTTONUP messages if the user clicked
			   onto the selection field (they're sent for the drop-down button). We could also use NM_RCLICK
			   instead of a hook (as we do for ExplorerTreeView), but NM_RCLICK isn't sent either.
			*/
			ATLASSERT(!mouseStatus_ComboBox.hHook);
			if(!mouseStatus_ComboBox.hHook) {
				HANDLE hMem = GlobalAlloc(GHND, sizeof(IMouseHookHandler*));
				ATLASSERT(hMem);
				if(hMem) {
					IMouseHookHandler** ppHandler = static_cast<IMouseHookHandler**>(GlobalLock(hMem));
					ATLASSERT_POINTER(ppHandler, IMouseHookHandler*);
					if(ppHandler) {
						*ppHandler = this;
						GlobalUnlock(hMem);

						ATLVERIFY(SetProp(containedComboBox, HOOKPROPIDENTIFIER, hMem));
						mouseStatus_ComboBox.hHook = SetWindowsHookEx(WH_MOUSE, MouseHookProc, NULL, GetCurrentThreadId());
					}
				}
				ATLASSERT(mouseStatus_ComboBox.hHook);
			}
		}
	}

	return 0;
}

LRESULT DriveComboBox::OnRButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	switch(index) {
		case 1:
			Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	return 0;
}

LRESULT DriveComboBox::OnXButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		switch(index) {
			case 1:
				Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
				break;
		}
	}

	return 0;
}

LRESULT DriveComboBox::OnXButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	switch(index) {
		case 1:
			Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	return 0;
}

LRESULT DriveComboBox::OnXButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	switch(index) {
		case 1:
			Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), htComboBoxPortion);
			break;
	}

	return 0;
}


LRESULT DriveComboBox::OnBeginEditNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	Raise_BeginSelectionChange();
	return 0;
}

LRESULT DriveComboBox::OnDragBeginNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMCBEDRAGBEGIN pDetails = reinterpret_cast<LPNMCBEDRAGBEGIN>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == CBEN_DRAGBEGINA) {
			ATLASSERT_POINTER(pDetails, NMCBEDRAGBEGINA);
		} else {
			ATLASSERT_POINTER(pDetails, NMCBEDRAGBEGINW);
		}
	#endif
	if(!pDetails) {
		return 0;
	}

	CComBSTR selectionFieldText = L"";
	if((GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
		CComPtr<IDriveComboBoxItem> pDCBoxItem;
		get_SelectionFieldItem(&pDCBoxItem);
		if(pDCBoxItem) {
			pDCBoxItem->get_Text(&selectionFieldText);
		}
	} else {
		if(pNotificationDetails->code == CBEN_DRAGBEGINA) {
			selectionFieldText = reinterpret_cast<LPNMCBEDRAGBEGINA>(pDetails)->szText;
		} else {
			selectionFieldText = reinterpret_cast<LPNMCBEDRAGBEGINW>(pDetails)->szText;
		}
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	CComPtr<IDriveComboBoxItem> pDCBoxItem;
	if(pDetails->iItemid == -1) {
		get_SelectionFieldItem(&pDCBoxItem);
	} else {
		pDCBoxItem = ClassFactory::InitDriveComboItem(pDetails->iItemid, this);
	}
	if(button & 2/*MouseButtonConstants.vbRightButton*/) {
		Raise_ItemBeginRDrag(pDCBoxItem, selectionFieldText, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
	} else {
		button |= 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_ItemBeginDrag(pDCBoxItem, selectionFieldText, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
	}

	if(properties.handleOLEDragDrop & hoddSourcePart) {
		// get the drive index of pDCBoxItem
		int itemIndex = pDetails->iItemid;
		if(pDetails->iItemid == -1 && (GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
			itemIndex = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
		// get the drive's pIDL
		PIDLIST_ABSOLUTE pIDLDrive = GetItemPIDL(itemIndex);
		ATLASSERT_POINTER(pIDLDrive, ITEMIDLIST_ABSOLUTE);
		if(pIDLDrive) {
			// get the Computer's IShellFolder
			PIDLIST_ABSOLUTE pIDLMyComputer = GetMyComputerPIDL();
			ATLASSERT_POINTER(pIDLMyComputer, ITEMIDLIST_ABSOLUTE);
			if(pIDLMyComputer) {
				CComPtr<IShellFolder> pMyComputerISF;
				ATLVERIFY(SUCCEEDED(BindToPIDL(pIDLMyComputer, IID_PPV_ARGS(&pMyComputerISF))));
				ATLASSUME(pMyComputerISF);
				if(pMyComputerISF) {
					// get the IDataObject
					PUITEMID_CHILD pRelativePIDLs[1] = {ILFindLastID(pIDLDrive)};
					LPDATAOBJECT pDataObject = NULL;
					HRESULT hr = pMyComputerISF->GetUIObjectOf(GethWndShellUIParentWindow(), 1, pRelativePIDLs, IID_IDataObject, NULL, reinterpret_cast<LPVOID*>(&pDataObject));
					if(SUCCEEDED(hr)) {
						ATLASSUME(pDataObject);
						// retrieve the supported verbs
						SFGAOF attributes = SFGAO_CANCOPY | SFGAO_CANLINK | SFGAO_CANMOVE;
						hr = pMyComputerISF->GetAttributesOf(1, pRelativePIDLs, &attributes);
						if(SUCCEEDED(hr)) {
							int supportedEffects = attributes & (SFGAO_CANCOPY | SFGAO_CANLINK | SFGAO_CANMOVE);
							if(supportedEffects != 0/*odeNone*/) {
								// create the item container
								LONG itemID = ItemIndexToID(itemIndex);
								CComPtr<IDriveComboBoxItemContainer> pDraggedItems;
								hr = CreateItemContainer(_variant_t(itemID), &pDraggedItems);
								if(SUCCEEDED(hr)) {
									ATLASSUME(pDraggedItems);

									// TODO: What shall we do with <performedEffects>?
									DWORD performedEffects = 0/*odeNone*/;

									if(properties.supportOLEDragImages) {
										// use SHDoDragDrop
										pDraggedItems->Clone(&dragDropStatus.pDraggedItems);
										CComPtr<IOLEDataObject> pOLEDataObject = NULL;
										CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
										CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
										pOLEDataObjectObj->AddRef();
										pOLEDataObjectObj->Attach(pDataObject);
										pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
										pOLEDataObjectObj->Release();
										pDataObject->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
										CComPtr<IDropSource> pDragSource = NULL;
										QueryInterface(IID_PPV_ARGS(&pDragSource));

										if(button & 2/*MouseButtonConstants.vbRightButton*/) {
											dragDropStatus.draggingMouseButton = MK_RBUTTON;
										} else {
											dragDropStatus.draggingMouseButton = MK_LBUTTON;
										}

										if(pOLEDataObject) {
											Raise_OLEStartDrag(pOLEDataObject);
										}
										hr = SHDoDragDrop((RunTimeHelper::IsVista() ? NULL : static_cast<HWND>(*this)), pDataObject, pDragSource, supportedEffects, &performedEffects);
										if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
											Raise_OLECompleteDrag(pOLEDataObject, static_cast<OLEDropEffectConstants>(performedEffects));
										}

										dragDropStatus.Reset();
									} else {
										// SHDoDragDrop always is with drag images
										OLEDrag(reinterpret_cast<LONG*>(&pDataObject), static_cast<OLEDropEffectConstants>(supportedEffects), static_cast<OLE_HANDLE>(-1), pDraggedItems, -1, reinterpret_cast<OLEDropEffectConstants*>(&performedEffects));
									}
								}
							}
						}
					}
				}
				ILFree(pIDLMyComputer);
			}
			ILFree(pIDLDrive);
		}
	}
	return 0;
}

LRESULT DriveComboBox::OnEndEditNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(properties.disabledEvents & deSelectionChangingEvents) {
		return FALSE;
	}

	PNMCBEENDEDIT pDetails = reinterpret_cast<PNMCBEENDEDIT>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == CBEN_ENDEDITA) {
			ATLASSERT_POINTER(pDetails, NMCBEENDEDITA);
		} else {
			ATLASSERT_POINTER(pDetails, NMCBEENDEDITW);
		}
	#endif
	if(!pDetails) {
		return FALSE;
	}

	CComBSTR selectionFieldText = L"";
	if((GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_DROPDOWNLIST) {
		CComPtr<IDriveComboBoxItem> pDCBoxItem;
		get_SelectedItem(&pDCBoxItem);
		if(pDCBoxItem) {
			pDCBoxItem->get_Text(&selectionFieldText);
		}
	} else {
		if(pNotificationDetails->code == CBEN_ENDEDITA) {
			selectionFieldText = reinterpret_cast<PNMCBEENDEDITA>(pDetails)->szText;
		} else {
			selectionFieldText = reinterpret_cast<PNMCBEENDEDITW>(pDetails)->szText;
		}
	}
	SelectionChangeReasonConstants reason;
	if(pNotificationDetails->code == CBEN_ENDEDITA) {
		reason = static_cast<SelectionChangeReasonConstants>(reinterpret_cast<PNMCBEENDEDITA>(pDetails)->iWhy);
	} else {
		reason = static_cast<SelectionChangeReasonConstants>(reinterpret_cast<PNMCBEENDEDITW>(pDetails)->iWhy);
	}

	VARIANT_BOOL cancel = VARIANT_FALSE;
	Raise_SelectionChanging(pDetails->iNewSelection, selectionFieldText, BOOL2VARIANTBOOL(pDetails->fChanged), reason, &cancel);
	return VARIANTBOOL2BOOL(cancel);
}

LRESULT DriveComboBox::OnGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	PNMCOMBOBOXEX pDetails = reinterpret_cast<PNMCOMBOBOXEX>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == CBEN_GETDISPINFOA) {
			ATLASSERT_POINTER(pDetails, NMCOMBOBOXEXA);
		} else {
			ATLASSERT_POINTER(pDetails, NMCOMBOBOXEXW);
		}
	#endif

	CComPtr<IDriveComboBoxItem> pDCBoxItem = ClassFactory::InitDriveComboItem(pDetails->ceItem.iItem, this);
	LONG iconIndex = 0;
	LONG selectedIconIndex = 0;
	LONG overlayIndex = 0;
	LONG indent = 0;
	CComBSTR itemText = L"";
	LONG maxItemTextLength = 0;

	RequestedInfoConstants requestedInfo = static_cast<RequestedInfoConstants>(0);
	if(pDetails->ceItem.mask & CBEIF_IMAGE) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riIconIndex);
	}
	if(pDetails->ceItem.mask & CBEIF_SELECTEDIMAGE) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riSelectedIconIndex);
	}
	if(pDetails->ceItem.mask & CBEIF_OVERLAY) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riOverlayIndex);
	}
	if(pDetails->ceItem.mask & CBEIF_INDENT) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riIndent);
	}
	if(pDetails->ceItem.mask & CBEIF_TEXT) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riItemText);
		// exclude the terminating 0
		maxItemTextLength = pDetails->ceItem.cchTextMax - 1;
	}

	VARIANT_BOOL dontAskAgain = VARIANT_FALSE;
	Raise_ItemGetDisplayInfo(pDCBoxItem, requestedInfo, &iconIndex, &selectedIconIndex, &overlayIndex, &indent, maxItemTextLength, &itemText, &dontAskAgain);

	if(pDetails->ceItem.mask & CBEIF_IMAGE) {
		pDetails->ceItem.iImage = iconIndex;
	}
	if(pDetails->ceItem.mask & CBEIF_SELECTEDIMAGE) {
		pDetails->ceItem.iSelectedImage = selectedIconIndex;
	}
	if(pDetails->ceItem.mask & CBEIF_OVERLAY) {
		pDetails->ceItem.iOverlay = overlayIndex;
	}
	if(pDetails->ceItem.mask & CBEIF_INDENT) {
		pDetails->ceItem.iIndent = indent;
	}
	if(pDetails->ceItem.mask & CBEIF_TEXT) {
		// don't forget the terminating 0
		int bufferSize = itemText.Length() + 1;
		if(bufferSize > pDetails->ceItem.cchTextMax) {
			bufferSize = pDetails->ceItem.cchTextMax;
		}
		if(pNotificationDetails->code == CBEN_GETDISPINFOA) {
			if(bufferSize > 1) {
				lstrcpynA(reinterpret_cast<PNMCOMBOBOXEXA>(pDetails)->ceItem.pszText, CW2A(itemText), bufferSize);
			} else {
				reinterpret_cast<PNMCOMBOBOXEXA>(pDetails)->ceItem.pszText[0] = '\0';
			}
		} else {
			if(bufferSize > 1) {
				lstrcpynW(reinterpret_cast<PNMCOMBOBOXEXW>(pDetails)->ceItem.pszText, itemText, bufferSize);
			} else {
				reinterpret_cast<PNMCOMBOBOXEXW>(pDetails)->ceItem.pszText[0] = L'\0';
			}
		}
	}

	if(dontAskAgain == VARIANT_FALSE) {
		pDetails->ceItem.mask &= ~CBEIF_DI_SETITEM;
	} else {
		pDetails->ceItem.mask |= CBEIF_DI_SETITEM;
	}

	return 0;
}


LRESULT DriveComboBox::OnCloseUp(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_ListCloseUp();

	// let ComboBoxEx32 handle this notification
	wasHandled = FALSE;
	return 0;
}

LRESULT DriveComboBox::OnDropDown(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_ListDropDown();

	// NOTE: Due to ComboBoxEx32's drag'n'drop support we receive CBN_DROPDOWN on mouse-up, not on mouse-down.
	//if(GetAsyncKeyState(VK_LBUTTON) & 0x8000) {
		// HACK: We capture mouse input on mouse-down and this interferes with drop-down handling of ComboBox.
		if(mouseStatus_ComboBox.IsClickCandidate(1/*MouseButtonConstants.vbLeftButton*/)) {
			if(GetCapture() == containedComboBox) {
				ReleaseCapture();
			}
			mouseStatus_ComboBox.RemoveClickCandidate(1/*MouseButtonConstants.vbLeftButton*/);
		}
	//}

	// let ComboBoxEx32 handle this notification
	wasHandled = FALSE;
	return 0;
}

LRESULT DriveComboBox::OnErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_OutOfMemory();

	// let ComboBoxEx32 handle this notification
	wasHandled = FALSE;
	return 0;
}

LRESULT DriveComboBox::OnSelChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	int newSelectedItem = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
	if(newSelectedItem != currentSelectedItem) {
		Raise_SelectionChanged(currentSelectedItem, newSelectedItem);
	}

	// let ComboBoxEx32 handle this notification
	wasHandled = FALSE;
	return 0;
}

LRESULT DriveComboBox::OnSelEndCancel(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deSelectionChangingEvents)) {
		Raise_SelectionCanceled();
	}
	return 0;
}

LRESULT DriveComboBox::OnSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		HWND h = GetFocus();
		IMEModeConstants ime = GetCurrentIMEContextMode(h);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(h, ime);
			}
		}
	}

	// let ComboBoxEx32 handle this notification
	wasHandled = FALSE;
	return 0;
}


void DriveComboBox::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
			properties.font.owningFontResource = FALSE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT DriveComboBox::Raise_BeginSelectionChange(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeginSelectionChange();
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu)
{
	//if(m_nFreezeEvents == 0) {
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// propose the middle of the control's client rectangle as the menu's position
				CRect clientRectangle;
				GetClientRect(&clientRectangle);
				CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
			} else {
				return S_OK;
			}
		}

		return Fire_ContextMenu(button, shift, x, y, hitTestDetails, pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_CreatedComboBoxControlWindow(HWND hWndComboBox)
{
	// subclass the combo box control
	containedComboBox.SubclassWindow(hWndComboBox);

	// configure the combo box control
	if(containedComboBox.IsWindow()) {
		// NOTE: CB_SETCUEBANNER isn't forwarded by ComboBoxEx32
		if(RunTimeHelper::IsCommCtrl6()) {
			containedComboBox.SendMessage(CB_SETCUEBANNER, 0, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
		}

		COMBOBOXINFO comboBoxInfo = {0};
		comboBoxInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(containedComboBox, &comboBoxInfo);
		if(comboBoxInfo.hwndList) {
			Raise_CreatedListBoxControlWindow(comboBoxInfo.hwndList);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedComboBoxControlWindow(HandleToLong(hWndComboBox));
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_CreatedListBoxControlWindow(HWND hWndListBox)
{
	// subclass the list box control
	containedListBox.SubclassWindow(hWndListBox);
	if(properties.registerForOLEDragDrop) {
		if(containedListBox.IsWindow()) {
			ATLVERIFY(RegisterDragDrop(containedListBox, static_cast<IDropTarget*>(this)) == S_OK);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedListBoxControlWindow(HandleToLong(hWndListBox));
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_DestroyedComboBoxControlWindow(HWND hWndComboBox)
{
	if(!(properties.disabledEvents & deMouseEvents)) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = containedComboBox;
		trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		if(mouseStatus_ComboBox.enteredControl) {
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			Raise_MouseLeave(button, shift, mouseStatus_ComboBox.lastPosition.x, mouseStatus_ComboBox.lastPosition.y, htComboBoxPortion);
		}
	}
	if(mouseStatus_ComboBox.hHook) {
		UnhookWindowsHookEx(mouseStatus_ComboBox.hHook);
		mouseStatus_ComboBox.hHook = NULL;
	}
	HANDLE hMem = RemoveProp(hWndComboBox, HOOKPROPIDENTIFIER);
	if(hMem) {
		GlobalFree(hMem);
	}

	//containedComboBox.UnsubclassWindow();     // done automatically
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedComboBoxControlWindow(HandleToLong(hWndComboBox));
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_DestroyedControlWindow(HWND hWnd)
{
	KillTimer(timers.ID_REDRAW);
	RemoveItemFromItemContainers(-1);
	#ifdef USE_STL
		itemIDs.clear();
		itemParams.clear();
	#else
		itemIDs.RemoveAll();
		itemParams.RemoveAll();
	#endif
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_DestroyedListBoxControlWindow(HWND hWndListBox)
{
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(containedListBox) == S_OK);
	}

	//containedListBox.UnsubclassWindow();     // done automatically
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedListBoxControlWindow(HandleToLong(hWndListBox));
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_FreeItemData(IDriveComboBoxItem* pComboItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeItemData(pComboItem);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_InsertedItem(IDriveComboBoxItem* pComboItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedItem(pComboItem);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_InsertingItem(IVirtualDriveComboBoxItem* pComboItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingItem(pComboItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ItemBeginDrag(IDriveComboBoxItem* pComboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemBeginDrag(pComboItem, selectionFieldText, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ItemBeginRDrag(IDriveComboBoxItem* pComboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemBeginRDrag(pComboItem, selectionFieldText, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ItemGetDisplayInfo(IDriveComboBoxItem* pComboItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pSelectedIconIndex, LONG* pOverlayIndex, LONG* pIndent, LONG maxItemTextLength, BSTR* pItemText, VARIANT_BOOL* pDontAskAgain)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetDisplayInfo(pComboItem, requestedInfo, pIconIndex, pSelectedIconIndex, pOverlayIndex, pIndent, maxItemTextLength, pItemText, pDontAskAgain);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ItemMouseEnter(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemMouseEnter(pComboItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ItemMouseLeave(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemMouseLeave(pComboItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ListClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
		CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(itemIndex, this);
		return Fire_ListClick(pDCBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ListCloseUp(void)
{
	if(flags.waitForCloseUp) {
		int selectedItem = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		if(selectedItem != selectedItemOnDropDown) {
			Raise_SelectedDriveChanged(selectedItemOnDropDown, selectedItem);
		}
	}
	flags.waitForCloseUp = FALSE;
	//if(m_nFreezeEvents == 0) {
		return Fire_ListCloseUp();
	//}
}

inline HRESULT DriveComboBox::Raise_ListDropDown(void)
{
	flags.waitForCloseUp = TRUE;
	selectedItemOnDropDown = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
	//if(m_nFreezeEvents == 0) {
		return Fire_ListDropDown();
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ListMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deListBoxMouseEvents)) {
		mouseStatus_ListBox.StoreClickCandidate(button);
		//if(button != 1/*MouseButtonConstants.vbLeftButton*/) {
		//	containedListBox.SetCapture();
		//}

		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
		CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(itemIndex, this);
		return Fire_ListMouseDown(pDCBItem, button, shift, x, y, hitTestDetails);
	} else {
		mouseStatus_ListBox.StoreClickCandidate(button);
		//if(button != 1/*MouseButtonConstants.vbLeftButton*/) {
		//	containedListBox.SetCapture();
		//}
		return S_OK;
	}
}

inline HRESULT DriveComboBox::Raise_ListMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus_ListBox.lastPosition.x = x;
	mouseStatus_ListBox.lastPosition.y = y;

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
	CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(itemIndex, this);

	// Do we move over another date than before?
	if(itemIndex != itemUnderMouse) {
		if(itemUnderMouse >= 0) {
			CComPtr<IDriveComboBoxItem> pPrevICBItem = ClassFactory::InitDriveComboItem(itemUnderMouse, this);
			Raise_ItemMouseLeave(pPrevICBItem, button, shift, x, y, hitTestDetails);
		}
		itemUnderMouse = itemIndex;
		if(pDCBItem) {
			Raise_ItemMouseEnter(pDCBItem, button, shift, x, y, hitTestDetails);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_ListMouseMove(pDCBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ListMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deListBoxMouseEvents)) {
		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
		CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(itemIndex, this);
		hr = Fire_ListMouseUp(pDCBItem, button, shift, x, y, hitTestDetails);
	}

	if(mouseStatus_ListBox.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		containedListBox.GetClientRect(&clientArea);
		containedListBox.ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != containedListBox) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		//if(button != 1/*MouseButtonConstants.vbLeftButton*/ && GetCapture() == containedListBox) {
		//	ReleaseCapture();
		//}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deListBoxClickEvents)) {
						Raise_ListClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus_ListBox.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT DriveComboBox::Raise_ListMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
	CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(itemIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_ListMouseWheel(pDCBItem, button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ListOLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

	POINTL cursorPosition = mousePosition;
	containedListBox.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = ListBoxHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IDriveComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitDriveComboItem(dropTarget, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtSelection, NULL, &yItem);
		yToItemTop = mousePosition.y - yItem;
	}

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	if(properties.handleOLEDragDrop & hoddTargetPart) {
		CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
		int dropTargetBeforeEvent = dropTarget;
		CComPtr<IDropTarget> pNewDropTarget = NULL;
		BOOL isSameDropTarget = (dropTargetBeforeEvent == dragDropStatus.currentDropTarget);
		PIDLIST_ABSOLUTE pIDLDropTarget = NULL;
		if(dropTargetBeforeEvent != -2) {
			pIDLDropTarget = GetItemPIDL(dropTargetBeforeEvent);
		}
		if(pIDLDropTarget) {
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}
		if(!pNewDropTarget) {
			*pEffect = DROPEFFECT_NONE;
		}

		LPDROPTARGET pEnteredDropTarget = NULL;
		// let the shell propose a drop effect
		if(pNewDropTarget) {
			if(isSameDropTarget) {
				ATLVERIFY(SUCCEEDED(pNewDropTarget->DragOver(dragDropStatus.lastKeyState, cursorPosition, pEffect)));
			} else {
				pNewDropTarget->DragEnter(pDataObject, dragDropStatus.lastKeyState, cursorPosition, pEffect);
				pEnteredDropTarget = pNewDropTarget;
			}
		}

		// fire event
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_ListOLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails);
			}
		//}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}
		dragDropStatus.lastDropTarget = dropTarget;

		// Did the event handler change the drop target?
		int dropTargetAfterEvent = dropTarget;
		BOOL eventHandlerHasChangedTarget = (dropTargetAfterEvent != dropTargetBeforeEvent);
		isSameDropTarget = (dropTargetAfterEvent == dragDropStatus.currentDropTarget);
		if(eventHandlerHasChangedTarget) {
			// update pNewDropTarget
			if(pNewDropTarget) {
				pNewDropTarget->DragLeave();
				pNewDropTarget = NULL;
			}
			if(dropTargetAfterEvent != -2) {
				if(!isSameDropTarget) {
					ILFree(pIDLDropTarget);
					pIDLDropTarget = GetItemPIDL(dropTargetAfterEvent);
				}
			} else {
				ILFree(pIDLDropTarget);
				pIDLDropTarget = NULL;
			}
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else if(pIDLDropTarget) {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}

		// now we don't have to worry about dropTargetBeforeEvent anymore
		if(!isSameDropTarget) {
			// leave the current drop target
			if(dragDropStatus.pCurrentDropTarget) {
				dragDropStatus.pCurrentDropTarget->DragLeave();
				dragDropStatus.pCurrentDropTarget = NULL;
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragLeave()));
			}

			// enter the new drop target
			dragDropStatus.pCurrentDropTarget = pNewDropTarget;
			if(dragDropStatus.pCurrentDropTarget && dragDropStatus.pCurrentDropTarget != pEnteredDropTarget) {
				DWORD effect = *pEffect;
				dragDropStatus.pCurrentDropTarget->DragEnter(pDataObject, dragDropStatus.lastKeyState, cursorPosition, &effect);
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragEnter(containedListBox, pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
			if(dragDropStatus.pCurrentDropTarget) {
				DWORD effect = *pEffect;
				ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragOver(dragDropStatus.lastKeyState, cursorPosition, &effect)));
			}
		}
		dragDropStatus.currentDropTarget = dropTargetAfterEvent;

		/* Now we have called DragEnter/DragOver for the correct drop target and left the old drop target if
		 * necessary. We've also made the necessary calls to IDropTargetHelper.
		 * Now we can handle the drop.
		 */

		if(dragDropStatus.pCurrentDropTarget) {
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->Drop(pDataObject, keyState, cursorPosition, pEffect)));
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(dragDropStatus.pDropTargetHelper) {
			ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->Drop(pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
		}
		dragDropStatus.currentDropTarget = -2;
		if(pCallDropTargetHelper) {
			*pCallDropTargetHelper = FALSE;
			dragDropStatus.drop_pDataObject = NULL;
		}
		ILFree(pIDLDropTarget);

	} else {
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_ListOLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails);
			}
		//}
	}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.ListOLEDragLeaveOrDrop();
	if(containedListBox.IsWindow()) {
		containedListBox.Invalidate();
	}

	return hr;
}

inline HRESULT DriveComboBox::Raise_ListOLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	POINTL cursorPosition = mousePosition;
	containedListBox.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.ListOLEDragEnter();

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = ListBoxHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IDriveComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitDriveComboItem(dropTarget, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtSelection, NULL, &yItem);
		yToItemTop = mousePosition.y - yItem;
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.listDragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone;
		containedListBox.GetClientRect(&noScrollZone);
		BOOL isInScrollZone = (noScrollZone.PtInRect(mousePos) == TRUE);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	dragDropStatus.lastKeyState = keyState;
	if(properties.handleOLEDragDrop & hoddTargetPart) {
		CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
		int dropTargetBeforeEvent = dropTarget;
		CComPtr<IDropTarget> pNewDropTarget = NULL;
		BOOL isSameDropTarget = (dropTargetBeforeEvent == dragDropStatus.currentDropTarget);
		PIDLIST_ABSOLUTE pIDLDropTarget = NULL;
		if(dropTargetBeforeEvent != -2) {
			pIDLDropTarget = GetItemPIDL(dropTargetBeforeEvent);
		}
		if(pIDLDropTarget) {
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}
		if(!pNewDropTarget) {
			*pEffect = DROPEFFECT_NONE;
		}

		LPDROPTARGET pEnteredDropTarget = NULL;
		// let the shell propose a drop effect
		if(pNewDropTarget) {
			if(isSameDropTarget) {
				ATLVERIFY(SUCCEEDED(pNewDropTarget->DragOver(keyState, cursorPosition, pEffect)));
			} else {
				pNewDropTarget->DragEnter(pDataObject, keyState, cursorPosition, pEffect);
				pEnteredDropTarget = pNewDropTarget;
			}
		}

		// fire event
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_ListOLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
			}
		//}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}
		dragDropStatus.lastDropTarget = dropTarget;

		// Did the event handler change the drop target?
		int dropTargetAfterEvent = dropTarget;
		BOOL eventHandlerHasChangedTarget = (dropTargetAfterEvent != dropTargetBeforeEvent);
		isSameDropTarget = (dropTargetAfterEvent == dragDropStatus.currentDropTarget);
		if(eventHandlerHasChangedTarget) {
			// update pNewDropTarget
			if(pNewDropTarget) {
				pNewDropTarget->DragLeave();
				pNewDropTarget = NULL;
			}
			if(dropTargetAfterEvent != -2) {
				if(!isSameDropTarget) {
					ILFree(pIDLDropTarget);
					pIDLDropTarget = GetItemPIDL(dropTargetAfterEvent);
				}
			} else {
				ILFree(pIDLDropTarget);
				pIDLDropTarget = NULL;
			}
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else if(pIDLDropTarget) {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}

		// now we don't have to worry about dropTargetBeforeEvent anymore
		if(isSameDropTarget) {
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragEnter(containedListBox, pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
		} else {
			// leave the current drop target
			if(dragDropStatus.pCurrentDropTarget) {
				dragDropStatus.pCurrentDropTarget->DragLeave();
				dragDropStatus.pCurrentDropTarget = NULL;
			}
			// enter the new drop target
			dragDropStatus.pCurrentDropTarget = pNewDropTarget;
			if(dragDropStatus.pCurrentDropTarget && dragDropStatus.pCurrentDropTarget != pEnteredDropTarget) {
				DWORD effect = *pEffect;
				dragDropStatus.pCurrentDropTarget->DragEnter(pDataObject, keyState, cursorPosition, &effect);
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragEnter(containedListBox, pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
		}
		dragDropStatus.currentDropTarget = dropTargetAfterEvent;
		if(pCallDropTargetHelper) {
			*pCallDropTargetHelper = FALSE;
		}
		ILFree(pIDLDropTarget);

	} else {
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_ListOLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
			}
		//}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}

		dragDropStatus.lastDropTarget = dropTarget;
	}

	if(properties.listDragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.listDragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.listDragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			ListBoxAutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT DriveComboBox::Raise_ListOLEDragLeave(BOOL fakedLeave, BOOL* pCallDropTargetHelper)
{
	KillTimer(timers.ID_DRAGSCROLL);
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = ListBoxHitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	CComPtr<IDriveComboBoxItem> pDropTarget = ClassFactory::InitDriveComboItem(dropTarget, this);

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtSelection, NULL, &yItem);
		yToItemTop = dragDropStatus.lastMousePosition.y - yItem;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		VARIANT_BOOL autoCloseUp = VARIANT_TRUE;
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_ListOLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, yToItemTop, hitTestDetails, &autoCloseUp);
		}
	//}

	if(properties.handleOLEDragDrop & hoddTargetPart) {
		if(dragDropStatus.pCurrentDropTarget) {
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragLeave()));
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(dragDropStatus.pDropTargetHelper) {
			ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragLeave()));
		}
		dragDropStatus.currentDropTarget = -2;
		if(pCallDropTargetHelper) {
			*pCallDropTargetHelper = FALSE;
		}
	}

	if(fakedLeave) {
		dragDropStatus.isOverListBox = FALSE;
	} else {
		dragDropStatus.pActiveDataObject = NULL;
		dragDropStatus.ListOLEDragLeaveOrDrop();
		containedListBox.Invalidate();
	}
	if(autoCloseUp != VARIANT_FALSE) {
		CloseDropDownWindow();
	}
	return hr;
}

inline HRESULT DriveComboBox::Raise_ListOLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	POINTL cursorPosition = mousePosition;
	containedListBox.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = ListBoxHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IDriveComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitDriveComboItem(dropTarget, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtSelection, NULL, &yItem);
		yToItemTop = mousePosition.y - yItem;
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.listDragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone;
		containedListBox.GetClientRect(&noScrollZone);
		BOOL isInScrollZone = (noScrollZone.PtInRect(mousePos) == TRUE);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	dragDropStatus.lastKeyState = keyState;
	if(properties.handleOLEDragDrop & hoddTargetPart) {
		CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
		int dropTargetBeforeEvent = dropTarget;
		CComPtr<IDropTarget> pNewDropTarget = NULL;
		BOOL isSameDropTarget = (dropTargetBeforeEvent == dragDropStatus.currentDropTarget);
		PIDLIST_ABSOLUTE pIDLDropTarget = NULL;
		if(dropTargetBeforeEvent != -2) {
			pIDLDropTarget = GetItemPIDL(dropTargetBeforeEvent);
		}
		if(pIDLDropTarget) {
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}
		if(!pNewDropTarget) {
			*pEffect = DROPEFFECT_NONE;
		}

		LPDROPTARGET pEnteredDropTarget = NULL;
		// let the shell propose a drop effect
		if(pNewDropTarget) {
			if(isSameDropTarget) {
				ATLVERIFY(SUCCEEDED(pNewDropTarget->DragOver(keyState, cursorPosition, pEffect)));
			} else {
				pNewDropTarget->DragEnter(pDataObject, keyState, cursorPosition, pEffect);
				pEnteredDropTarget = pNewDropTarget;
			}
		}

		// fire event
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_ListOLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
			}
		//}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}
		dragDropStatus.lastDropTarget = dropTarget;

		// Did the event handler change the drop target?
		int dropTargetAfterEvent = dropTarget;
		BOOL eventHandlerHasChangedTarget = (dropTargetAfterEvent != dropTargetBeforeEvent);
		isSameDropTarget = (dropTargetAfterEvent == dragDropStatus.currentDropTarget);
		if(eventHandlerHasChangedTarget) {
			// update pNewDropTarget
			if(pNewDropTarget) {
				pNewDropTarget->DragLeave();
				pNewDropTarget = NULL;
			}
			if(dropTargetAfterEvent != -2) {
				if(!isSameDropTarget) {
					ILFree(pIDLDropTarget);
					pIDLDropTarget = GetItemPIDL(dropTargetAfterEvent);
				}
			} else {
				ILFree(pIDLDropTarget);
				pIDLDropTarget = NULL;
			}
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else if(pIDLDropTarget) {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}

		// now we don't have to worry about dropTargetBeforeEvent anymore
		if(isSameDropTarget) {
			// just call DragOver
			if(eventHandlerHasChangedTarget) {
				// we entered this target on a previous call of this method, but did not yet call DragOver this time
				if(dragDropStatus.pCurrentDropTarget) {
					DWORD effect = *pEffect;
					ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragOver(keyState, cursorPosition, &effect)));
				}
			} else {
				// we already called DragOver on this target before firing the event
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragOver(reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
		} else {
			// leave the current drop target
			if(dragDropStatus.pCurrentDropTarget) {
				dragDropStatus.pCurrentDropTarget->DragLeave();
				dragDropStatus.pCurrentDropTarget = NULL;
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragLeave()));
			}

			// enter the new drop target
			dragDropStatus.pCurrentDropTarget = pNewDropTarget;
			if(dragDropStatus.pCurrentDropTarget && dragDropStatus.pCurrentDropTarget != pEnteredDropTarget) {
				DWORD effect = *pEffect;
				dragDropStatus.pCurrentDropTarget->DragEnter(pDataObject, keyState, cursorPosition, &effect);
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragEnter(containedListBox, pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
		}
		dragDropStatus.currentDropTarget = dropTargetAfterEvent;
		if(pCallDropTargetHelper) {
			*pCallDropTargetHelper = FALSE;
		}
		ILFree(pIDLDropTarget);

	} else {
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_ListOLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
			}
		//}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}

		dragDropStatus.lastDropTarget = dropTarget;
	}

	if(properties.listDragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.listDragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.listDragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			ListBoxAutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT DriveComboBox::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htComboBoxPortion) {
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT(hitTestDetails & htComboBoxPortion);
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!pMouseStatusToUse->enteredControl) {
			Raise_MouseEnter(button, shift, x, y, hitTestDetails);
		}
		if(!pMouseStatusToUse->hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			trackingOptions.hwndTrack = *this;
			TrackMouseEvent(&trackingOptions);
			if(containedComboBox.IsWindow()) {
				trackingOptions.hwndTrack = containedComboBox;
				TrackMouseEvent(&trackingOptions);
			}

			Raise_MouseHover(button, shift, x, y, hitTestDetails);
		}
		pMouseStatusToUse->StoreClickCandidate(button);
		if(hitTestDetails & htComboBoxPortion) {
			containedComboBox.SetCapture();
		}

		return Fire_MouseDown(button, shift, x, y, hitTestDetails);
	} else {
		if(!pMouseStatusToUse->enteredControl) {
			pMouseStatusToUse->EnterControl();
		}
		if(!pMouseStatusToUse->hoveredControl) {
			pMouseStatusToUse->HoverControl();
		}
		pMouseStatusToUse->StoreClickCandidate(button);
		if(hitTestDetails & htComboBoxPortion) {
			containedComboBox.SetCapture();
		}
		return S_OK;
	}
}

inline HRESULT DriveComboBox::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htComboBoxPortion) {
		trackingOptions.hwndTrack = containedComboBox;
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT(hitTestDetails & htComboBoxPortion);
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	pMouseStatusToUse->EnterControl();

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseEnter(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(hitTestDetails & htComboBoxPortion) {
		mouseStatus_ComboBox.HoverControl();
	} else {
		ATLASSERT(hitTestDetails & htComboBoxPortion);
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT DriveComboBox::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(hitTestDetails & htComboBoxPortion) {
		mouseStatus_ComboBox.LeaveControl();
	} else {
		ATLASSERT(hitTestDetails & htComboBoxPortion);
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT DriveComboBox::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htComboBoxPortion) {
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT(hitTestDetails & htComboBoxPortion);
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(!pMouseStatusToUse->enteredControl) {
		Raise_MouseEnter(button, shift, x, y, hitTestDetails);
	}
	mouseStatus_ComboBox.lastPosition.x = x;
	mouseStatus_ComboBox.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y, hitTestDetails);
	}

	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htComboBoxPortion) {
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT(hitTestDetails & htComboBoxPortion);
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(pMouseStatusToUse->IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			HWND hWindowBelowMouse = WindowFromPoint(cursorPosition);
			if((hWindowBelowMouse != *this) && !IsChild(hWindowBelowMouse)) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(hitTestDetails & htComboBoxPortion) {
			if(GetCapture() == containedComboBox) {
				ReleaseCapture();
			}
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y, hitTestDetails);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y, hitTestDetails);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y, hitTestDetails);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y, hitTestDetails);
					}
					break;
			}
		}

		pMouseStatusToUse->RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT DriveComboBox::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta, HitTestConstants hitTestDetails)
{
	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htComboBoxPortion) {
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT(hitTestDetails & htComboBoxPortion);
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(!pMouseStatusToUse->enteredControl) {
		Raise_MouseEnter(button, shift, x, y, hitTestDetails);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGDROPDOWN);
	dragDropStatus.dropDownTimerIsRunning = FALSE;

	POINTL cursorPosition = mousePosition;
	containedComboBox.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = -1;
	if(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails))) {
		if(hitTestDetails == htComboBoxPortion) {
			dropTarget = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
	}
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IDriveComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitDriveComboItem(dropTarget, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	if(properties.handleOLEDragDrop & hoddTargetPart) {
		CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
		int dropTargetBeforeEvent = dropTarget;
		CComPtr<IDropTarget> pNewDropTarget = NULL;
		BOOL isSameDropTarget = (dropTargetBeforeEvent == dragDropStatus.currentDropTarget);
		PIDLIST_ABSOLUTE pIDLDropTarget = NULL;
		if(dropTargetBeforeEvent != -2) {
			pIDLDropTarget = GetItemPIDL(dropTargetBeforeEvent);
		}
		if(pIDLDropTarget) {
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}
		if(!pNewDropTarget) {
			*pEffect = DROPEFFECT_NONE;
		}

		LPDROPTARGET pEnteredDropTarget = NULL;
		// let the shell propose a drop effect
		if(pNewDropTarget) {
			if(isSameDropTarget) {
				ATLVERIFY(SUCCEEDED(pNewDropTarget->DragOver(dragDropStatus.lastKeyState, cursorPosition, pEffect)));
			} else {
				pNewDropTarget->DragEnter(pDataObject, dragDropStatus.lastKeyState, cursorPosition, pEffect);
				pEnteredDropTarget = pNewDropTarget;
			}
		}

		// fire event
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
			}
		//}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}
		dragDropStatus.lastDropTarget = dropTarget;

		// Did the event handler change the drop target?
		int dropTargetAfterEvent = dropTarget;
		BOOL eventHandlerHasChangedTarget = (dropTargetAfterEvent != dropTargetBeforeEvent);
		isSameDropTarget = (dropTargetAfterEvent == dragDropStatus.currentDropTarget);
		if(eventHandlerHasChangedTarget) {
			// update pNewDropTarget
			if(pNewDropTarget) {
				pNewDropTarget->DragLeave();
				pNewDropTarget = NULL;
			}
			if(dropTargetAfterEvent != -2) {
				if(!isSameDropTarget) {
					ILFree(pIDLDropTarget);
					pIDLDropTarget = GetItemPIDL(dropTargetAfterEvent);
				}
			} else {
				ILFree(pIDLDropTarget);
				pIDLDropTarget = NULL;
			}
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else if(pIDLDropTarget) {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}

		// now we don't have to worry about dropTargetBeforeEvent anymore
		if(!isSameDropTarget) {
			// leave the current drop target
			if(dragDropStatus.pCurrentDropTarget) {
				dragDropStatus.pCurrentDropTarget->DragLeave();
				dragDropStatus.pCurrentDropTarget = NULL;
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragLeave()));
			}

			// enter the new drop target
			dragDropStatus.pCurrentDropTarget = pNewDropTarget;
			if(dragDropStatus.pCurrentDropTarget && dragDropStatus.pCurrentDropTarget != pEnteredDropTarget) {
				DWORD effect = *pEffect;
				dragDropStatus.pCurrentDropTarget->DragEnter(pDataObject, dragDropStatus.lastKeyState, cursorPosition, &effect);
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
			if(dragDropStatus.pCurrentDropTarget) {
				DWORD effect = *pEffect;
				ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragOver(dragDropStatus.lastKeyState, cursorPosition, &effect)));
			}
		}
		dragDropStatus.currentDropTarget = dropTargetAfterEvent;

		/* Now we have called DragEnter/DragOver for the correct drop target and left the old drop target if
		 * necessary. We've also made the necessary calls to IDropTargetHelper.
		 * Now we can handle the drop.
		 */

		if(dragDropStatus.pCurrentDropTarget) {
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->Drop(pDataObject, keyState, cursorPosition, pEffect)));
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(dragDropStatus.pDropTargetHelper) {
			ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->Drop(pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
		}
		dragDropStatus.currentDropTarget = -2;
		if(pCallDropTargetHelper) {
			*pCallDropTargetHelper = FALSE;
			dragDropStatus.drop_pDataObject = NULL;
		}
		ILFree(pIDLDropTarget);

	} else {
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
			}
		//}
	}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	containedComboBox.Invalidate();

	return hr;
}

inline HRESULT DriveComboBox::Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	POINTL cursorPosition = mousePosition;
	containedComboBox.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	VARIANT_BOOL autoDropDown = VARIANT_FALSE;
	if(GetStyle() & (CBS_DROPDOWN | CBS_DROPDOWNLIST)) {
		POINT pt = {mousePosition.x, mousePosition.y};
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(containedComboBox, &controlInfo);
		autoDropDown = BOOL2VARIANTBOOL(PtInRect(&controlInfo.rcButton, pt));
	}

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = -1;
	if(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails))) {
		if(hitTestDetails == htComboBoxPortion) {
			dropTarget = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
	}
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IDriveComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitDriveComboItem(dropTarget, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	dragDropStatus.lastKeyState = keyState;
	if(properties.handleOLEDragDrop & hoddTargetPart) {
		CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
		int dropTargetBeforeEvent = dropTarget;
		CComPtr<IDropTarget> pNewDropTarget = NULL;
		BOOL isSameDropTarget = (dropTargetBeforeEvent == dragDropStatus.currentDropTarget);
		PIDLIST_ABSOLUTE pIDLDropTarget = NULL;
		if(dropTargetBeforeEvent != -2) {
			pIDLDropTarget = GetItemPIDL(dropTargetBeforeEvent);
		}
		if(pIDLDropTarget) {
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}
		if(!pNewDropTarget) {
			*pEffect = DROPEFFECT_NONE;
		}

		LPDROPTARGET pEnteredDropTarget = NULL;
		// let the shell propose a drop effect
		if(pNewDropTarget) {
			if(isSameDropTarget) {
				ATLVERIFY(SUCCEEDED(pNewDropTarget->DragOver(keyState, cursorPosition, pEffect)));
			} else {
				pNewDropTarget->DragEnter(pDataObject, keyState, cursorPosition, pEffect);
				pEnteredDropTarget = pNewDropTarget;
			}
		}

		// fire event
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails, &autoDropDown);
			}
		//}

		if(autoDropDown == VARIANT_FALSE) {
			// cancel automatic drop-down
			KillTimer(timers.ID_DRAGDROPDOWN);
			dragDropStatus.dropDownTimerIsRunning = FALSE;
		} else {
			if(properties.dragDropDownTime && !dragDropStatus.dropDownTimerIsRunning) {
				// start timered automatic drop-down
				dragDropStatus.dropDownTimerIsRunning = TRUE;
				SetTimer(timers.ID_DRAGDROPDOWN, (properties.dragDropDownTime == -1 ? GetDoubleClickTime() * 4 : properties.dragDropDownTime));
			}
		}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}
		dragDropStatus.lastDropTarget = dropTarget;

		// Did the event handler change the drop target?
		int dropTargetAfterEvent = dropTarget;
		BOOL eventHandlerHasChangedTarget = (dropTargetAfterEvent != dropTargetBeforeEvent);
		isSameDropTarget = (dropTargetAfterEvent == dragDropStatus.currentDropTarget);
		if(eventHandlerHasChangedTarget) {
			// update pNewDropTarget
			if(pNewDropTarget) {
				pNewDropTarget->DragLeave();
				pNewDropTarget = NULL;
			}
			if(dropTargetAfterEvent != -2) {
				if(!isSameDropTarget) {
					ILFree(pIDLDropTarget);
					pIDLDropTarget = GetItemPIDL(dropTargetAfterEvent);
				}
			} else {
				ILFree(pIDLDropTarget);
				pIDLDropTarget = NULL;
			}
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else if(pIDLDropTarget) {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}

		// now we don't have to worry about dropTargetBeforeEvent anymore
		if(isSameDropTarget) {
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
		} else {
			// leave the current drop target
			if(dragDropStatus.pCurrentDropTarget) {
				dragDropStatus.pCurrentDropTarget->DragLeave();
				dragDropStatus.pCurrentDropTarget = NULL;
			}
			// enter the new drop target
			dragDropStatus.pCurrentDropTarget = pNewDropTarget;
			if(dragDropStatus.pCurrentDropTarget && dragDropStatus.pCurrentDropTarget != pEnteredDropTarget) {
				DWORD effect = *pEffect;
				dragDropStatus.pCurrentDropTarget->DragEnter(pDataObject, keyState, cursorPosition, &effect);
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
		}
		dragDropStatus.currentDropTarget = dropTargetAfterEvent;
		if(pCallDropTargetHelper) {
			*pCallDropTargetHelper = FALSE;
		}
		ILFree(pIDLDropTarget);

	} else {
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails, &autoDropDown);
			}
		//}

		if(autoDropDown == VARIANT_FALSE) {
			// cancel automatic drop-down
			KillTimer(timers.ID_DRAGDROPDOWN);
			dragDropStatus.dropDownTimerIsRunning = FALSE;
		} else {
			if(properties.dragDropDownTime && !dragDropStatus.dropDownTimerIsRunning) {
				// start timered automatic drop-down
				dragDropStatus.dropDownTimerIsRunning = TRUE;
				SetTimer(timers.ID_DRAGDROPDOWN, (properties.dragDropDownTime == -1 ? GetDoubleClickTime() * 4 : properties.dragDropDownTime));
			}
		}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}

		dragDropStatus.lastDropTarget = dropTarget;
	}
	return hr;
}

inline HRESULT DriveComboBox::Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_OLEDragLeave(BOOL fakedLeave, BOOL* pCallDropTargetHelper)
{
	KillTimer(timers.ID_DRAGDROPDOWN);
	dragDropStatus.dropDownTimerIsRunning = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = -1;
	if(SUCCEEDED(HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails))) {
		if(hitTestDetails == htComboBoxPortion) {
			dropTarget = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
	}
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	CComPtr<IDriveComboBoxItem> pDropTarget = ClassFactory::InitDriveComboItem(dropTarget, this);

	// fire event
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, hitTestDetails);
		}
	//}

	if(properties.handleOLEDragDrop & hoddTargetPart) {
		if(dragDropStatus.pCurrentDropTarget) {
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragLeave()));
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(dragDropStatus.pDropTargetHelper) {
			ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragLeave()));
		}
		dragDropStatus.currentDropTarget = -2;
		if(pCallDropTargetHelper) {
			*pCallDropTargetHelper = FALSE;
		}
	}

	if(!fakedLeave) {
		dragDropStatus.pActiveDataObject = NULL;
		dragDropStatus.OLEDragLeaveOrDrop();
		containedComboBox.Invalidate();
	}
	return hr;
}

inline HRESULT DriveComboBox::Raise_OLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	POINTL cursorPosition = mousePosition;
	containedComboBox.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	VARIANT_BOOL autoDropDown = VARIANT_FALSE;
	if(GetStyle() & (CBS_DROPDOWN | CBS_DROPDOWNLIST)) {
		POINT pt = {mousePosition.x, mousePosition.y};
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(containedComboBox, &controlInfo);
		autoDropDown = BOOL2VARIANTBOOL(PtInRect(&controlInfo.rcButton, pt));
	}

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = -1;
	if(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails))) {
		if(hitTestDetails == htComboBoxPortion) {
			dropTarget = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
	}
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IDriveComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitDriveComboItem(dropTarget, this, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	HRESULT hr = S_OK;
	dragDropStatus.lastKeyState = keyState;
	if(properties.handleOLEDragDrop & hoddTargetPart) {
		CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
		int dropTargetBeforeEvent = dropTarget;
		CComPtr<IDropTarget> pNewDropTarget = NULL;
		BOOL isSameDropTarget = (dropTargetBeforeEvent == dragDropStatus.currentDropTarget);
		PIDLIST_ABSOLUTE pIDLDropTarget = NULL;
		if(dropTargetBeforeEvent != -2) {
			pIDLDropTarget = GetItemPIDL(dropTargetBeforeEvent);
		}
		if(pIDLDropTarget) {
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}
		if(!pNewDropTarget) {
			*pEffect = DROPEFFECT_NONE;
		}

		LPDROPTARGET pEnteredDropTarget = NULL;
		// let the shell propose a drop effect
		if(pNewDropTarget) {
			if(isSameDropTarget) {
				ATLVERIFY(SUCCEEDED(pNewDropTarget->DragOver(keyState, cursorPosition, pEffect)));
			} else {
				pNewDropTarget->DragEnter(pDataObject, keyState, cursorPosition, pEffect);
				pEnteredDropTarget = pNewDropTarget;
			}
		}

		// fire event
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails, &autoDropDown);
			}
		//}

		if(autoDropDown == VARIANT_FALSE) {
			// cancel automatic drop-down
			KillTimer(timers.ID_DRAGDROPDOWN);
			dragDropStatus.dropDownTimerIsRunning = FALSE;
		} else {
			if(properties.dragDropDownTime && !dragDropStatus.dropDownTimerIsRunning) {
				// start timered automatic drop-down
				dragDropStatus.dropDownTimerIsRunning = TRUE;
				SetTimer(timers.ID_DRAGDROPDOWN, (properties.dragDropDownTime == -1 ? GetDoubleClickTime() * 4 : properties.dragDropDownTime));
			}
		}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}
		dragDropStatus.lastDropTarget = dropTarget;

		// Did the event handler change the drop target?
		int dropTargetAfterEvent = dropTarget;
		BOOL eventHandlerHasChangedTarget = (dropTargetAfterEvent != dropTargetBeforeEvent);
		isSameDropTarget = (dropTargetAfterEvent == dragDropStatus.currentDropTarget);
		if(eventHandlerHasChangedTarget) {
			// update pNewDropTarget
			if(pNewDropTarget) {
				pNewDropTarget->DragLeave();
				pNewDropTarget = NULL;
			}
			if(dropTargetAfterEvent != -2) {
				if(!isSameDropTarget) {
					ILFree(pIDLDropTarget);
					pIDLDropTarget = GetItemPIDL(dropTargetAfterEvent);
				}
			} else {
				ILFree(pIDLDropTarget);
				pIDLDropTarget = NULL;
			}
			if(isSameDropTarget) {
				pNewDropTarget = dragDropStatus.pCurrentDropTarget;
			} else if(pIDLDropTarget) {
				::GetDropTarget(GethWndShellUIParentWindow(), pIDLDropTarget, &pNewDropTarget, FALSE);
			}
		}

		// now we don't have to worry about dropTargetBeforeEvent anymore
		if(isSameDropTarget) {
			// just call DragOver
			if(eventHandlerHasChangedTarget) {
				// we entered this target on a previous call of this method, but did not yet call DragOver this time
				if(dragDropStatus.pCurrentDropTarget) {
					DWORD effect = *pEffect;
					ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragOver(keyState, cursorPosition, &effect)));
				}
			} else {
				// we already called DragOver on this target before firing the event
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragOver(reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
		} else {
			// leave the current drop target
			if(dragDropStatus.pCurrentDropTarget) {
				dragDropStatus.pCurrentDropTarget->DragLeave();
				dragDropStatus.pCurrentDropTarget = NULL;
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragLeave()));
			}

			// enter the new drop target
			dragDropStatus.pCurrentDropTarget = pNewDropTarget;
			if(dragDropStatus.pCurrentDropTarget && dragDropStatus.pCurrentDropTarget != pEnteredDropTarget) {
				DWORD effect = *pEffect;
				dragDropStatus.pCurrentDropTarget->DragEnter(pDataObject, keyState, cursorPosition, &effect);
			}
			if(dragDropStatus.pDropTargetHelper) {
				ATLVERIFY(SUCCEEDED(dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, reinterpret_cast<LPPOINT>(&cursorPosition), *pEffect)));
			}
		}
		dragDropStatus.currentDropTarget = dropTargetAfterEvent;
		if(pCallDropTargetHelper) {
			*pCallDropTargetHelper = FALSE;
		}
		ILFree(pIDLDropTarget);

	} else {
		//if(m_nFreezeEvents == 0) {
			if(dragDropStatus.pActiveDataObject) {
				hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails, &autoDropDown);
			}
		//}

		if(autoDropDown == VARIANT_FALSE) {
			// cancel automatic drop-down
			KillTimer(timers.ID_DRAGDROPDOWN);
			dragDropStatus.dropDownTimerIsRunning = FALSE;
		} else {
			if(properties.dragDropDownTime && !dragDropStatus.dropDownTimerIsRunning) {
				// start timered automatic drop-down
				dragDropStatus.dropDownTimerIsRunning = TRUE;
				SetTimer(timers.ID_DRAGDROPDOWN, (properties.dragDropDownTime == -1 ? GetDoubleClickTime() * 4 : properties.dragDropDownTime));
			}
		}

		if(pDropTarget) {
			// TODO: Retrieving just the index isn't enough!
			LONG l = -1;
			pDropTarget->get_Index(&l);
			dropTarget = l;
			// we're using a raw pointer
			pDropTarget->Release();
		} else {
			dropTarget = -1;
		}

		dragDropStatus.lastDropTarget = dropTarget;
	}
	return hr;
}

inline HRESULT DriveComboBox::Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
{
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
		return Fire_OLEQueryContinueDrag(BOOL2VARIANTBOOL(pressedEscape), button, shift, reinterpret_cast<OLEActionToContinueWithConstants*>(pActionToContinueWith));
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT DriveComboBox::Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT DriveComboBox::Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLESetData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_OLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_OutOfMemory(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OutOfMemory();
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_RecreatedControlWindow(HWND hWnd)
{
	properties.selectionFieldItemProperties.mask = 0;
	properties.selectionFieldItemProperties.iImage = I_IMAGECALLBACK;
	properties.selectionFieldItemProperties.iSelectedImage = I_IMAGECALLBACK;
	properties.selectionFieldItemProperties.iOverlay = I_IMAGECALLBACK;

	HWND hWndCombo = reinterpret_cast<HWND>(SendMessage(CBEM_GETCOMBOCONTROL, 0, 0));
	if(hWndCombo) {
		Raise_CreatedComboBoxControlWindow(hWndCombo);
	}

	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_RemovedItem(IVirtualDriveComboBoxItem* pComboItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedItem(pComboItem);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_RemovingItem(IDriveComboBoxItem* pComboItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingItem(pComboItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_SelectedDriveChanged(int previousSelectedItem, int newSelectedItem)
{
	//if(m_nFreezeEvents == 0) {
		int selectedItem = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		if(selectedItem != selectedItemOnDropDown) {
			selectedItemOnDropDown = selectedItem;
			if(previousSelectedItem != newSelectedItem) {
				// NOTE: We can't use selectedItemOnDropDown directly, because selection may change without ListDropDown.
				CComPtr<IDriveComboBoxItem> pDCBPreviousSelectedItem = ClassFactory::InitDriveComboItem(previousSelectedItem, this);
				CComPtr<IDriveComboBoxItem> pDCBNewSelectedItem = ClassFactory::InitDriveComboItem(newSelectedItem, this);
				return Fire_SelectedDriveChanged(pDCBPreviousSelectedItem, pDCBNewSelectedItem);
			}
		}
	//}
	return S_OK;
}

inline HRESULT DriveComboBox::Raise_SelectionCanceled(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectionCanceled();
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_SelectionChanged(int previousSelectedItem, int newSelectedItem)
{
	currentSelectedItem = newSelectedItem;
	//if(m_nFreezeEvents == 0) {
		if(previousSelectedItem != newSelectedItem) {
			CComPtr<IDriveComboBoxItem> pDCBPreviousSelectedItem = ClassFactory::InitDriveComboItem(previousSelectedItem, this);
			CComPtr<IDriveComboBoxItem> pDCBNewSelectedItem = ClassFactory::InitDriveComboItem(newSelectedItem, this);
			if(!flags.waitForCloseUp) {
				Raise_SelectedDriveChanged(previousSelectedItem, newSelectedItem);
			}
			return Fire_SelectionChanged(pDCBPreviousSelectedItem, pDCBNewSelectedItem);
		}
	//}
	return S_OK;
}

inline HRESULT DriveComboBox::Raise_SelectionChanging(int newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* pCancelChange)
{
	//if(m_nFreezeEvents == 0) {
		CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(newSelectedItem, this);
		return Fire_SelectionChanging(pDCBItem, selectionFieldText, selectionFieldHasBeenEdited, selectionChangeReason, pCancelChange);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT DriveComboBox::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}


void DriveComboBox::RecreateControlWindow(void)
{
	if(m_bInPlaceActive && !flags.dontRecreate) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));

		// HACK: Sometimes the inner combo box or edit box gets created with a width of 0.
		CRect rc(0, 0, 1, 1);
		if(containedComboBox.IsWindow()) {
			containedComboBox.GetWindowRect(&rc);
		}
		if(rc.Width() == 0) {
			GetWindowRect(&rc);
			GetParent().ScreenToClient(&rc);
			MoveWindow(rc.left, rc.top, rc.Width() + 1, rc.Height());
			MoveWindow(rc.left, rc.top, rc.Width(), rc.Height());
		}
	}
}

IMEModeConstants DriveComboBox::GetCurrentIMEContextMode(HWND hWnd)
{
	IMEModeConstants imeContextMode = imeNoControl;

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		imeContextMode = imeInherit;
	} else {
		HIMC hIMC = ImmGetContext(hWnd);
		if(hIMC) {
			if(ImmGetOpenStatus(hIMC)) {
				DWORD conversionMode = 0;
				DWORD sentenceMode = 0;
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				if(conversionMode & IME_CMODE_NATIVE) {
					if(conversionMode & IME_CMODE_KATAKANA) {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeKatakanaHalf];
						} else {
							imeContextMode = countryTable[imeAlphaFull];
						}
					} else {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeHiragana];
						} else {
							imeContextMode = countryTable[imeKatakana];
						}
					}
				} else {
					if(conversionMode & IME_CMODE_FULLSHAPE) {
						imeContextMode = countryTable[imeAlpha];
					} else {
						imeContextMode = countryTable[imeHangulFull];
					}
				}
			} else {
				imeContextMode = countryTable[imeDisable];
			}
			ImmReleaseContext(hWnd, hIMC);
		} else {
			imeContextMode = countryTable[imeOn];
		}
	}
	return imeContextMode;
}

void DriveComboBox::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
{
	if((IMEMode == imeInherit) || (IMEMode == imeNoControl) || !::IsWindow(hWnd)) {
		return;
	}

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		return;
	}

	// update IME mode
	HIMC hIMC = ImmGetContext(hWnd);
	if(IMEMode == imeDisable) {
		// disable IME
		if(hIMC) {
			// close the IME
			if(ImmGetOpenStatus(hIMC)) {
				ImmSetOpenStatus(hIMC, FALSE);
			}
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
		// remove the control's association to the IME context
		HIMC h = ImmAssociateContext(hWnd, NULL);
		if(h) {
			IMEFlags.hDefaultIMC = h;
		}
		return;
	} else {
		// enable IME
		if(!hIMC) {
			if(!IMEFlags.hDefaultIMC) {
				// create an IME context
				hIMC = ImmCreateContext();
				if(hIMC) {
					// associate the control with the IME context
					ImmAssociateContext(hWnd, hIMC);
				}
			} else {
				// associate the control with the default IME context
				ImmAssociateContext(hWnd, IMEFlags.hDefaultIMC);
			}
		} else {
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
	}

	hIMC = ImmGetContext(hWnd);
	if(hIMC) {
		DWORD conversionMode = 0;
		DWORD sentenceMode = 0;
		switch(IMEMode) {
			case imeOn:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				break;
			case imeOff:
				// close IME
				ImmSetOpenStatus(hIMC, FALSE);
				break;
			default:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				// switch conversion
				switch(IMEMode) {
					case imeHiragana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_KATAKANA;
						break;
					case imeKatakana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeKatakanaHalf:
						conversionMode |= (IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
					case imeAlphaFull:
						conversionMode |= IME_CMODE_FULLSHAPE;
						conversionMode &= ~(IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeAlpha:
						conversionMode |= IME_CMODE_ALPHANUMERIC;
						conversionMode &= ~(IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeHangulFull:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeHangul:
						conversionMode |= IME_CMODE_NATIVE;
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
				}
				ImmSetConversionStatus(hIMC, conversionMode, sentenceMode);
				break;
		}
		// each ImmGetContext() needs a ImmReleaseContext()
		ImmReleaseContext(hWnd, hIMC);
		hIMC = NULL;
	}
}

IMEModeConstants DriveComboBox::GetEffectiveIMEMode(void)
{
	IMEModeConstants IMEMode = properties.IMEMode;
	if((IMEMode == imeInherit) && IsWindow()) {
		CWindow wnd = GetParent();
		while((IMEMode == imeInherit) && wnd.IsWindow()) {
			// retrieve the parent's IME mode
			IMEMode = GetCurrentIMEContextMode(wnd);
			wnd = wnd.GetParent();
		}
	}

	if(IMEMode == imeInherit) {
		// use imeNoControl as fallback
		IMEMode = imeNoControl;
	}
	return IMEMode;
}

DWORD DriveComboBox::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	/*switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}*/
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD DriveComboBox::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE | WS_VSCROLL;
	/*switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}*/
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(properties.doOEMConversion) {
		style |= CBS_OEMCONVERT;
	}
	/*if(!properties.integralHeight) {
		style |= CBS_NOINTEGRALHEIGHT;
	}*/
	if(properties.listAlwaysShowVerticalScrollBar) {
		style |= CBS_DISABLENOSCROLL;
	}
	style |= CBS_DROPDOWNLIST;
	return style;
}

void DriveComboBox::SendConfigurationMessages(void)
{
	DWORD extendedStyle = 0;
	if(properties.caseSensitiveItemSearching) {
		extendedStyle |= CBES_EX_CASESENSITIVE;
	}
	switch(properties.iconVisibility) {
		case ivHiddenButIndent:
			extendedStyle |= CBES_EX_NOEDITIMAGE;
			break;
		case ivHiddenDontIndent:
			extendedStyle |= CBES_EX_NOEDITIMAGEINDENT;
			break;
	}
	if(IsComctl32Version610OrNewer()) {
		if(properties.textEndEllipsis) {
			extendedStyle |= CBES_EX_TEXTENDELLIPSIS;
		}
	}
	SendMessage(CBEM_SETEXTENDEDSTYLE, 0, extendedStyle);

	/*// ComboBoxEx32 fiddles around with the borders and edges, so enforce our own styles here
	switch(properties.appearance) {
		case a2D:
			ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
		case a3D:
			ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
		case a3DLight:
			ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
	}
	switch(properties.borderStyle) {
		case bsNone:
			ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
		case bsFixedSingle:
			ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
	}
	SendMessage(WM_WINDOWPOSCHANGING, 0, 0);*/

	SetWindowText(COLE2CT(properties.text));
	// NOTE: CB_SETHORIZONTALEXTENT isn't forwarded by ComboBoxEx32
	//containedComboBox.SendMessage(CB_SETHORIZONTALEXTENT, properties.listScrollableWidth, 0);
	SendMessage(CB_SETDROPPEDWIDTH, properties.listWidth, 0);
	switch(properties.dropDownKey) {
		case ddkF4:
			SendMessage(CB_SETEXTENDEDUI, FALSE, 0);
			break;
		case ddkDownArrow:
			SendMessage(CB_SETEXTENDEDUI, TRUE, 0);
			break;
	}

	ApplyFont();

	CRect rc;
	containedComboBox.GetWindowRect(&rc);
	if(properties.listHeight == -1) {
		// make room for 8 items
		int itemHeight = properties.itemHeight;
		if(itemHeight == -1) {
			itemHeight = static_cast<int>(SendMessage(CB_GETITEMHEIGHT, 0, 0));
		}
		rc.bottom += 8 * itemHeight + 2 * GetSystemMetrics(SM_CYBORDER);
	} else {
		rc.bottom += properties.listHeight;
	}
	ScreenToClient(&rc);
	containedComboBox.MoveWindow(&rc, FALSE);
	if(properties.setItemHeight != -1) {
		SendMessage(CB_SETITEMHEIGHT, 0, properties.setItemHeight);
	}
	if(properties.setSelectionFieldHeight != -1) {
		SendMessage(CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), properties.setSelectionFieldHeight);
	}
	/*if(RunTimeHelper::IsCommCtrl6()) {
		containedComboBox.SendMessage(CB_SETMINVISIBLE, properties.minVisibleItems, 0);
	}*/

	SetSystemImageList();
	ReloadDrives();
}

HCURSOR DriveComboBox::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

int DriveComboBox::IDToItemIndex(LONG ID)
{
	if(ID == 0) {
		// selection field item
		return -1;
	}
	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(itemIDs.begin(), itemIDs.end(), ID);
		if(iter != itemIDs.end()) {
			return std::distance(itemIDs.begin(), iter);
		}
	#else
		for(size_t i = 0; i < itemIDs.GetCount(); ++i) {
			if(itemIDs[i] == ID) {
				return i;
			}
		}
	#endif
	return -2;
}

LONG DriveComboBox::ItemIndexToID(int itemIndex)
{
	if(itemIndex == -1) {
		// selection field item
		return 0;
	}
	#ifdef USE_STL
		if(itemIndex >= 0 && itemIndex < static_cast<int>(itemIDs.size())) {
			return itemIDs[itemIndex];
		}
	#else
		if(itemIndex >= 0 && itemIndex < static_cast<int>(itemIDs.GetCount())) {
			return itemIDs[itemIndex];
		}
	#endif
	return -1;
}

LONG DriveComboBox::GetNewItemID(void)
{
	return ++lastItemID;
}

int DriveComboBox::GetDriveIndex(int itemIndex)
{
	COMBOBOXEXITEM item = {0};
	item.iItem = itemIndex;
	item.mask = CBEIF_LPARAM | CBEIF_NOINTERCEPTION;
	if(SendMessage(CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return item.lParam - 1;
	}
	return -1;
}

PIDLIST_ABSOLUTE DriveComboBox::GetItemPIDL(int itemIndex)
{
	int driveIndex = GetDriveIndex(itemIndex);
	if(driveIndex >= 0) {
		// now the path
		TCHAR pDrive[4] = TEXT("A:\\");
		pDrive[0] += static_cast<TCHAR>(driveIndex);
		// now use the path to get the pIDL
		return ILCreateFromPath(pDrive);
	}
	return NULL;
}

void DriveComboBox::RegisterItemContainer(IItemContainer* pContainer)
{
	ATLASSUME(pContainer);
	#ifdef DEBUG
		#ifdef USE_STL
			std::unordered_map<DWORD, IItemContainer*>::iterator iter = itemContainers.find(pContainer->GetID());
			ATLASSERT(iter == itemContainers.end());
		#else
			CAtlMap<DWORD, IItemContainer*>::CPair* pEntry = itemContainers.Lookup(pContainer->GetID());
			ATLASSERT(pEntry == NULL);
		#endif
	#endif
	itemContainers[pContainer->GetID()] = pContainer;
}

void DriveComboBox::DeregisterItemContainer(DWORD containerID)
{
	#ifdef USE_STL
		std::unordered_map<DWORD, IItemContainer*>::iterator iter = itemContainers.find(containerID);
		ATLASSERT(iter != itemContainers.end());
		if(iter != itemContainers.end()) {
			itemContainers.erase(iter);
		}
	#else
		itemContainers.RemoveKey(containerID);
	#endif
}

void DriveComboBox::RemoveItemFromItemContainers(LONG itemIdentifier)
{
	#ifdef USE_STL
		for(std::unordered_map<DWORD, IItemContainer*>::const_iterator iter = itemContainers.begin(); iter != itemContainers.end(); ++iter) {
			iter->second->RemovedItem(itemIdentifier);
		}
	#else
		POSITION p = itemContainers.GetStartPosition();
		while(p) {
			itemContainers.GetValueAt(p)->RemovedItem(itemIdentifier);
			itemContainers.GetNextValue(p);
		}
	#endif
}

void DriveComboBox::UpdateItemIDInItemContainers(LONG oldItemID, LONG newItemID)
{
	#ifdef USE_STL
		for(std::unordered_map<DWORD, IItemContainer*>::const_iterator iter = itemContainers.begin(); iter != itemContainers.end(); ++iter) {
			iter->second->ReplacedItemID(oldItemID, newItemID);
		}
	#else
		POSITION p = itemContainers.GetStartPosition();
		while(p) {
			itemContainers.GetValueAt(p)->ReplacedItemID(oldItemID, newItemID);
			itemContainers.GetNextValue(p);
		}
	#endif
}

int DriveComboBox::ListBoxHitTest(LONG x, LONG y, HitTestConstants* pFlags, BOOL autoScroll/* = FALSE*/)
{
	ATLASSERT(containedListBox.IsWindow());

	int itemIndex = -1;

	POINT pt = {x, y};
	containedListBox.ClientToScreen(&pt);
	CRect rc;
	containedListBox.GetWindowRect(&rc);
	if(rc.PtInRect(pt)) {
		*pFlags = htNotOverItem;
		itemIndex = LBItemFromPt(containedListBox, pt, autoScroll);
		if(itemIndex != -1) {
			*pFlags = htItem;
		}
	} else {
		if(autoScroll) {
			LBItemFromPt(containedListBox, pt, TRUE);
		}
		*pFlags = static_cast<HitTestConstants>(0);
		if(pt.x < rc.left) {
			*pFlags = static_cast<HitTestConstants>(*pFlags | htToLeft);
		} else if(pt.x >= rc.right) {
			*pFlags = static_cast<HitTestConstants>(*pFlags | htToRight);
		}
		if(pt.y < rc.top) {
			*pFlags = static_cast<HitTestConstants>(*pFlags | htAbove);
		} else if(pt.y >= rc.bottom) {
			*pFlags = static_cast<HitTestConstants>(*pFlags | htBelow);
		}
	}
	return itemIndex;
}

BOOL DriveComboBox::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void DriveComboBox::SetSystemImageList(void)
{
	if(properties.useSystemImageList == usilNone || !IsWindow()) {
		return;
	}

	HIMAGELIST hItemImageList = NULL;
	HIMAGELIST hHighResImageList = NULL;
	if(APIWrapper::IsSupported_SHGetImageList()) {
		int imageList = SHIL_SMALL;
		switch(properties.useSystemImageList) {
			/*case usilSmallImageList:
				imageList = SHIL_SMALL;
				break;*/
			case usilLargeImageList:
				imageList = SHIL_LARGE;
				break;
			case usilExtraLargeImageList:
				imageList = SHIL_EXTRALARGE;
				break;
			case usilJumboImageList:
				imageList = (RunTimeHelper::IsVista() ? SHIL_JUMBO : SHIL_EXTRALARGE);
				break;
		}
		APIWrapper::SHGetImageList(imageList, IID_IImageList, reinterpret_cast<LPVOID*>(&hItemImageList), NULL);
		imageList = (RunTimeHelper::IsVista() ? SHIL_JUMBO : SHIL_EXTRALARGE);
		APIWrapper::SHGetImageList(imageList, IID_IImageList, reinterpret_cast<LPVOID*>(&hHighResImageList), NULL);
	} else {
		UINT imageList = SHGFI_SMALLICON;
		switch(properties.useSystemImageList) {
			/*case usilSmallImageList:
				imageList = SHGFI_SMALLICON;
				break;*/
			case usilLargeImageList:
			case usilExtraLargeImageList:
			case usilJumboImageList:
				imageList = SHGFI_LARGEICON;
				break;
		}
		SHFILEINFO details = {0};
		hItemImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES | imageList));
		hHighResImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES | SHGFI_LARGEICON));
	}

	SendMessage(CBEM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(hItemImageList));
	properties.hHighResImageList = hHighResImageList;
	FireViewChange();
}

void DriveComboBox::ListBoxAutoScroll(void)
{
	LONG realScrollTimeBase = properties.listDragScrollTimeBase;
	if(realScrollTimeBase == -1) {
		realScrollTimeBase = GetDoubleClickTime() / 4;
	}

	if((dragDropStatus.autoScrolling.currentHScrollVelocity != 0) && ((containedListBox.GetStyle() & WS_HSCROLL) == WS_HSCROLL)) {
		if(dragDropStatus.autoScrolling.currentHScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll to the left?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Left) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				containedListBox.GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll left
					dragDropStatus.autoScrolling.lastScroll_Left = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					containedListBox.SendMessage(WM_HSCROLL, SB_LINELEFT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll to the right?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Right) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				containedListBox.GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll right
					dragDropStatus.autoScrolling.lastScroll_Right = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					containedListBox.SendMessage(WM_HSCROLL, SB_LINERIGHT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}

	if((dragDropStatus.autoScrolling.currentVScrollVelocity != 0) && ((containedListBox.GetStyle() & WS_VSCROLL) == WS_VSCROLL)) {
		if(dragDropStatus.autoScrolling.currentVScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll upwardly?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Up) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				containedListBox.GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll up
					dragDropStatus.autoScrolling.lastScroll_Up = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					containedListBox.SendMessage(WM_VSCROLL, SB_LINEUP, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll downwards?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Down) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				containedListBox.GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll down
					dragDropStatus.autoScrolling.lastScroll_Down = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					containedListBox.SendMessage(WM_VSCROLL, SB_LINEDOWN, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}
}

BOOL DriveComboBox::IsLeftMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	}
}

BOOL DriveComboBox::IsRightMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	}
}


HRESULT DriveComboBox::CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing)
{
	if(!hIcon || !pBits || !pWICImagingFactory) {
		return E_FAIL;
	}

	ICONINFO iconInfo;
	GetIconInfo(hIcon, &iconInfo);
	ATLASSERT(iconInfo.hbmColor);
	BITMAP bitmapInfo = {0};
	if(iconInfo.hbmColor) {
		GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bitmapInfo);
	} else if(iconInfo.hbmMask) {
		GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bitmapInfo);
	}
	bitmapInfo.bmHeight = abs(bitmapInfo.bmHeight);
	BOOL needsFrame = ((bitmapInfo.bmWidth < size.cx) || (bitmapInfo.bmHeight < size.cy));
	if(iconInfo.hbmColor) {
		DeleteObject(iconInfo.hbmColor);
	}
	if(iconInfo.hbmMask) {
		DeleteObject(iconInfo.hbmMask);
	}

	HRESULT hr = E_FAIL;

	CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
	if(!needsFrame) {
		hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapScaler);
	}
	if(needsFrame || SUCCEEDED(hr)) {
		CComPtr<IWICBitmap> pWICBitmapSource = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmapSource);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapSource);
		if(SUCCEEDED(hr)) {
			if(!needsFrame) {
				hr = pWICBitmapScaler->Initialize(pWICBitmapSource, size.cx, size.cy, WICBitmapInterpolationModeFant);
			}
			if(SUCCEEDED(hr)) {
				WICRect rc = {0};
				if(needsFrame) {
					rc.Height = bitmapInfo.bmHeight;
					rc.Width = bitmapInfo.bmWidth;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					LPRGBQUAD pIconBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, rc.Width * rc.Height * sizeof(RGBQUAD)));
					hr = pWICBitmapSource->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pIconBits));
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						// center the icon
						int xIconStart = (size.cx - bitmapInfo.bmWidth) / 2;
						int yIconStart = (size.cy - bitmapInfo.bmHeight) / 2;
						LPRGBQUAD pIconPixel = pIconBits;
						LPRGBQUAD pPixel = pBits;
						pPixel += yIconStart * size.cx;
						for(int y = yIconStart; y < yIconStart + bitmapInfo.bmHeight; ++y, pPixel += size.cx, pIconPixel += bitmapInfo.bmWidth) {
							CopyMemory(pPixel + xIconStart, pIconPixel, bitmapInfo.bmWidth * sizeof(RGBQUAD));
						}
						HeapFree(GetProcessHeap(), 0, pIconBits);

						rc.Height = size.cy;
						rc.Width = size.cx;

						// TODO: now draw a frame around it
					}
				} else {
					rc.Height = size.cy;
					rc.Width = size.cx;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					hr = pWICBitmapScaler->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pBits));
					ATLASSERT(SUCCEEDED(hr));

					if(SUCCEEDED(hr) && doAlphaChannelPostProcessing) {
						for(int i = 0; i < rc.Width * rc.Height; ++i, ++pBits) {
							if(pBits->rgbReserved == 0x00) {
								ZeroMemory(pBits, sizeof(RGBQUAD));
							}
						}
					}
				}
			} else {
				ATLASSERT(FALSE && "Bitmap scaler failed");
			}
		}
	}
	return hr;
}


BOOL DriveComboBox::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}