// ListBox.cpp: Superclasses ListBox.

#include "stdafx.h"
#include "ListBox.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
IMEModeConstants ListBox::IMEFlags::chineseIMETable[10] = {
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

IMEModeConstants ListBox::IMEFlags::japaneseIMETable[10] = {
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

IMEModeConstants ListBox::IMEFlags::koreanIMETable[10] = {
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

FONTDESC ListBox::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


ListBox::ListBox()
{
	properties.font.InitializePropertyWatcher(this, DISPID_LB_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_LB_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
	lastItemID = -1;
	currentCaretItem = -1;
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


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListBox::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListBox, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ListBox::GetSizeMax(ULARGE_INTEGER* pSize)
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
	// ...and 15 VT_BOOL properties...
	pSize->QuadPart += 15 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));

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

STDMETHODIMP ListBox::Load(LPSTREAM pStream)
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
	if(version > 0x0100) {
		return E_FAIL;
	}
	LONG subSignature = 0;
	if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}
	if(subSignature != 0x04040404/*4x 0x04 (-> ListBox)*/) {
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

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowItemSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL alwaysShowVerticalScrollBar = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG columnWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragScrollTimeBase = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	VARTYPE vt;
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL hasStrings = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants IMEMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR insertMarkColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	InsertMarkStyleConstants insertMarkStyle = static_cast<InsertMarkStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL integralHeight = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS itemHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG locale = propertyValue.lVal;

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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL multiColumn = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MultiSelectConstants multiSelect = static_cast<MultiSelectConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLEDragImageStyleConstants oleDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OwnerDrawItemsConstants ownerDrawItems = static_cast<OwnerDrawItemsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processTabs = propertyValue.boolVal;
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
	OLE_XSIZE_PIXELS scrollableWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL sorted = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS tabWidthInPixels = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ToolTipsConstants toolTips = static_cast<ToolTipsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL virtualMode = propertyValue.boolVal;


	hr = put_AllowDragDrop(allowDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ColumnWidth(columnWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragScrollTimeBase(dragScrollTimeBase);
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
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IMEMode(IMEMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InsertMarkColor(insertMarkColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InsertMarkStyle(insertMarkStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemHeight(itemHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Locale(locale);
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
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollableWidth(scrollableWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TabWidth(tabWidthInPixels);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ToolTips(toolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	BOOL b = flags.dontRecreate;
	flags.dontRecreate = TRUE;
	hr = put_AllowItemSelection(allowItemSelection);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AlwaysShowVerticalScrollBar(alwaysShowVerticalScrollBar);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HasStrings(hasStrings);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IntegralHeight(integralHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MultiColumn(multiColumn);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MultiSelect(multiSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OwnerDrawItems(ownerDrawItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessTabs(processTabs);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Sorted(sorted);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_VirtualMode(virtualMode);
	if(FAILED(hr)) {
		return hr;
	}
	flags.dontRecreate = b;
	RecreateControlWindow();

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ListBox::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
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
	LONG version = 0x0100;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x04040404/*4x 0x04 (-> ListBox)*/;
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

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowItemSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.alwaysShowVerticalScrollBar);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.columnWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.dragScrollTimeBase;
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
	VARTYPE vt = VT_DISPATCH;
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
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.hasStrings);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.IMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.insertMarkColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.insertMarkStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.integralHeight);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.setItemHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.locale;
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
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiColumn);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.multiSelect;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.oleDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.ownerDrawItems;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processTabs);
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
	propertyValue.lVal = properties.scrollableWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.sorted);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.tabWidthInPixels;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.toolTips;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.virtualMode);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND ListBox::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_STANDARD_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<ListBox>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT ListBox::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("ListBox ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void ListBox::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP ListBox::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<ListBox>::IOleObject_SetClientSite(pClientSite);

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

	/*LCID lcid = LOCALE_USER_DEFAULT;
	GetAmbientLocaleID(lcid);
	put_Locale(lcid);*/

	return hr;
}

STDMETHODIMP ListBox::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL ListBox::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
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
	return CComControl<ListBox>::PreTranslateAccelerator(pMessage, hReturnValue);
}

HIMAGELIST ListBox::CreateLegacyDragImage(int itemIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	/********************************************************************************************************
	 * Known problems:                                                                                      *
	 * - We use hardcoded margins.                                                                          *
	 * - Owner-drawn items are not supported.                                                               *
	 ********************************************************************************************************/

	// retrieve window details
	BOOL hasFocus = (GetFocus() == *this);
	DWORD style = GetExStyle();
	DWORD textDrawStyle = DT_EDITCONTROL | DT_NOPREFIX | DT_SINGLELINE | DT_VCENTER;
	if(style & WS_EX_RTLREADING) {
		textDrawStyle |= DT_RTLREADING;
	}
	BOOL layoutRTL = ((style & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
	BOOL containsStrings;
	style = GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		containsStrings = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		containsStrings = TRUE;
	}

	// retrieve item details
	BOOL itemIsSelected = SendMessage(LB_GETSEL, itemIndex, 0);
	BOOL itemIsFocused = (itemIndex == SendMessage(LB_GETCARETINDEX, 0, 0));
	int itemTextLength = 0;
	LPTSTR pItemText = NULL;
	if(containsStrings) {
		itemTextLength = static_cast<int>(SendMessage(LB_GETTEXTLEN, itemIndex, 0));
		if(itemTextLength == LB_ERR) {
			return NULL;
		}
		pItemText = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (itemTextLength + 1) * sizeof(TCHAR)));
		if(!pItemText) {
			return NULL;
		}
		SendMessage(LB_GETTEXT, itemIndex, reinterpret_cast<LPARAM>(pItemText));
	}

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
	CRect focusRect;
	CRect labelBoundingRect;

	SendMessage(LB_GETITEMRECT, itemIndex, reinterpret_cast<LPARAM>(&selectionBoundingRect));
	labelBoundingRect = selectionBoundingRect;
	itemBoundingRect = selectionBoundingRect;
	// center the label rectangle vertically
	int cy = labelBoundingRect.Height();
	labelBoundingRect.top = itemBoundingRect.top + (itemBoundingRect.Height() - cy) / 2;
	labelBoundingRect.bottom = labelBoundingRect.top + cy;
	itemBoundingRect.UnionRect(&itemBoundingRect, &labelBoundingRect);
	focusRect = selectionBoundingRect;
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

	if(itemTextLength > 0) {
		// draw the text
		CRect rc = labelBoundingRect;
		if(layoutRTL) {
			// TODO: Don't use hard-coded margins
			rc.OffsetRect(1, 0);
		}
		// TODO: Don't use hard-coded margins
		rc.InflateRect(-2, 0);
		if(themedListItems) {
			CT2W converter(pItemText);
			LPWSTR pLabelText = converter;
			themingEngine.DrawThemeText(memoryDC, LBCP_ITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle, 0, &rc);
		} else {
			memoryDC.DrawText(pItemText, itemTextLength, &rc, textDrawStyle);
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
								if((pPixel->rgbReserved != 0x00) || selectionBoundingRect.PtInRect(pt2)) {
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
								if((pPixel->rgbReserved != 0x00) || selectionBoundingRect.PtInRect(pt)) {
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

	if(pItemText) {
		HeapFree(GetProcessHeap(), 0, pItemText);
	}
	ReleaseDC(hCompatibleDC);

	return hDragImageList;
}

BOOL ListBox::CreateLegacyOLEDragImage(IListBoxItemContainer* pItems, LPSHDRAGIMAGE pDragImage)
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

BOOL ListBox::CreateVistaOLEDragImage(IListBoxItemContainer* pItems, LPSHDRAGIMAGE pDragImage)
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
				CComPtr<IListBoxItem> pItem = NULL;
				while(pEnum->Next(1, &v, NULL) == S_OK) {
					if(v.vt == VT_DISPATCH) {
						v.pdispVal->QueryInterface(IID_IListBoxItem, reinterpret_cast<LPVOID*>(&pItem));
						ATLASSUME(pItem);
						if(pItem) {
							// get the item's icon
							LONG icon = 0;
							LONG overlay = 0;
							Raise_ItemGetDisplayInfo(pItem, riIconIndex, &icon, &overlay);

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

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP ListBox::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
		if(dragDropStatus.useItemCountLabelHack) {
			dragDropStatus.pDropTargetHelper->DragLeave();
			dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
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

STDMETHODIMP ListBox::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP ListBox::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ListBox::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
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
STDMETHODIMP ListBox::GiveFeedback(DWORD effect)
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

STDMETHODIMP ListBox::QueryContinueDrag(BOOL pressedEscape, DWORD keyState)
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
STDMETHODIMP ListBox::DragEnterTarget(HWND hWndTarget)
{
	Raise_OLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	return S_OK;
}

STDMETHODIMP ListBox::DragLeaveTarget(void)
{
	Raise_OLEDragLeavePotentialTarget();
	return S_OK;
}
// implementation of IDropSourceNotify
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP ListBox::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
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

STDMETHODIMP ListBox::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_LB_APPEARANCE:
		case DISPID_LB_BORDERSTYLE:
		case DISPID_LB_COLUMNWIDTH:
		case DISPID_LB_ITEMHEIGHT:
		case DISPID_LB_MOUSEICON:
		case DISPID_LB_MOUSEPOINTER:
		case DISPID_LB_PROCESSTABS:
		case DISPID_LB_TABSTOPS:
		case DISPID_LB_TABWIDTH:
		case DISPID_LB_VIRTUALITEMCOUNT:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_LB_ALLOWITEMSELECTION:
		case DISPID_LB_ALWAYSSHOWVERTICALSCROLLBAR:
		case DISPID_LB_ANCHORITEM:
		case DISPID_LB_CARETITEM:
		case DISPID_LB_DISABLEDEVENTS:
		case DISPID_LB_DONTREDRAW:
		case DISPID_LB_HOVERTIME:
		case DISPID_LB_IMEMODE:
		case DISPID_LB_MULTICOLUMN:
		case DISPID_LB_MULTISELECT:
		case DISPID_LB_OWNERDRAWITEMS:
		case DISPID_LB_PROCESSCONTEXTMENUKEYS:
		case DISPID_LB_RIGHTTOLEFT:
		case DISPID_LB_SCROLLABLEWIDTH:
		case DISPID_LB_SORTED:
		case DISPID_LB_TOOLTIPS:
		case DISPID_LB_VIRTUALMODE:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_LB_BACKCOLOR:
		case DISPID_LB_FORECOLOR:
		case DISPID_LB_INSERTMARKCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_LB_APPID:
		case DISPID_LB_APPNAME:
		case DISPID_LB_APPSHORTNAME:
		case DISPID_LB_BUILD:
		case DISPID_LB_CHARSET:
		case DISPID_LB_HIMAGELIST:
		case DISPID_LB_HWND:
		case DISPID_LB_HWNDTOOLTIP:
		case DISPID_LB_ISRELEASE:
		case DISPID_LB_PROGRAMMER:
		case DISPID_LB_TESTER:
		case DISPID_LB_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_LB_ALLOWDRAGDROP:
		case DISPID_LB_DRAGGEDITEMS:
		case DISPID_LB_DRAGSCROLLTIMEBASE:
		case DISPID_LB_HDRAGIMAGELIST:
		case DISPID_LB_INSERTMARKSTYLE:
		case DISPID_LB_OLEDRAGIMAGESTYLE:
		case DISPID_LB_REGISTERFOROLEDRAGDROP:
		case DISPID_LB_SHOWDRAGIMAGE:
		case DISPID_LB_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_LB_FONT:
		case DISPID_LB_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_LB_FIRSTVISIBLEITEM:
		case DISPID_LB_HASSTRINGS:
		case DISPID_LB_INTEGRALHEIGHT:
		case DISPID_LB_LISTITEMS:
		case DISPID_LB_SELECTEDITEM:
		case DISPID_LB_SELECTEDITEMSARRAY:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_LB_ENABLED:
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
CAtlString ListBox::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ListBox::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ListBox::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ComboListBoxControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ListBox::GetSpecialThanks(void)
{
	return TEXT("Wine Headquarters");
}

CAtlString ListBox::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ListBox::GetVersion(void)
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
// implementation of IPerPropertyBrowsing
STDMETHODIMP ListBox::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_LB_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_LB_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_LB_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_LB_INSERTMARKSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.insertMarkStyle), description);
			break;
		case DISPID_LB_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_LB_MULTISELECT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.multiSelect), description);
			break;
		case DISPID_LB_OLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.oleDragImageStyle), description);
			break;
		case DISPID_LB_OWNERDRAWITEMS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.ownerDrawItems), description);
			break;
		case DISPID_LB_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_LB_TOOLTIPS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.toolTips), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<ListBox>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ListBox::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_LB_BORDERSTYLE:
		case DISPID_LB_INSERTMARKSTYLE:
		case DISPID_LB_OLEDRAGIMAGESTYLE:
			c = 2;
			break;
		case DISPID_LB_APPEARANCE:
		case DISPID_LB_MULTISELECT:
		case DISPID_LB_OWNERDRAWITEMS:
			c = 3;
			break;
		case DISPID_LB_RIGHTTOLEFT:
		case DISPID_LB_TOOLTIPS:
			c = 4;
			break;
		case DISPID_LB_IMEMODE:
			c = 12;
			break;
		case DISPID_LB_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<ListBox>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_LB_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
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

STDMETHODIMP ListBox::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_LB_APPEARANCE:
		case DISPID_LB_BORDERSTYLE:
		case DISPID_LB_IMEMODE:
		case DISPID_LB_INSERTMARKSTYLE:
		case DISPID_LB_MOUSEPOINTER:
		case DISPID_LB_MULTISELECT:
		case DISPID_LB_OLEDRAGIMAGESTYLE:
		case DISPID_LB_OWNERDRAWITEMS:
		case DISPID_LB_RIGHTTOLEFT:
		case DISPID_LB_TOOLTIPS:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<ListBox>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ListBox::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	return IPerPropertyBrowsingImpl<ListBox>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ListBox::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_LB_APPEARANCE:
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
		case DISPID_LB_BORDERSTYLE:
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
			break;
		case DISPID_LB_IMEMODE:
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
		case DISPID_LB_INSERTMARKSTYLE:
			switch(cookie) {
				case imsNative:
					description = GetResStringWithNumber(IDP_INSERTMARKSTYLENATIVE, imsNative);
					break;
				case imsImproved:
					description = GetResStringWithNumber(IDP_INSERTMARKSTYLEIMPROVED, imsImproved);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_LB_MOUSEPOINTER:
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
		case DISPID_LB_MULTISELECT:
			switch(cookie) {
				case msNone:
					description = GetResStringWithNumber(IDP_MULTISELECTNONE, msNone);
					break;
				case msNormal:
					description = GetResStringWithNumber(IDP_MULTISELECTNORMAL, msNormal);
					break;
				case msSelectByClick:
					description = GetResStringWithNumber(IDP_MULTISELECTSELECTBYCLICK, msSelectByClick);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_LB_OLEDRAGIMAGESTYLE:
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
		case DISPID_LB_OWNERDRAWITEMS:
			switch(cookie) {
				case odiDontOwnerDraw:
					description = GetResStringWithNumber(IDP_OWNERDRAWITEMSDONTOWNERDRAW, odiDontOwnerDraw);
					break;
				case odiOwnerDrawFixedHeight:
					description = GetResStringWithNumber(IDP_OWNERDRAWITEMSFIXEDHEIGHT, odiOwnerDrawFixedHeight);
					break;
				case odiOwnerDrawVariableHeight:
					description = GetResStringWithNumber(IDP_OWNERDRAWITEMSVARIABLEHEIGHT, odiOwnerDrawVariableHeight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_LB_RIGHTTOLEFT:
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
		case DISPID_LB_TOOLTIPS:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_TOOLTIPSNONE, 0);
					break;
				case ttLabelTips:
					description = GetResStringWithNumber(IDP_TOOLTIPSLABELTIPS, ttLabelTips);
					break;
				case ttInfoTips:
					description = GetResStringWithNumber(IDP_TOOLTIPSINFOTIPS, ttInfoTips);
					break;
				case ttLabelTips | ttInfoTips:
					description = GetResStringWithNumber(IDP_TOOLTIPSLABELANDINFOTIPS, ttLabelTips | ttInfoTips);
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
STDMETHODIMP ListBox::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 4;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StockColorPage;
		pPropertyPages->pElems[2] = CLSID_StockFontPage;
		pPropertyPages->pElems[3] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP ListBox::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ListBox>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ListBox::EnumVerbs(IEnumOLEVERB** ppEnumerator)
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
STDMETHODIMP ListBox::GetControlInfo(LPCONTROLINFO pControlInfo)
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

HRESULT ListBox::DoVerbAbout(HWND hWndParent)
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

HRESULT ListBox::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_LB_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT ListBox::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP ListBox::get_AllowDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowDragDrop);
	return S_OK;
}

STDMETHODIMP ListBox::put_AllowDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_ALLOWDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowDragDrop != b) {
		properties.allowDragDrop = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_LB_ALLOWDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_AllowItemSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowItemSelection = !(GetStyle() & LBS_NOSEL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowItemSelection);
	return S_OK;
}

STDMETHODIMP ListBox::put_AllowItemSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_ALLOWITEMSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowItemSelection != b) {
		properties.allowItemSelection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_ALLOWITEMSELECTION);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_AlwaysShowVerticalScrollBar(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.alwaysShowVerticalScrollBar = ((GetStyle() & LBS_DISABLENOSCROLL) == LBS_DISABLENOSCROLL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.alwaysShowVerticalScrollBar);
	return S_OK;
}

STDMETHODIMP ListBox::put_AlwaysShowVerticalScrollBar(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_ALWAYSSHOWVERTICALSCROLLBAR);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.alwaysShowVerticalScrollBar != b) {
		properties.alwaysShowVerticalScrollBar = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_ALWAYSSHOWVERTICALSCROLLBAR);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_AnchorItem(IListBoxItem** ppAnchorItem)
{
	ATLASSERT_POINTER(ppAnchorItem, IListBoxItem*);
	if(!ppAnchorItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		ClassFactory::InitListItem(static_cast<int>(SendMessage(LB_GETANCHORINDEX, 0, 0)), this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppAnchorItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::putref_AnchorItem(IListBoxItem* pNewAnchorItem)
{
	PUTPROPPROLOG(DISPID_LB_ANCHORITEM);
	int newAnchorItem = -1;
	if(pNewAnchorItem) {
		LONG l = -1;
		pNewAnchorItem->get_Index(&l);
		newAnchorItem = l;
		// TODO: Shouldn't we AddRef' pNewAnchorItem?
	}

	if(IsWindow()) {
		SendMessage(LB_SETANCHORINDEX, newAnchorItem, 0);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_LB_ANCHORITEM);
	return S_OK;
}

STDMETHODIMP ListBox::get_Appearance(AppearanceConstants* pValue)
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

STDMETHODIMP ListBox::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_APPEARANCE);
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
			FireViewChange();
		}
		FireOnChanged(DISPID_LB_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 16;
	return S_OK;
}

STDMETHODIMP ListBox::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ComboListBoxControls");
	return S_OK;
}

STDMETHODIMP ListBox::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"CBLCtls");
	return S_OK;
}

STDMETHODIMP ListBox::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP ListBox::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_LB_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_LB_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_BorderStyle(BorderStyleConstants* pValue)
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

STDMETHODIMP ListBox::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_BORDERSTYLE);
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
			FireViewChange();
		}
		FireOnChanged(DISPID_LB_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ListBox::get_CaretItem(VARIANT_BOOL /*partialVisibilityIsOkay = VARIANT_FALSE*/, IListBoxItem** ppCaretItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppCaretItem, IListBoxItem*);
	if(!ppCaretItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		ClassFactory::InitListItem(static_cast<int>(SendMessage(LB_GETCARETINDEX, 0, 0)), this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppCaretItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::putref_CaretItem(VARIANT_BOOL partialVisibilityIsOkay/* = VARIANT_FALSE*/, IListBoxItem* pNewCaretItem/* = NULL*/)
{
	PUTPROPPROLOG(DISPID_LB_CARETITEM);
	HRESULT hr = E_FAIL;

	int newCaretItem = -1;
	if(pNewCaretItem) {
		LONG l = -1;
		pNewCaretItem->get_Index(&l);
		newCaretItem = l;
		// TODO: Shouldn't we AddRef' pNewCaretItem?
	}

	if(IsWindow()) {
		if(SendMessage(LB_SETCARETINDEX, newCaretItem, VARIANTBOOL2BOOL(partialVisibilityIsOkay)) != LB_ERR) {
			hr = S_OK;
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_LB_CARETITEM);
	FireViewChange();
	return hr;
}

STDMETHODIMP ListBox::get_CharSet(BSTR* pValue)
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

STDMETHODIMP ListBox::get_ColumnWidth(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.columnWidth;
	return S_OK;
}

STDMETHODIMP ListBox::put_ColumnWidth(LONG newValue)
{
	PUTPROPPROLOG(DISPID_LB_COLUMNWIDTH);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.columnWidth != newValue) {
		properties.columnWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LB_SETCOLUMNWIDTH, properties.columnWidth, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_LB_COLUMNWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents & ~deProcessKeyboardInput);
		if(GetStyle() & LBS_WANTKEYBOARDINPUT) {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents | deProcessKeyboardInput);
		}
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ListBox::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deMouseEvents) != (newValue & deMouseEvents)) {
			if(IsWindow()) {
				if(newValue & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
				itemUnderMouse = -1;
			}
		}
		BOOL needsRecreation = FALSE;
		if((properties.disabledEvents & deProcessKeyboardInput) != (newValue & deProcessKeyboardInput)) {
			needsRecreation = IsWindow();
		}

		properties.disabledEvents = newValue;
		if(needsRecreation) {
			RecreateControlWindow();
		}
		SetDirty(TRUE);
		FireOnChanged(DISPID_LB_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// NOTE: We could check for LBS_NOREDRAW here.

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP ListBox::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_LB_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_DraggedItems(IListBoxItemContainer** ppItems)
{
	ATLASSERT_POINTER(ppItems, IListBoxItemContainer*);
	if(!ppItems) {
		return E_POINTER;
	}

	*ppItems = NULL;
	if(dragDropStatus.pDraggedItems) {
		return dragDropStatus.pDraggedItems->Clone(ppItems);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_DragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP ListBox::put_DragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_LB_DRAGSCROLLTIMEBASE);

	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragScrollTimeBase != newValue) {
		properties.dragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_LB_DRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_Enabled(VARIANT_BOOL* pValue)
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

STDMETHODIMP ListBox::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_ENABLED);
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

		FireOnChanged(DISPID_LB_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_FirstVisibleItem(IListBoxItem** ppFirstItem)
{
	ATLASSERT_POINTER(ppFirstItem, IListBoxItem*);
	if(!ppFirstItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		ClassFactory::InitListItem(static_cast<int>(SendMessage(LB_GETTOPINDEX, 0, 0)), this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFirstItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::putref_FirstVisibleItem(IListBoxItem* pNewFirstItem)
{
	PUTPROPPROLOG(DISPID_LB_FIRSTVISIBLEITEM);
	HRESULT hr = E_FAIL;

	int newFirstItem = -1;
	if(pNewFirstItem) {
		LONG l = -1;
		pNewFirstItem->get_Index(&l);
		newFirstItem = l;
		// TODO: Shouldn't we AddRef' pNewFirstItem?
	}

	if(IsWindow()) {
		SendMessage(LB_SETTOPINDEX, newFirstItem, 0);
		hr = S_OK;
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_LB_FIRSTVISIBLEITEM);
	FireViewChange();
	return hr;
}

STDMETHODIMP ListBox::get_Font(IFontDisp** ppFont)
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

STDMETHODIMP ListBox::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_LB_FONT);
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
	FireOnChanged(DISPID_LB_FONT);
	return S_OK;
}

STDMETHODIMP ListBox::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_LB_FONT);
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
	FireOnChanged(DISPID_LB_FONT);
	return S_OK;
}

STDMETHODIMP ListBox::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP ListBox::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_LB_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_LB_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_HasStrings(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.hasStrings = ((GetStyle() & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.hasStrings);
	return S_OK;
}

STDMETHODIMP ListBox::put_HasStrings(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_HASSTRINGS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.hasStrings != b) {
		properties.hasStrings = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_HASSTRINGS);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_hDragImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(dragDropStatus.IsDragging()) {
		*pValue = HandleToLong(dragDropStatus.hDragImageList);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = NULL;
	switch(imageList) {
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

STDMETHODIMP ListBox::put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_LB_HIMAGELIST);
	switch(imageList) {
		case ilHighResolution:
			properties.hHighResImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	FireOnChanged(DISPID_LB_HIMAGELIST);
	return S_OK;
}

STDMETHODIMP ListBox::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP ListBox::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_LB_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_LB_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP ListBox::get_hWndToolTip(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(toolTipStatus.hWndToolTip);
	}
	return S_OK;
}

STDMETHODIMP ListBox::put_hWndToolTip(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_LB_HWNDTOOLTIP);
	toolTipStatus.hWndToolTip = static_cast<HWND>(LongToHandle(newValue));

	SetDirty(TRUE);
	FireOnChanged(DISPID_LB_HWNDTOOLTIP);
	return S_OK;
}

STDMETHODIMP ListBox::get_IMEMode(IMEModeConstants* pValue)
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

STDMETHODIMP ListBox::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == *this) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(*this, ime);
				}
			}
		}
		FireOnChanged(DISPID_LB_IMEMODE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_InsertMarkColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(insertMark.color != OLECOLOR2COLORREF(properties.insertMarkColor)) {
			properties.insertMarkColor = insertMark.color;
		}
	}

	*pValue = properties.insertMarkColor;
	return S_OK;
}

STDMETHODIMP ListBox::put_InsertMarkColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_LB_INSERTMARKCOLOR);
	if(properties.insertMarkColor != newValue) {
		properties.insertMarkColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			insertMark.color = OLECOLOR2COLORREF(properties.insertMarkColor);
			if(!insertMark.hidden) {
				switch(properties.insertMarkStyle) {
					case imsImproved:
					{
						RECT itemBoundingRectangle = {0};
						CRect insertMarkRect;
						if(insertMark.itemIndex != -1) {
							SendMessage(LB_GETITEMRECT, insertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
							if(insertMark.afterItem) {
								insertMarkRect.top = itemBoundingRectangle.bottom - 3;
								insertMarkRect.bottom = itemBoundingRectangle.bottom + 3;
							} else {
								insertMarkRect.top = itemBoundingRectangle.top - 3;
								insertMarkRect.bottom = itemBoundingRectangle.top + 3;
							}
							insertMarkRect.left = itemBoundingRectangle.left;
							insertMarkRect.right = itemBoundingRectangle.right;
						}

						// redraw
						if(insertMarkRect.Width() > 0) {
							InvalidateRect(&insertMarkRect);
						}
						break;
					}
				}
			}
		}
		FireOnChanged(DISPID_LB_INSERTMARKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_InsertMarkStyle(InsertMarkStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, InsertMarkStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.insertMarkStyle;
	return S_OK;
}

STDMETHODIMP ListBox::put_InsertMarkStyle(InsertMarkStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_INSERTMARKSTYLE);
	if(properties.insertMarkStyle != newValue) {
		properties.insertMarkStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow() && !insertMark.hidden) {
			dragDropStatus.HideDragImage(TRUE);
			Invalidate();
			UpdateWindow();
			dragDropStatus.ShowDragImage(TRUE);
		}
		FireOnChanged(DISPID_LB_INSERTMARKSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_IntegralHeight(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.integralHeight = !(GetStyle() & LBS_NOINTEGRALHEIGHT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.integralHeight);
	return S_OK;
}

STDMETHODIMP ListBox::put_IntegralHeight(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_INTEGRALHEIGHT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.integralHeight != b) {
		properties.integralHeight = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_INTEGRALHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_IsRelease(VARIANT_BOOL* pValue)
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

STDMETHODIMP ListBox::get_ItemHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.itemHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(LB_GETITEMHEIGHT, 0, 0));
	}

	*pValue = properties.itemHeight;
	return S_OK;
}

STDMETHODIMP ListBox::put_ItemHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_LB_ITEMHEIGHT);
	if(properties.itemHeight != newValue) {
		properties.itemHeight = newValue;
		properties.setItemHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.itemHeight == -1) {
				ApplyFont();
			} else {
				if(!(GetStyle() & LBS_OWNERDRAWVARIABLE)) {
					SendMessage(LB_SETITEMHEIGHT, 0, properties.itemHeight);
				}
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_LB_ITEMHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_ItemsPerColumn(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<LONG>(GetListBoxInfo(*this));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::get_ListItems(IListBoxItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IListBoxItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	ClassFactory::InitListItems(this, IID_IListBoxItems, reinterpret_cast<LPUNKNOWN*>(ppItems));
	return S_OK;
}

STDMETHODIMP ListBox::get_Locale(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.locale = SendMessage(LB_GETLOCALE, 0, 0);
	}

	*pValue = properties.locale;
	return S_OK;
}

STDMETHODIMP ListBox::put_Locale(LONG newValue)
{
	PUTPROPPROLOG(DISPID_LB_LOCALE);
	if(properties.locale != newValue) {
		properties.locale = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LB_SETLOCALE, properties.locale, 0);
		}
		FireOnChanged(DISPID_LB_LOCALE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_MouseIcon(IPictureDisp** ppMouseIcon)
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

STDMETHODIMP ListBox::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_LB_MOUSEICON);
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
	FireOnChanged(DISPID_LB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ListBox::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_LB_MOUSEICON);
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
	FireOnChanged(DISPID_LB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ListBox::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP ListBox::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_LB_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_MultiColumn(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiColumn = ((GetStyle() & LBS_MULTICOLUMN) == LBS_MULTICOLUMN);
	}

	*pValue = BOOL2VARIANTBOOL(properties.multiColumn);
	return S_OK;
}

STDMETHODIMP ListBox::put_MultiColumn(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_MULTICOLUMN);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiColumn != b) {
		properties.multiColumn = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_MULTICOLUMN);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_MultiSelect(MultiSelectConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MultiSelectConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & LBS_EXTENDEDSEL) {
			properties.multiSelect = msNormal;
		} else if(style & LBS_MULTIPLESEL) {
			properties.multiSelect = msSelectByClick;
		} else {
			properties.multiSelect = msNone;
		}
	}

	*pValue = properties.multiSelect;
	return S_OK;
}

STDMETHODIMP ListBox::put_MultiSelect(MultiSelectConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_MULTISELECT);
	if(properties.multiSelect != newValue) {
		properties.multiSelect = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_MULTISELECT);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.oleDragImageStyle;
	return S_OK;
}

STDMETHODIMP ListBox::put_OLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_OLEDRAGIMAGESTYLE);
	if(properties.oleDragImageStyle != newValue) {
		properties.oleDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_LB_OLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_OwnerDrawItems(OwnerDrawItemsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OwnerDrawItemsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
			case LBS_OWNERDRAWFIXED:
				properties.ownerDrawItems = odiOwnerDrawFixedHeight;
				break;
			case LBS_OWNERDRAWVARIABLE:
				properties.ownerDrawItems = odiOwnerDrawVariableHeight;
				break;
			default:
				properties.ownerDrawItems = odiDontOwnerDraw;
				break;
		}
	}

	*pValue = properties.ownerDrawItems;
	return S_OK;
}

STDMETHODIMP ListBox::put_OwnerDrawItems(OwnerDrawItemsConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_OWNERDRAWITEMS);
	if(properties.ownerDrawItems != newValue) {
		properties.ownerDrawItems = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_OWNERDRAWITEMS);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP ListBox::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_LB_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_ProcessTabs(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.processTabs = ((GetStyle() & LBS_USETABSTOPS) == LBS_USETABSTOPS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.processTabs);
	return S_OK;
}

STDMETHODIMP ListBox::put_ProcessTabs(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_PROCESSTABS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processTabs != b) {
		properties.processTabs = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_PROCESSTABS);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ListBox::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP ListBox::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
		}
		FireOnChanged(DISPID_LB_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_RightToLeft(RightToLeftConstants* pValue)
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

STDMETHODIMP ListBox::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.rightToLeft & rtlLayout) {
				ModifyStyleEx(0, WS_EX_LAYOUTRTL);
			} else {
				ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
			}
			if(properties.rightToLeft & rtlText) {
				ModifyStyleEx(0, WS_EX_RTLREADING);
			} else {
				ModifyStyleEx(WS_EX_RTLREADING, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_LB_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_ScrollableWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.scrollableWidth = static_cast<LONG>(SendMessage(LB_GETHORIZONTALEXTENT, 0, 0));
	}

	*pValue = properties.scrollableWidth;
	return S_OK;
}

STDMETHODIMP ListBox::put_ScrollableWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_LB_SCROLLABLEWIDTH);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.scrollableWidth != newValue) {
		properties.scrollableWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LB_SETHORIZONTALEXTENT, properties.scrollableWidth, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_LB_SCROLLABLEWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_SelectedItem(IListBoxItem** ppSelectedItem)
{
	ATLASSERT_POINTER(ppSelectedItem, IListBoxItem*);
	if(!ppSelectedItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(GetStyle() & (LBS_MULTIPLESEL | LBS_EXTENDEDSEL)) {
			return E_FAIL;
		}
		int itemIndex = static_cast<int>(SendMessage(LB_GETCURSEL, 0, 0));
		if(itemIndex != LB_ERR) {
			ClassFactory::InitListItem(itemIndex, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppSelectedItem));
			return S_OK;
		} else {
			*ppSelectedItem = NULL;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::putref_SelectedItem(IListBoxItem* pNewSelectedItem)
{
	PUTPROPPROLOG(DISPID_LB_SELECTEDITEM);
	int newSelectedItem = -1;
	if(pNewSelectedItem) {
		LONG l = -1;
		pNewSelectedItem->get_Index(&l);
		newSelectedItem = l;
		// TODO: Shouldn't we AddRef' pNewSelectedItem?
	}

	if(IsWindow()) {
		if(GetStyle() & (LBS_MULTIPLESEL | LBS_EXTENDEDSEL)) {
			return E_FAIL;
		}
		SendMessage(LB_SETCURSEL, newSelectedItem, 0);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_LB_SELECTEDITEM);
	return S_OK;
}

STDMETHODIMP ListBox::get_SelectedItemsArray(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(pValue == NULL) {
		return E_POINTER;
	}

	VariantClear(pValue);

	int items = static_cast<int>(SendMessage(LB_GETSELCOUNT, 0, 0));
	if(items == LB_ERR) {
		return E_FAIL;
	}
	if(items == 0) {
		pValue->parray = NULL;
		pValue->vt = VT_ARRAY | VT_I4;
		return S_OK;
	}

	PINT pBuffer = static_cast<PINT>(HeapAlloc(GetProcessHeap(), 0, items * sizeof(int)));
	if(!pBuffer) {
		return E_OUTOFMEMORY;
	}
	if(SendMessage(LB_GETSELITEMS, items, reinterpret_cast<LPARAM>(pBuffer))) {
		CComSafeArray<LONG> selectedItems;
		selectedItems.Create(items);
		for(int i = 0; i < items; ++i) {
			selectedItems.SetAt(i, pBuffer[i]);
		}
		HeapFree(GetProcessHeap(), 0, pBuffer);

		pValue->parray = SafeArrayCreateVectorEx(VT_I4, 0, items, NULL);
		selectedItems.CopyTo(&pValue->parray);
		pValue->vt = VT_ARRAY | VT_I4;
		return S_OK;
	}

	HeapFree(GetProcessHeap(), 0, pBuffer);
	return E_FAIL;
}

STDMETHODIMP ListBox::put_SelectedItemsArray(VARIANT newValue)
{
	VARIANT* pValue = &newValue;
	if((newValue.vt & VT_VARIANT) == VT_VARIANT && !(newValue.vt & VT_ARRAY) && newValue.pvarVal) {
		pValue = newValue.pvarVal;
	}

	int numberOfItemsToSelect = 0;
	PINT pItems = NULL;
	if(pValue->vt & VT_ARRAY) {
		// an array
		LONG l = 0;
		LONG u = -1;
		VARTYPE vt = 0;
		if(pValue->parray) {
			SafeArrayGetLBound(pValue->parray, 1, &l);
			SafeArrayGetUBound(pValue->parray, 1, &u);
			if(u < l) {
				return E_FAIL;
			}
			numberOfItemsToSelect = u - l + 1;
			pItems = static_cast<PINT>(HeapAlloc(GetProcessHeap(), 0, numberOfItemsToSelect * sizeof(int)));
			if(!pItems) {
				return E_OUTOFMEMORY;
			}
			SafeArrayGetVartype(pValue->parray, &vt);
		}

		for(LONG i = l; i <= u; ++i) {
			if(vt == VT_VARIANT) {
				VARIANT v;
				SafeArrayGetElement(pValue->parray, &i, &v);
				if(SUCCEEDED(VariantChangeType(&v, &v, 0, VT_INT))) {
					pItems[i - l] = v.intVal;
				} else {
					VariantClear(&v);
					HeapFree(GetProcessHeap(), 0, pItems);
					return E_FAIL;
				}
				VariantClear(&v);
			} else if(vt == VT_I1 || vt == VT_UI1 || vt == VT_I2 || vt == VT_UI2 || vt == VT_I4 || vt == VT_UI4 || vt == VT_INT || vt == VT_UINT) {
				SafeArrayGetElement(pValue->parray, &i, &pItems[i - l]);
			} else {
				HeapFree(GetProcessHeap(), 0, pItems);
				return E_FAIL;
			}
		}
	} else {
		// a single value
		numberOfItemsToSelect = 1;
		pItems = static_cast<PINT>(HeapAlloc(GetProcessHeap(), 0, numberOfItemsToSelect * sizeof(int)));
		VARIANT v;
		VariantInit(&v);
		if(SUCCEEDED(VariantChangeType(&v, pValue, 0, VT_INT))) {
			pItems[0] = v.intVal;
		} else {
			VariantClear(&v);
			HeapFree(GetProcessHeap(), 0, pItems);
			return E_FAIL;
		}
		VariantClear(&v);
	}

	// deselect all items
	SendMessage(LB_SELITEMRANGEEX, static_cast<int>(SendMessage(LB_GETCOUNT, 0, 0)) - 1, 0);
	// select all items in pItems
	for(int i = 0; i < numberOfItemsToSelect; i++) {
		SendMessage(LB_SETSEL, TRUE, pItems[i]);
	}
	if(pItems) {
		HeapFree(GetProcessHeap(), 0, pItems);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP ListBox::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_SHOWDRAGIMAGE);
	if(!dragDropStatus.hDragImageList && !dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage(FALSE);
	} else {
		dragDropStatus.ShowDragImage(FALSE);
	}

	FireOnChanged(DISPID_LB_SHOWDRAGIMAGE);
	return S_OK;
}

STDMETHODIMP ListBox::get_Sorted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.sorted = ((GetStyle() & LBS_SORT) == LBS_SORT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.sorted);
	return S_OK;
}

STDMETHODIMP ListBox::put_Sorted(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_SORTED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.sorted != b) {
		properties.sorted = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_SORTED);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP ListBox::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_LB_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_TabStops(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	VariantClear(pValue);
	#ifdef USE_STL
		LONG numberOfStops = static_cast<LONG>(properties.tabStops.size());
	#else
		LONG numberOfStops = static_cast<LONG>(properties.tabStops.GetCount());
	#endif
	if(numberOfStops > 0) {
		// create the array
		pValue->vt = VT_ARRAY | VT_I4;
		pValue->parray = SafeArrayCreateVectorEx(VT_I4, 1, numberOfStops, NULL);

		HFONT hFont = NULL;
		if(IsWindow()) {
			hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
		} else {
			hFont = (properties.useSystemFont ? static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)) : properties.font.currentFont);
		}
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		HFONT hPreviousFont = memoryDC.SelectFont(hFont);
		SIZE textExtent = {0};
		memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
		int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
		memoryDC.SelectFont(hPreviousFont);
		int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

		// transfer the tab stops
		for(LONG i = 1; i <= numberOfStops; ++i) {
			// convert to pixels
			long tabStop = 2 * properties.tabStops[i - 1] * averageCharWidthOfUsedFont / averageCharWidthOfSystemFont;

			SafeArrayPutElement(pValue->parray, &i, &tabStop);
		}
	}
	return S_OK;
}

STDMETHODIMP ListBox::put_TabStops(VARIANT newValue)
{
	PUTPROPPROLOG(DISPID_LB_TABSTOPS);
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		if(newValue.vt == VT_EMPTY) {
			#ifdef USE_STL
				properties.tabStops.clear();
			#else
				properties.tabStops.RemoveAll();
			#endif
			if(properties.tabWidthInPixels == -1) {
				if(SendMessage(LB_SETTABSTOPS, 0, 0)) {
					hr = S_OK;
				}
			} else {
				UINT distance = properties.tabWidthInDTUs;
				if(SendMessage(LB_SETTABSTOPS, 1, reinterpret_cast<LPARAM>(&distance))) {
					hr = S_OK;
				}
			}
			FireViewChange();
		} else if(newValue.vt & VT_ARRAY) {
			// an array
			if(newValue.parray) {
				LONG l = 0;
				SafeArrayGetLBound(newValue.parray, 1, &l);
				LONG u = 0;
				SafeArrayGetUBound(newValue.parray, 1, &u);
				if(u < l) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				UINT* pTabStops = new UINT[u - l + 1];

				HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
				CDC memoryDC;
				memoryDC.CreateCompatibleDC();
				HFONT hPreviousFont = memoryDC.SelectFont(hFont);
				SIZE textExtent = {0};
				memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
				int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
				memoryDC.SelectFont(hPreviousFont);
				int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

				VARTYPE vt = 0;
				SafeArrayGetVartype(newValue.parray, &vt);
				for(LONG i = l; i <= u; ++i) {
					UINT tabStop = 0;
					if(vt == VT_I1 || vt == VT_UI1 || vt == VT_I2 || vt == VT_UI2 || vt == VT_I4 || vt == VT_UI4 || vt == VT_INT || vt == VT_UINT) {
						SafeArrayGetElement(newValue.parray, &i, &tabStop);
					} else if(vt == VT_VARIANT) {
						VARIANT buffer;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						if(SUCCEEDED(VariantChangeType(&buffer, &buffer, 0, VT_UINT))) {
							tabStop = buffer.uintVal;
						} else {
							// invalid arg - raise VB runtime error 380
							delete[] pTabStops;
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
						}
					} else {
						// invalid arg - raise VB runtime error 380
						delete[] pTabStops;
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}

					// convert to dialog template units
					pTabStops[i - l] = tabStop * averageCharWidthOfSystemFont / (2 * averageCharWidthOfUsedFont);
				}
				if(SendMessage(LB_SETTABSTOPS, (u - l + 1), reinterpret_cast<LPARAM>(pTabStops))) {
					hr = S_OK;
				}
				delete[] pTabStops;
				FireViewChange();
			}
		} else {
			// a single value
			VARIANT v;
			VariantInit(&v);
			if(FAILED(VariantChangeType(&v, &newValue, 0, VT_UINT))) {
				// invalid arg - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			UINT tabStop = v.uintVal;

			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			CDC memoryDC;
			memoryDC.CreateCompatibleDC();
			HFONT hPreviousFont = memoryDC.SelectFont(hFont);
			SIZE textExtent = {0};
			memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
			int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
			memoryDC.SelectFont(hPreviousFont);
			int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

			// convert to dialog template units
			tabStop = tabStop * averageCharWidthOfSystemFont / (2 * averageCharWidthOfUsedFont);

			if(SendMessage(LB_SETTABSTOPS, 1, reinterpret_cast<LPARAM>(&tabStop))) {
				hr = S_OK;
			}
			FireViewChange();
		}
	}

	FireOnChanged(DISPID_LB_TABSTOPS);
	return hr;
}

STDMETHODIMP ListBox::get_TabWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.tabWidthInPixels;
	} else {
		HFONT hFont = NULL;
		if(IsWindow()) {
			hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
		} else {
			hFont = (properties.useSystemFont ? static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)) : static_cast<HFONT>(properties.font.currentFont));
		}
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		HFONT hPreviousFont = memoryDC.SelectFont(hFont);
		SIZE textExtent = {0};
		memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
		int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
		memoryDC.SelectFont(hPreviousFont);
		int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

		// convert to pixels
		*pValue = 2 * properties.tabWidthInDTUs * averageCharWidthOfUsedFont / averageCharWidthOfSystemFont;
	}
	return S_OK;
}

STDMETHODIMP ListBox::put_TabWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_LB_TABWIDTH);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.tabWidthInPixels != newValue) {
		LONG oldValueInPixels = properties.tabWidthInPixels;
		properties.tabWidthInPixels = newValue;
		SetDirty(TRUE);
		LONG oldValueInDTUs = properties.tabWidthInDTUs;

		HFONT hFont = NULL;
		if(IsWindow()) {
			hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
		} else {
			hFont = (properties.useSystemFont ? static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)) : static_cast<HFONT>(properties.font.currentFont));
		}
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		HFONT hPreviousFont = memoryDC.SelectFont(hFont);
		SIZE textExtent = {0};
		memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
		int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
		memoryDC.SelectFont(hPreviousFont);
		int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

		// convert to dialog template units
		properties.tabWidthInDTUs = static_cast<LONG>(newValue) * averageCharWidthOfSystemFont / (2 * averageCharWidthOfUsedFont);
		if(IsWindow()) {
			#ifdef USE_STL
				if(properties.tabStops.empty()) {
			#else
				if(properties.tabStops.IsEmpty()) {
			#endif
				if(properties.tabWidthInPixels == -1) {
					if(!SendMessage(LB_SETTABSTOPS, 0, 0)) {
						properties.tabWidthInDTUs = oldValueInDTUs;
						properties.tabWidthInPixels = oldValueInPixels;
						return E_FAIL;
					}
				} else {
					UINT distance = properties.tabWidthInDTUs;
					if(!SendMessage(LB_SETTABSTOPS, 1, reinterpret_cast<LPARAM>(&distance))) {
						properties.tabWidthInDTUs = oldValueInDTUs;
						properties.tabWidthInPixels = oldValueInPixels;
						return E_FAIL;
					}
				}
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_LB_TABWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ListBox::get_ToolTips(ToolTipsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ToolTipsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(!toolTipStatus.hWndToolTip) {
			properties.toolTips = static_cast<ToolTipsConstants>(0);
		}
	}

	*pValue = properties.toolTips;
	return S_OK;
}

STDMETHODIMP ListBox::put_ToolTips(ToolTipsConstants newValue)
{
	PUTPROPPROLOG(DISPID_LB_TOOLTIPS);
	if(properties.toolTips != newValue) {
		properties.toolTips = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.toolTips) {
				if(!toolTipStatus.hWndToolTip) {
					toolTipStatus.hWndToolTip = CreateWindowEx(0, TOOLTIPS_CLASS, NULL, WS_POPUP | TTS_NOPREFIX, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, *this, NULL, NULL, NULL);
					if(toolTipStatus.hWndToolTip) {
						TOOLINFO toolInfo = {0};
						toolInfo.cbSize = sizeof(TOOLINFO);
						toolInfo.uFlags = TTF_TRANSPARENT;
						toolInfo.hwnd = *this;
						toolInfo.lpszText = LPSTR_TEXTCALLBACK;
						GetClientRect(&toolInfo.rect);
						SendMessage(toolTipStatus.hWndToolTip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&toolInfo));
						SendMessage(toolTipStatus.hWndToolTip, WM_SETFONT, reinterpret_cast<WPARAM>(GetFont()), MAKELPARAM(TRUE, 0));
					}
				}
			} else if(::IsWindow(toolTipStatus.hWndToolTip)) {
				::DestroyWindow(toolTipStatus.hWndToolTip);
				toolTipStatus.hWndToolTip = NULL;
				toolTipStatus.toolTipItem = -1;
			}
		}
		FireOnChanged(DISPID_LB_TOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP ListBox::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_LB_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP ListBox::get_Version(BSTR* pValue)
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

STDMETHODIMP ListBox::get_VirtualItemCount(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	LONG virtualItemCount = 0;
	if(!IsInDesignMode() && IsWindow()) {
		virtualItemCount = static_cast<LONG>(SendMessage(LB_GETCOUNT, 0, 0));
	}

	*pValue = virtualItemCount;
	return S_OK;
}

STDMETHODIMP ListBox::put_VirtualItemCount(LONG newValue)
{
	PUTPROPPROLOG(DISPID_LB_VIRTUALITEMCOUNT);
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(IsWindow()) {
		SendMessage(LB_SETCOUNT, newValue, 0);
	}
	FireOnChanged(DISPID_LB_VIRTUALITEMCOUNT);
	return S_OK;
}

STDMETHODIMP ListBox::get_VirtualMode(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.virtualMode = ((GetStyle() & LBS_NODATA) == LBS_NODATA);
	}

	*pValue = BOOL2VARIANTBOOL(properties.virtualMode);
	return S_OK;
}

STDMETHODIMP ListBox::put_VirtualMode(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_LB_VIRTUALMODE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.virtualMode != b) {
		properties.virtualMode = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_LB_VIRTUALMODE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ListBox::BeginDrag(IListBoxItemContainer* pDraggedItems, OLE_HANDLE hDragImageList/* = NULL*/, OLE_XPOS_PIXELS* pXHotSpot/* = NULL*/, OLE_YPOS_PIXELS* pYHotSpot/* = NULL*/)
{
	ATLASSUME(pDraggedItems);
	if(!pDraggedItems) {
		return E_POINTER;
	}

	int xHotSpot = 0;
	if(pXHotSpot) {
		xHotSpot = *pXHotSpot;
	}
	int yHotSpot = 0;
	if(pYHotSpot) {
		yHotSpot = *pYHotSpot;
	}
	HRESULT hr = dragDropStatus.BeginDrag(*this, pDraggedItems, static_cast<HIMAGELIST>(LongToHandle(hDragImageList)), &xHotSpot, &yHotSpot);
	SetCapture();
	if(pXHotSpot) {
		*pXHotSpot = xHotSpot;
	}
	if(pYHotSpot) {
		*pYHotSpot = yHotSpot;
	}

	if(dragDropStatus.hDragImageList) {
		ImageList_BeginDrag(dragDropStatus.hDragImageList, 0, xHotSpot, yHotSpot);
		dragDropStatus.dragImageIsHidden = 0;
		ImageList_DragEnter(0, 0, 0);

		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}
	return hr;
}

STDMETHODIMP ListBox::CreateItemContainer(VARIANT items/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IListBoxItemContainer** ppContainer/* = NULL*/)
{
	ATLASSERT_POINTER(ppContainer, IListBoxItemContainer*);
	if(!ppContainer) {
		return E_POINTER;
	}

	*ppContainer = NULL;
	CComObject<ListBoxItemContainer>* pLBItemContainerObj = NULL;
	CComObject<ListBoxItemContainer>::CreateInstance(&pLBItemContainerObj);
	pLBItemContainerObj->AddRef();

	// clone all settings
	pLBItemContainerObj->SetOwner(this);
	pLBItemContainerObj->properties.useIndexes = ((GetStyle() & LBS_NODATA) == LBS_NODATA);

	pLBItemContainerObj->QueryInterface(__uuidof(IListBoxItemContainer), reinterpret_cast<LPVOID*>(ppContainer));
	pLBItemContainerObj->Release();

	if(*ppContainer) {
		(*ppContainer)->Add(items);
		RegisterItemContainer(static_cast<IItemContainer*>(pLBItemContainerObj));
	}
	return S_OK;
}

STDMETHODIMP ListBox::DeselectItems(VARIANT firstItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, VARIANT lastItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/)
{
	LONG numberOfItems = static_cast<LONG>(SendMessage(LB_GETCOUNT, 0, 0));

	LONG firstItemIndex = 0;
	if(firstItem.vt != VT_ERROR) {
		firstItemIndex = -1;
		if(firstItem.vt == VT_DISPATCH) {
			CComQIPtr<IListBoxItem> pItem = firstItem.pdispVal;
			if(pItem) {
				pItem->get_Index(&firstItemIndex);
			}
		} else {
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &firstItem, 0, VT_UINT))) {
				firstItemIndex = v.uintVal;
			}
		}
		if(firstItemIndex == -1) {
			return E_INVALIDARG;
		}
	}
	LONG lastItemIndex;
	if(lastItem.vt != VT_ERROR) {
		lastItemIndex = -1;
		if(lastItem.vt == VT_DISPATCH) {
			CComQIPtr<IListBoxItem> pItem = lastItem.pdispVal;
			if(pItem) {
				pItem->get_Index(&lastItemIndex);
			}
		} else {
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &lastItem, 0, VT_UINT))) {
				lastItemIndex = v.uintVal;
			}
		}
		if(lastItemIndex == -1) {
			return E_INVALIDARG;
		}
	} else {
		lastItemIndex = numberOfItems - 1;
	}
	if(lastItemIndex == -1) {
		return S_OK;
	}

	if(firstItemIndex == 0 && lastItemIndex == numberOfItems - 1) {
		if(SendMessage(LB_SETSEL, FALSE, -1) != LB_ERR) {
			return S_OK;
		}
		return E_FAIL;
	}

	if(SendMessage(LB_SELITEMRANGEEX, lastItemIndex, firstItemIndex) != LB_ERR) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::EndDrag(VARIANT_BOOL abort)
{
	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	KillTimer(timers.ID_DRAGSCROLL);
	ReleaseCapture();
	if(dragDropStatus.hDragImageList) {
		dragDropStatus.HideDragImage(TRUE);
		ImageList_EndDrag();
	}

	HRESULT hr = S_OK;
	if(abort) {
		hr = Raise_AbortedDrag();
	} else {
		hr = Raise_Drop();
	}

	dragDropStatus.EndDrag();
	Invalidate();

	return hr;
}

STDMETHODIMP ListBox::FinishOLEDragDrop(void)
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

STDMETHODIMP ListBox::FindItemByItemData(LONG itemData, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IListBoxItem** ppFoundItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppFoundItem, IListBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IListBoxItem> pItem = startAfterItem.pdispVal;
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

	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}

	if(expectsString) {
		int numberOfItems = static_cast<int>(SendMessage(LB_GETCOUNT, 0, 0));
		for(int i = itemIndex + 1; i < numberOfItems; i++) {
			if(static_cast<LONG>(SendMessage(LB_GETITEMDATA, i, 0)) == itemData) {
				ClassFactory::InitListItem(i, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
				return S_OK;
			}
		}
		for(int i = 0; i <= itemIndex; i++) {
			if(static_cast<LONG>(SendMessage(LB_GETITEMDATA, i, 0)) == itemData) {
				ClassFactory::InitListItem(i, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
				return S_OK;
			}
		}
		*ppFoundItem = NULL;
		return S_OK;
	} else {
		itemIndex = static_cast<LONG>(SendMessage(LB_FINDSTRING, itemIndex, itemData));
		ClassFactory::InitListItem(itemIndex, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
		return S_OK;
	}
}

STDMETHODIMP ListBox::FindItemByText(BSTR searchString, VARIANT_BOOL exactMatch/* = VARIANT_TRUE*/, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IListBoxItem** ppFoundItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppFoundItem, IListBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IListBoxItem> pItem = startAfterItem.pdispVal;
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

	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	if(!expectsString) {
		return E_FAIL;
	}

	COLE2T converter(searchString);
	LPTSTR pBuffer = converter;
	itemIndex = static_cast<LONG>(SendMessage((exactMatch == VARIANT_FALSE ? LB_FINDSTRING : LB_FINDSTRINGEXACT), itemIndex, reinterpret_cast<LPARAM>(pBuffer)));
	ClassFactory::InitListItem(itemIndex, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
	return S_OK;
}

STDMETHODIMP ListBox::GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, IListBoxItem** ppListItem)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppListItem, IListBoxItem*);
	if(!ppListItem) {
		return E_POINTER;
	}

	int proposedItemIndex = -1;
	InsertMarkPositionConstants proposedRelativePosition = impNowhere;

	POINT pt = {x, y};
	int numberOfItems = SendMessage(LB_GETCOUNT, 0, 0);
	int firstVisibleItem = static_cast<int>(SendMessage(LB_GETTOPINDEX, 0, 0));
	RECT clientRectangle = {0};
	GetClientRect(&clientRectangle);
	BOOL abort = FALSE;
	for(int itemIndex = max(firstVisibleItem - 1, 0); itemIndex < numberOfItems; ++itemIndex) {
		CRect itemBoundingRectangle;
		SendMessage(LB_GETITEMRECT, itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
		if(pt.x >= itemBoundingRectangle.left && pt.x <= itemBoundingRectangle.right) {
			if(itemIndex == firstVisibleItem && pt.y < itemBoundingRectangle.top) {
				proposedItemIndex = itemIndex;
				proposedRelativePosition = impBefore;
			} else if(pt.y >= itemBoundingRectangle.top) {
				proposedItemIndex = itemIndex;
				if(pt.y < itemBoundingRectangle.CenterPoint().y) {
					proposedRelativePosition = impBefore;
				} else {
					proposedRelativePosition = impAfter;
				}
			}
		}
		if(abort) {
			break;
		} else if(itemBoundingRectangle.top > clientRectangle.bottom || itemBoundingRectangle.left > clientRectangle.right) {
			abort = TRUE;
		}
	}

	if(proposedItemIndex == -1) {
		*ppListItem = NULL;
		*pRelativePosition = impNowhere;
	} else {
		*pRelativePosition = proposedRelativePosition;
		ClassFactory::InitListItem(proposedItemIndex, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppListItem), FALSE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, IListBoxItem** ppListItem)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppListItem, IListBoxItem*);
	if(!ppListItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(insertMark.hidden) {
			*ppListItem = NULL;
			*pRelativePosition = impNowhere;
		} else {
			if(insertMark.afterItem) {
				*pRelativePosition = impAfter;
			} else {
				*pRelativePosition = impBefore;
			}
			ClassFactory::InitListItem(insertMark.itemIndex, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppListItem), FALSE);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::GetInsertMarkRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		switch(properties.insertMarkStyle) {
			case imsImproved:
			{
				RECT itemBoundingRectangle = {0};
				RECT insertMarkRect = {0};
				if(insertMark.itemIndex != -1) {
					SendMessage(LB_GETITEMRECT, insertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
					if(insertMark.afterItem) {
						insertMarkRect.top = itemBoundingRectangle.bottom - 3;
						insertMarkRect.bottom = itemBoundingRectangle.bottom + 3;
					} else {
						insertMarkRect.top = itemBoundingRectangle.top - 3;
						insertMarkRect.bottom = itemBoundingRectangle.top + 3;
					}
					insertMarkRect.left = itemBoundingRectangle.left;
					insertMarkRect.right = itemBoundingRectangle.right;
				}

				if(pXLeft) {
					*pXLeft = insertMarkRect.left;
				}
				if(pYTop) {
					*pYTop = insertMarkRect.top;
				}
				if(pXRight) {
					*pXRight = insertMarkRect.right;
				}
				if(pYBottom) {
					*pYBottom = insertMarkRect.bottom;
				}

				hr = S_OK;
				break;
			}
		}
	}
	return hr;
}

STDMETHODIMP ListBox::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IListBoxItem** ppHitItem)
{
	ATLASSERT_POINTER(pHitTestDetails, HitTestConstants);
	ATLASSERT_POINTER(ppHitItem, IListBoxItem*);
	if(!pHitTestDetails || !ppHitItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		ClassFactory::InitListItem(HitTest(x, y, pHitTestDetails), this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppHitItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP ListBox::OLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, IListBoxItemContainer* pDraggedItems/* = NULL*/, LONG itemCountToDisplay/* = -1*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
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

STDMETHODIMP ListBox::PrepareForItemInsertions(LONG numberOfItems, LONG averageStringWidth, LONG* pAllocatedItems)
{
	if(IsWindow()) {
		*pAllocatedItems = SendMessage(LB_INITSTORAGE, numberOfItems, numberOfItems * averageStringWidth * sizeof(TCHAR));
		if(*pAllocatedItems == LB_ERRSPACE) {
			*pAllocatedItems = 0;
			return E_OUTOFMEMORY;
		}
		#ifdef USE_STL
			itemIDs.reserve(numberOfItems);
		#endif
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::Refresh(void)
{
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		Invalidate();
		UpdateWindow();
		dragDropStatus.ShowDragImage(TRUE);
	}
	return S_OK;
}

STDMETHODIMP ListBox::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP ListBox::SelectItemByItemData(LONG itemData, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IListBoxItem** ppFoundItem/* = NULL*/)
{
	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}

	if(expectsString) {
		// cannot use LB_SELECTSTRING, so use emulation
		HRESULT hr = FindItemByItemData(itemData, startAfterItem, ppFoundItem);
		if(SUCCEEDED(hr)) {
			hr = putref_SelectedItem(*ppFoundItem);
		}
		return hr;
	}

	ATLASSERT_POINTER(ppFoundItem, IListBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IListBoxItem> pItem = startAfterItem.pdispVal;
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

	itemIndex = static_cast<LONG>(SendMessage(LB_SELECTSTRING, itemIndex, itemData));
	ClassFactory::InitListItem(itemIndex, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
	return S_OK;
}

STDMETHODIMP ListBox::SelectItemByText(BSTR searchString, VARIANT_BOOL exactMatch/* = VARIANT_TRUE*/, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IListBoxItem** ppFoundItem/* = NULL*/)
{
	if(exactMatch != VARIANT_FALSE) {
		// cannot use LB_SELECTSTRING, so use emulation
		HRESULT hr = FindItemByText(searchString, exactMatch, startAfterItem, ppFoundItem);
		if(SUCCEEDED(hr)) {
			hr = putref_SelectedItem(*ppFoundItem);
		}
		return hr;
	}

	ATLASSERT_POINTER(ppFoundItem, IListBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IListBoxItem> pItem = startAfterItem.pdispVal;
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

	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	if(!expectsString) {
		return E_FAIL;
	}

	COLE2T converter(searchString);
	LPTSTR pBuffer = converter;
	itemIndex = static_cast<LONG>(SendMessage(LB_SELECTSTRING, itemIndex, reinterpret_cast<LPARAM>(pBuffer)));
	ClassFactory::InitListItem(itemIndex, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
	return S_OK;
}

STDMETHODIMP ListBox::SelectItems(VARIANT firstItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, VARIANT lastItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/)
{
	LONG numberOfItems = static_cast<LONG>(SendMessage(LB_GETCOUNT, 0, 0));

	LONG firstItemIndex = 0;
	if(firstItem.vt != VT_ERROR) {
		firstItemIndex = -1;
		if(firstItem.vt == VT_DISPATCH) {
			CComQIPtr<IListBoxItem> pItem = firstItem.pdispVal;
			if(pItem) {
				pItem->get_Index(&firstItemIndex);
			}
		} else {
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &firstItem, 0, VT_UINT))) {
				firstItemIndex = v.uintVal;
			}
		}
		if(firstItemIndex == -1) {
			return E_INVALIDARG;
		}
	}
	LONG lastItemIndex;
	if(lastItem.vt != VT_ERROR) {
		lastItemIndex = -1;
		if(lastItem.vt == VT_DISPATCH) {
			CComQIPtr<IListBoxItem> pItem = lastItem.pdispVal;
			if(pItem) {
				pItem->get_Index(&lastItemIndex);
			}
		} else {
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &lastItem, 0, VT_UINT))) {
				lastItemIndex = v.uintVal;
			}
		}
		if(lastItemIndex == -1) {
			return E_INVALIDARG;
		}
	} else {
		lastItemIndex = numberOfItems - 1;
	}
	if(lastItemIndex == -1) {
		return S_OK;
	}

	if(firstItemIndex == 0 && lastItemIndex == numberOfItems - 1) {
		if(SendMessage(LB_SETSEL, TRUE, -1) != LB_ERR) {
			return S_OK;
		}
		return E_FAIL;
	}

	if(firstItemIndex == lastItemIndex) {
		// we've to select 2 items and then deselect one
		if(lastItemIndex + 1 == numberOfItems) {
			if(SendMessage(LB_SELITEMRANGEEX, firstItemIndex - 1, lastItemIndex) != LB_ERR) {
				if(SendMessage(LB_SELITEMRANGEEX, firstItemIndex - 1, lastItemIndex - 1) != LB_ERR) {
					return S_OK;
				}
			}
		} else {
			if(SendMessage(LB_SELITEMRANGEEX, firstItemIndex, lastItemIndex + 1) != LB_ERR) {
				if(SendMessage(LB_SELITEMRANGEEX, firstItemIndex + 1, lastItemIndex + 1) != LB_ERR) {
					return S_OK;
				}
			}
		}
	} else {
		if(SendMessage(LB_SELITEMRANGEEX, firstItemIndex, lastItemIndex) != LB_ERR) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListBox::SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, IListBoxItem* pListItem)
{
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		int itemIndex = -1;
		if(pListItem) {
			LONG l = -1;
			pListItem->get_Index(&l);
			itemIndex = l;
		}

		int oldItemIndex = insertMark.itemIndex;
		BOOL oldAfterItem = insertMark.afterItem;
		switch(relativePosition) {
			case impNowhere:
				insertMark.itemIndex = -1;
				insertMark.afterItem = FALSE;
				insertMark.hidden = TRUE;
				break;
			case impBefore:
				insertMark.itemIndex = itemIndex;
				insertMark.afterItem = FALSE;
				insertMark.hidden = FALSE;
				break;
			case impAfter:
				insertMark.itemIndex = itemIndex;
				insertMark.afterItem = TRUE;
				insertMark.hidden = FALSE;
				break;
		}

		switch(properties.insertMarkStyle) {
			case imsNative:
				dragDropStatus.HideDragImage(TRUE);
				DrawInsert(GetParent(), *this, insertMark.afterItem ? insertMark.itemIndex + 1 : insertMark.itemIndex);
				dragDropStatus.ShowDragImage(TRUE);
				hr = S_OK;
				break;
			case imsImproved:
			{
				RECT itemBoundingRectangle = {0};
				CRect oldInsertMarkRect;
				if(oldItemIndex != -1) {
					SendMessage(LB_GETITEMRECT, oldItemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
					if(oldAfterItem) {
						oldInsertMarkRect.top = itemBoundingRectangle.bottom - 3;
						oldInsertMarkRect.bottom = itemBoundingRectangle.bottom + 3;
					} else {
						oldInsertMarkRect.top = itemBoundingRectangle.top - 3;
						oldInsertMarkRect.bottom = itemBoundingRectangle.top + 3;
					}
					oldInsertMarkRect.left = itemBoundingRectangle.left;
					oldInsertMarkRect.right = itemBoundingRectangle.right;
				}

				CRect newInsertMarkRect;
				if(insertMark.itemIndex != -1) {
					SendMessage(LB_GETITEMRECT, insertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
					if(insertMark.afterItem) {
						newInsertMarkRect.top = itemBoundingRectangle.bottom - 3;
						newInsertMarkRect.bottom = itemBoundingRectangle.bottom + 3;
					} else {
						newInsertMarkRect.top = itemBoundingRectangle.top - 3;
						newInsertMarkRect.bottom = itemBoundingRectangle.top + 3;
					}
					newInsertMarkRect.left = itemBoundingRectangle.left;
					newInsertMarkRect.right = itemBoundingRectangle.right;
				}

				if(newInsertMarkRect != oldInsertMarkRect) {
					// redraw
					if(oldInsertMarkRect.Width() > 0) {
						InvalidateRect(&oldInsertMarkRect);
					}
					if(newInsertMarkRect.Width() > 0) {
						InvalidateRect(&newInsertMarkRect);
					}
				}

				hr = S_OK;
				break;
			}
		}
	}

	return hr;
}


LRESULT ListBox::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT ListBox::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);

	wasHandled = FALSE;
	return 0;
}

LRESULT ListBox::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(*this);
	}
	return lr;
}

LRESULT ListBox::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(*this);

	wasHandled = FALSE;
	return 0;
}

LRESULT ListBox::OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if((ime != imeNoControl) && (GetFocus() == *this)) {
		// we've the focus, so configure the IME
		SetCurrentIMEContextMode(*this, ime);
	}
	return lr;
}

LRESULT ListBox::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ListBox::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ListBox::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<ListBox>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	return lr;
}

LRESULT ListBox::OnLButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ListBox::OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	int itemIndex = -1;
	LRESULT lr = 0;
	mouseStatus.mouseDownMessageToSendOnMouseUp.hwnd = NULL;
	if(properties.allowDragDrop) {
		HitTestConstants dummy = static_cast<HitTestConstants>(0);
		itemIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
		if(itemIndex != -1 && SendMessage(LB_GETSEL, itemIndex, 0)) {
			// the item is selected - eat the message
			mouseStatus.mouseDownMessageToSendOnMouseUp.hwnd = *this;
			mouseStatus.mouseDownMessageToSendOnMouseUp.message = message;
			mouseStatus.mouseDownMessageToSendOnMouseUp.wParam = wParam;
			mouseStatus.mouseDownMessageToSendOnMouseUp.lParam = lParam;
			/*mouseStatus.mouseDownMessageToSendOnMouseUp.time = GetMessageTime();
			DWORD pos = GetMessagePos();
			mouseStatus.mouseDownMessageToSendOnMouseUp.pt.x = GET_X_LPARAM(pos);
			mouseStatus.mouseDownMessageToSendOnMouseUp.pt.y = GET_Y_LPARAM(pos);*/
		} else {
			lr = DefWindowProc(message, wParam, lParam);

			int newCaretItem = static_cast<int>(SendMessage(LB_GETCARETINDEX, 0, 0));
			if(newCaretItem != currentCaretItem) {
				Raise_CaretChanged(currentCaretItem, newCaretItem);
			}
			Raise_SelectionChanged();
		}
	} else {
		lr = DefWindowProc(message, wParam, lParam);

		int newCaretItem = static_cast<int>(SendMessage(LB_GETCARETINDEX, 0, 0));
		if(newCaretItem != currentCaretItem) {
			Raise_CaretChanged(currentCaretItem, newCaretItem);
		}
		Raise_SelectionChanged();
	}

	if(properties.allowDragDrop) {
		if(itemIndex != -1) {
			dragDropStatus.candidate.itemIndex = itemIndex;
			dragDropStatus.candidate.button = MK_LBUTTON;
			dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
			dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);
			SetCapture();
		}
	}

	return lr;
}

LRESULT ListBox::OnLButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.allowDragDrop) {
		dragDropStatus.candidate.itemIndex = -1;
		if(mouseStatus.mouseDownMessageToSendOnMouseUp.hwnd == *this) {
			DefWindowProc(mouseStatus.mouseDownMessageToSendOnMouseUp.message, mouseStatus.mouseDownMessageToSendOnMouseUp.wParam, mouseStatus.mouseDownMessageToSendOnMouseUp.lParam);

			int newCaretItem = static_cast<int>(SendMessage(LB_GETCARETINDEX, 0, 0));
			if(newCaretItem != currentCaretItem) {
				Raise_CaretChanged(currentCaretItem, newCaretItem);
			}
			Raise_SelectionChanged();
		}
	}

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ListBox::OnMButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ListBox::OnMButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ListBox::OnMButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ListBox::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT ListBox::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ListBox::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);
	}

	return 0;
}

LRESULT ListBox::OnMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &hitTestDetails);
		if(itemIndex != toolTipStatus.toolTipItem) {
			SendMessage(toolTipStatus.hWndToolTip, TTM_POP, 0, 0);
		}
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	if(properties.allowDragDrop) {
		if(dragDropStatus.candidate.itemIndex != -1) {
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
				mouseStatus.mouseDownMessageToSendOnMouseUp.hwnd = NULL;
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
				HitTest(dragDropStatus.candidate.position.x, dragDropStatus.candidate.position.y, &hitTestDetails);
				CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(dragDropStatus.candidate.itemIndex, this);
				switch(dragDropStatus.candidate.button) {
					case MK_LBUTTON:
						button |= 1/*MouseButtonConstants.vbLeftButton*/;
						Raise_ItemBeginDrag(pLBItem, button, shift, dragDropStatus.candidate.position.x, dragDropStatus.candidate.position.y, hitTestDetails);
						break;
					case MK_RBUTTON:
						button |= 2/*MouseButtonConstants.vbRightButton*/;
						Raise_ItemBeginRDrag(pLBItem, button, shift, dragDropStatus.candidate.position.x, dragDropStatus.candidate.position.y, hitTestDetails);
						break;
				}
				dragDropStatus.candidate.itemIndex = -1;
			}
		}
	}

	if(dragDropStatus.IsDragging()) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_DragMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ListBox::OnNCMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}
	return 0;
}

LRESULT ListBox::OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ListBox::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	switch(properties.insertMarkStyle) {
		case imsImproved:
			if(message == WM_PAINT) {
				HDC hDC = GetDC();
				DrawInsertionMark(hDC);
				ReleaseDC(hDC);
			} else {
				DrawInsertionMark(reinterpret_cast<HDC>(wParam));
			}
			break;
	}
	return lr;
}

LRESULT ListBox::OnRButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ListBox::OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(properties.allowDragDrop) {
		HitTestConstants dummy = static_cast<HitTestConstants>(0);
		int itemIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &dummy);
		if(itemIndex != -1) {
			dragDropStatus.candidate.itemIndex = itemIndex;
			dragDropStatus.candidate.button = MK_RBUTTON;
			dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
			dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);
			SetCapture();
		}
	}

	return lr;
}

LRESULT ListBox::OnRButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.allowDragDrop) {
		dragDropStatus.candidate.itemIndex = -1;
	}

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ListBox::OnScroll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL tmp = insertMark.hidden;
	if(!insertMark.hidden) {
		insertMark.hidden = TRUE;
		switch(properties.insertMarkStyle) {
			case imsNative:
				dragDropStatus.HideDragImage(TRUE);
				DrawInsert(GetParent(), *this, -1);
				dragDropStatus.ShowDragImage(TRUE);
				break;
			case imsImproved:
			{
				// remove the insertion mark - we might get drawing glitches otherwise
				CRect oldInsertMarkRect;
				// calculate the current insertion mark's rectangle
				RECT itemBoundingRectangle = {0};
				SendMessage(LB_GETITEMRECT, insertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
				if(insertMark.afterItem) {
					oldInsertMarkRect.top = itemBoundingRectangle.bottom - 3;
					oldInsertMarkRect.bottom = itemBoundingRectangle.bottom + 3;
				} else {
					oldInsertMarkRect.top = itemBoundingRectangle.top - 3;
					oldInsertMarkRect.bottom = itemBoundingRectangle.top + 3;
				}
				oldInsertMarkRect.left = itemBoundingRectangle.left;
				oldInsertMarkRect.right = itemBoundingRectangle.right;

				// redraw
				if(oldInsertMarkRect.Width() > 0) {
					InvalidateRect(&oldInsertMarkRect);
				}
				break;
			}
		}
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!tmp) {
		// now draw the insertion mark again
		insertMark.hidden = FALSE;
		switch(properties.insertMarkStyle) {
			case imsNative:
				dragDropStatus.HideDragImage(TRUE);
				DrawInsert(GetParent(), *this, insertMark.afterItem ? insertMark.itemIndex + 1 : insertMark.itemIndex);
				dragDropStatus.ShowDragImage(TRUE);
				break;
			case imsImproved:
			{
				HDC hDC = GetDC();
				DrawInsertionMark(hDC);
				ReleaseDC(hDC);
				break;
			}
		}
	}

	return lr;
}

LRESULT ListBox::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
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

LRESULT ListBox::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<ListBox>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<ListBox>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT ListBox::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_LB_FONT) == S_FALSE) {
		return 0;
	}

	if(!IsInDesignMode()) {
		if(properties.setItemHeight != -1) {
			if(!(GetStyle() & LBS_OWNERDRAWVARIABLE)) {
				properties.itemHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(LB_GETITEMHEIGHT, 0, 0));
			}
		}
	}
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(toolTipStatus.hWndToolTip) {
		SendMessage(toolTipStatus.hWndToolTip, WM_SETFONT, wParam, lParam);
	}
	if(!IsInDesignMode()) {
		if(properties.setItemHeight != -1) {
			if(!(GetStyle() & LBS_OWNERDRAWVARIABLE)) {
				SendMessage(LB_SETITEMHEIGHT, 0, properties.itemHeight);
			}
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
		FireOnChanged(DISPID_LB_FONT);
	}

	return lr;
}

LRESULT ListBox::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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

LRESULT ListBox::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT ListBox::OnSize(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(toolTipStatus.hWndToolTip) {
		TOOLINFO toolInfo;
		if(properties.toolTips & ttLabelTips) {
			toolTipStatus.InvalidateToolTipItem();
		}
		toolInfo.cbSize = sizeof(TOOLINFO);
		toolInfo.hwnd = *this;
		toolInfo.uId = 0;
		GetClientRect(&toolInfo.rect);
		SendMessage(toolTipStatus.hWndToolTip, TTM_NEWTOOLRECT, 0, reinterpret_cast<LPARAM>(&toolInfo));
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ListBox::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT ListBox::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);

				int newCaretItem = static_cast<int>(SendMessage(LB_GETCARETINDEX, 0, 0));
				if(newCaretItem != currentCaretItem) {
					Raise_CaretChanged(currentCaretItem, newCaretItem);
				}
				Raise_SelectionChanged();
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		case timers.ID_DRAGSCROLL:
			AutoScroll();
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ListBox::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
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

LRESULT ListBox::OnXButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

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
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ListBox::OnXButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

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
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ListBox::OnXButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

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
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ListBox::OnReflectedCharToItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	LONG itemToSelect = -1;
	Raise_ProcessCharacterInput(LOWORD(wParam), shift, &itemToSelect);
	return itemToSelect;
}

LRESULT ListBox::OnReflectedCompareItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPCOMPAREITEMSTRUCT pDetails = reinterpret_cast<LPCOMPAREITEMSTRUCT>(lParam);
	ATLASSERT(pDetails->CtlType == ODT_LISTBOX);
	ATLASSERT(pDetails->hwndItem == *this);

	// NOTE: The indexes may be -1.
	CComPtr<IListBoxItem> pFirstItem = ClassFactory::InitListItem(pDetails->itemID1, this, FALSE, pDetails->itemData1);
	CComPtr<IListBoxItem> pSecondItem = ClassFactory::InitListItem(pDetails->itemID2, this, FALSE, pDetails->itemData2);
	CompareResultConstants result = crEqual;
	Raise_CompareItems(pFirstItem, pSecondItem, pDetails->dwLocaleId, &result);
	return result;
}

LRESULT ListBox::OnReflectedCtlColorListBox(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
	SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.foreColor));
	if(!(properties.backColor & 0x80000000)) {
		SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
		return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
	} else {
		return reinterpret_cast<LRESULT>(GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
	}
}

LRESULT ListBox::OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPDRAWITEMSTRUCT pDetails = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);

	ATLASSERT(pDetails->CtlType == ODT_LISTBOX);
	ATLASSERT((pDetails->itemState & (ODS_GRAYED | ODS_CHECKED | ODS_DEFAULT | ODS_COMBOBOXEDIT | ODS_HOTLIGHT | ODS_INACTIVE)) == 0);

	if(pDetails->hwndItem == *this) {
		int itemIndex = pDetails->itemID;
		CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);
		Raise_OwnerDrawItem(pLBItem, static_cast<OwnerDrawActionConstants>(pDetails->itemAction), static_cast<OwnerDrawItemStateConstants>(pDetails->itemState), HandleToLong(pDetails->hDC), reinterpret_cast<RECTANGLE*>(&pDetails->rcItem));
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ListBox::OnReflectedMeasureItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPMEASUREITEMSTRUCT pDetails = reinterpret_cast<LPMEASUREITEMSTRUCT>(lParam);
	ATLASSERT(pDetails->CtlType == ODT_LISTBOX);

	if(pDetails->CtlType == ODT_LISTBOX) {
		int itemIndex = pDetails->itemID;
		CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
		Raise_MeasureItem(pLBItem, reinterpret_cast<OLE_YSIZE_PIXELS*>(&pDetails->itemHeight));
		return TRUE;
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT ListBox::OnReflectedVKeyToItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	LONG itemToSelect = -1;
	Raise_ProcessKeyStroke(LOWORD(wParam), shift, &itemToSelect);
	return itemToSelect;
}

LRESULT ListBox::OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
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

LRESULT ListBox::OnAddString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	toolTipStatus.InvalidateToolTipItem();

	if(flags.silentItemInsertions > 0) {
		LONG itemID = GetNewItemID();
		int insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem != LB_ERR) {
			#ifdef USE_STL
				if(insertedItem >= static_cast<int>(itemIDs.size())) {
					itemIDs.push_back(itemID);
				} else {
					itemIDs.insert(itemIDs.begin() + insertedItem, itemID);
				}
			#else
				if(insertedItem >= static_cast<int>(itemIDs.GetCount())) {
					itemIDs.Add(itemID);
				} else {
					itemIDs.InsertAt(insertedItem, itemID);
				}
			#endif
		}
		return insertedItem;
	}

	int insertedItem = LB_ERR;
	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	LPARAM itemData = 0;
	if(expectsString) {
		#ifdef USE_STL
			if(properties.itemDataOfInsertedItems.size() > 0) {
				itemData = properties.itemDataOfInsertedItems.top();
				properties.itemDataOfInsertedItems.pop();
			}
		#else
			if(properties.itemDataOfInsertedItems.GetCount() > 0) {
				itemData = properties.itemDataOfInsertedItems.RemoveHead();
			}
		#endif
	}

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!(properties.disabledEvents & deItemInsertionEvents)) {
		// TODO: For LBS_SORT, this might be the wrong itemIndex.
		int itemIndex = static_cast<int>(SendMessage(LB_GETCOUNT, 0, 0));

		CComObject<VirtualListBoxItem>* pVLBoxItemObj = NULL;
		CComPtr<IVirtualListBoxItem> pVLBoxItemItf = NULL;
		CComObject<VirtualListBoxItem>::CreateInstance(&pVLBoxItemObj);
		pVLBoxItemObj->AddRef();
		pVLBoxItemObj->SetOwner(this);
		pVLBoxItemObj->LoadState(itemIndex, lParam, (expectsString ? itemData : lParam));
		pVLBoxItemObj->QueryInterface(IID_IVirtualListBoxItem, reinterpret_cast<LPVOID*>(&pVLBoxItemItf));
		pVLBoxItemObj->Release();
		Raise_InsertingItem(pVLBoxItemItf, &cancel);
		pVLBoxItemObj = NULL;
	}

	if(cancel == VARIANT_FALSE) {
		LONG itemID = GetNewItemID();
		// finally pass the message to the list box
		insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem != LB_ERR) {
			if(expectsString) {
				SendMessage(LB_SETITEMDATA, insertedItem, itemData);
			}
			#ifdef USE_STL
				if(insertedItem >= static_cast<int>(itemIDs.size())) {
					itemIDs.push_back(itemID);
				} else {
					itemIDs.insert(itemIDs.begin() + insertedItem, itemID);
				}
			#else
				if(insertedItem >= static_cast<int>(itemIDs.GetCount())) {
					itemIDs.Add(itemID);
				} else {
					itemIDs.InsertAt(insertedItem, itemID);
				}
			#endif
			if(!(properties.disabledEvents & deItemInsertionEvents)) {
				CComPtr<IListBoxItem> pLBoxItem = ClassFactory::InitListItem(insertedItem, this);
				if(pLBoxItem) {
					Raise_InsertedItem(pLBoxItem);
				}
			}
		}
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		if(insertedItem != LB_ERR) {
			// maybe we have a new item below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
			int newItemUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			if(newItemUnderMouse != itemUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(itemUnderMouse >= 0) {
					CComPtr<IListBoxItem> pLBoxItem = ClassFactory::InitListItem(itemUnderMouse, this);
					Raise_ItemMouseLeave(pLBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				if(itemUnderMouse >= 0) {
					CComPtr<IListBoxItem> pLBoxItem = ClassFactory::InitListItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pLBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return insertedItem;
}

LRESULT ListBox::OnDeleteString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	toolTipStatus.InvalidateToolTipItem();

	if(flags.silentItemDeletions > 0) {
		LRESULT lr = DefWindowProc(message, wParam, lParam);
		if(lr != LB_ERR) {
			ATLASSERT(!(GetStyle() & LBS_NODATA));
			// ListBoxItem::put_Text will call SetItemID, which will call UpdateItemIDInItemContainers
			//LONG itemID = ItemIndexToID(static_cast<int>(wParam));
			//RemoveItemFromItemContainers(itemID);
			#ifdef USE_STL
				if(static_cast<int>(wParam) >= 0 && static_cast<int>(wParam) < static_cast<int>(itemIDs.size())) {
					itemIDs.erase(itemIDs.begin() + static_cast<int>(wParam));
				}
			#else
				if(static_cast<int>(wParam) >= 0 && static_cast<int>(wParam) < static_cast<int>(itemIDs.GetCount())) {
					itemIDs.RemoveAt(static_cast<int>(wParam));
				}
			#endif
		}
		return lr;
	}

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IListBoxItem> pLBoxItemItf = NULL;
	CComObject<ListBoxItem>* pLBoxItemObj = NULL;
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		CComObject<ListBoxItem>::CreateInstance(&pLBoxItemObj);
		pLBoxItemObj->AddRef();
		pLBoxItemObj->SetOwner(this);
		pLBoxItemObj->Attach(static_cast<int>(wParam));
		pLBoxItemObj->QueryInterface(IID_IListBoxItem, reinterpret_cast<LPVOID*>(&pLBoxItemItf));
		pLBoxItemObj->Release();
		Raise_RemovingItem(pLBoxItemItf, &cancel);
	}

	LRESULT lr = LB_ERR;
	if(cancel == VARIANT_FALSE) {
		if(!(properties.disabledEvents & deMouseEvents)) {
			if(static_cast<int>(wParam) == itemUnderMouse) {
				// we're removing the item below the mouse cursor
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
				HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				Raise_ItemMouseLeave(pLBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				itemUnderMouse = -1;
			}
		}

		if(!(properties.disabledEvents & deFreeItemData)) {
			Raise_FreeItemData(pLBoxItemItf);
		}

		CComPtr<IVirtualListBoxItem> pVLBoxItemItf = NULL;
		if(!(properties.disabledEvents & deItemDeletionEvents)) {
			CComObject<VirtualListBoxItem>* pVLBoxItemObj = NULL;
			CComObject<VirtualListBoxItem>::CreateInstance(&pVLBoxItemObj);
			pVLBoxItemObj->AddRef();
			pVLBoxItemObj->SetOwner(this);
			if(pLBoxItemObj) {
				pLBoxItemObj->SaveState(pVLBoxItemObj);
			}

			pVLBoxItemObj->QueryInterface(IID_IVirtualListBoxItem, reinterpret_cast<LPVOID*>(&pVLBoxItemItf));
			pVLBoxItemObj->Release();
		}
		lr = DefWindowProc(message, wParam, lParam);
		if(lr != LB_ERR) {
			ATLASSERT(!(GetStyle() & LBS_NODATA));
			LONG itemID = ItemIndexToID(static_cast<int>(wParam));
			RemoveItemFromItemContainers(itemID);
			if(itemID >= 0) {
				#ifdef USE_STL
					itemIDs.erase(itemIDs.begin() + static_cast<int>(wParam));
				#else
					itemIDs.RemoveAt(static_cast<int>(wParam));
				#endif
			}
			if(!(properties.disabledEvents & deItemDeletionEvents)) {
				Raise_RemovedItem(pVLBoxItemItf);
			}
			if(static_cast<int>(wParam) == currentCaretItem) {
				Raise_CaretChanged(currentCaretItem, static_cast<int>(SendMessage(LB_GETCARETINDEX, 0, 0)));
			}
			Raise_SelectionChanged();

			if(!(properties.disabledEvents & deMouseEvents)) {
				// maybe we have a new item below the mouse cursor now
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
				int newItemUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				if(newItemUnderMouse != itemUnderMouse) {
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					if(itemUnderMouse >= 0) {
						pLBoxItemItf = ClassFactory::InitListItem(itemUnderMouse, this);
						Raise_ItemMouseLeave(pLBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
					itemUnderMouse = newItemUnderMouse;
					if(itemUnderMouse >= 0) {
						pLBoxItemItf = ClassFactory::InitListItem(itemUnderMouse, this);
						Raise_ItemMouseEnter(pLBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}
	return lr;
}

LRESULT ListBox::OnInsertString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	toolTipStatus.InvalidateToolTipItem();

	if(flags.silentItemInsertions > 0) {
		LONG itemID = GetNewItemID();
		int insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem != LB_ERR) {
			#ifdef USE_STL
				if(insertedItem >= static_cast<int>(itemIDs.size())) {
					itemIDs.push_back(itemID);
				} else {
					itemIDs.insert(itemIDs.begin() + insertedItem, itemID);
				}
			#else
				if(insertedItem >= static_cast<int>(itemIDs.GetCount())) {
					itemIDs.Add(itemID);
				} else {
					itemIDs.InsertAt(insertedItem, itemID);
				}
			#endif
		}
		return insertedItem;
	}

	int insertedItem = LB_ERR;
	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	LPARAM itemData = 0;
	if(expectsString) {
		#ifdef USE_STL
			if(properties.itemDataOfInsertedItems.size() > 0) {
				itemData = properties.itemDataOfInsertedItems.top();
				properties.itemDataOfInsertedItems.pop();
			}
		#else
			if(properties.itemDataOfInsertedItems.GetCount() > 0) {
				itemData = properties.itemDataOfInsertedItems.RemoveHead();
			}
		#endif
	}

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!(properties.disabledEvents & deItemInsertionEvents)) {
		CComObject<VirtualListBoxItem>* pVLBoxItemObj = NULL;
		CComPtr<IVirtualListBoxItem> pVLBoxItemItf = NULL;
		CComObject<VirtualListBoxItem>::CreateInstance(&pVLBoxItemObj);
		pVLBoxItemObj->AddRef();
		pVLBoxItemObj->SetOwner(this);
		pVLBoxItemObj->LoadState(static_cast<int>(wParam), lParam, (expectsString ? itemData : lParam));
		pVLBoxItemObj->QueryInterface(IID_IVirtualListBoxItem, reinterpret_cast<LPVOID*>(&pVLBoxItemItf));
		pVLBoxItemObj->Release();
		Raise_InsertingItem(pVLBoxItemItf, &cancel);
		pVLBoxItemObj = NULL;
	}

	if(cancel == VARIANT_FALSE) {
		LONG itemID = GetNewItemID();
		// finally pass the message to the list box
		insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem >= 0) {
			if(expectsString) {
				SendMessage(LB_SETITEMDATA, insertedItem, itemData);
			}
			#ifdef USE_STL
				if(insertedItem >= static_cast<int>(itemIDs.size())) {
					itemIDs.push_back(itemID);
				} else {
					itemIDs.insert(itemIDs.begin() + insertedItem, itemID);
				}
			#else
				if(insertedItem >= static_cast<int>(itemIDs.GetCount())) {
					itemIDs.Add(itemID);
				} else {
					itemIDs.InsertAt(insertedItem, itemID);
				}
			#endif
			if(!(properties.disabledEvents & deItemInsertionEvents)) {
				CComPtr<IListBoxItem> pLBoxItem = ClassFactory::InitListItem(insertedItem, this);
				if(pLBoxItem) {
					Raise_InsertedItem(pLBoxItem);
				}
			}
		}
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		if(insertedItem != LB_ERR) {
			// maybe we have a new item below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
			int newItemUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			if(newItemUnderMouse != itemUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(itemUnderMouse >= 0) {
					CComPtr<IListBoxItem> pLBoxItem = ClassFactory::InitListItem(itemUnderMouse, this);
					Raise_ItemMouseLeave(pLBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				if(itemUnderMouse >= 0) {
					CComPtr<IListBoxItem> pLBoxItem = ClassFactory::InitListItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pLBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return insertedItem;
}

LRESULT ListBox::OnResetContent(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	toolTipStatus.InvalidateToolTipItem();

	/* NOTE: If LB_DELETESTRING is called for the last item in the control, the control sends LB_RESETCONTENT
	         to itself. */
	if(flags.silentItemDeletions > 0) {
		LRESULT lr = DefWindowProc(message, wParam, lParam);
		#ifdef USE_STL
			itemIDs.clear();
		#else
			itemIDs.RemoveAll();
		#endif
		return lr;
	}

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
			CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemUnderMouse, this);
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
			HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			Raise_ItemMouseLeave(pLBItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
			itemUnderMouse = -1;
		}
	}
	if(!(properties.disabledEvents & deFreeItemData)) {
		Raise_FreeItemData(NULL);
	}
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	#ifdef USE_STL
		itemIDs.clear();
	#else
		itemIDs.RemoveAll();
	#endif
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		Raise_RemovedItem(NULL);
	}
	if(currentCaretItem != -1) {
		Raise_CaretChanged(currentCaretItem, -1);
	}
	Raise_SelectionChanged();

	RemoveItemFromItemContainers(-1);
	return lr;
}

LRESULT ListBox::OnSelectionChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr != LB_ERR) {
		if(GetStyle() & LBS_NOTIFY) {
			GetParent().PostMessage(WM_COMMAND, MAKEWPARAM(GetWindowLongPtr(GWLP_ID) & 0xFFFF, LBN_SELCHANGE), reinterpret_cast<LPARAM>(static_cast<HWND>(*this)));
		}
	}
	return lr;
}

LRESULT ListBox::OnSetCaretIndex(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr != LB_ERR) {
		int newCaretItem = static_cast<int>(SendMessage(LB_GETCARETINDEX, 0, 0));
		if(newCaretItem != currentCaretItem) {
			Raise_CaretChanged(currentCaretItem, newCaretItem);
		}
	}
	return lr;
}

LRESULT ListBox::OnSetColumnWidth(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(wParam == -1) {
		////////////////////////////////////////////////////////////////
		// taken from ReactOS: dll\win32\gdi32\objects\font.c -> GdiGetCharDimensions() (revision 39839)
		static const LPTSTR pAlphabet = TEXT("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ");
		HDC hDC = GetDC();
		ATLASSERT(hDC);
		SIZE sz;
		if(GetTextExtentPoint(hDC, pAlphabet, 52, &sz)) {
			wParam = 15 * ((sz.cx / 26 + 1) / 2);
		}
		ReleaseDC(hDC);
		////////////////////////////////////////////////////////////////
	}

	if(!IsInDesignMode()) {
		properties.columnWidth = wParam;
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ListBox::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	typedef BOOL WINAPI Str_SetPtrAFn(__deref_inout_opt LPSTR*, __in_opt LPCSTR);
	Str_SetPtrAFn* pfnStr_SetPtrA = NULL;
	HMODULE hComctl32DLL = LoadLibrary(TEXT("comctl32.dll"));
	if(hComctl32DLL) {
		pfnStr_SetPtrA = reinterpret_cast<Str_SetPtrAFn*>(GetProcAddress(hComctl32DLL, MAKEINTRESOURCEA(234)));
		FreeLibrary(hComctl32DLL);
	}

	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);

	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	VARIANT_BOOL cancelToolTip = VARIANT_FALSE;
	//if(itemIndex != toolTipStatus.toolTipItem) {
		if(pfnStr_SetPtrA) {
			pfnStr_SetPtrA(&toolTipStatus.pCurrentToolTipA, NULL);
		}
		toolTipStatus.toolTipItem = itemIndex;
		toolTipStatus.inplaceToolTip = FALSE;
		if(itemIndex >= 0) {
			CHAR pBuffer[INFOTIPSIZE];
			CHAR pLabelTip[INFOTIPSIZE];
			pBuffer[0] = 0;
			pLabelTip[0] = 0;
			LPSTR pText = pBuffer;

			if(properties.toolTips & ttLabelTips) {
				TCHAR p[INFOTIPSIZE];
				if(IsItemTruncated(itemIndex, p, INFOTIPSIZE)) {
					CT2A converter(p);
					LPSTR p2 = converter;
					lstrcpynA(pLabelTip, p2, INFOTIPSIZE);
					lstrcpynA(pBuffer, pLabelTip, INFOTIPSIZE);
				}
			}

			BOOL infoTipMode = FALSE;
			if(properties.toolTips & ttInfoTips) {
				CComBSTR text = L"";
				text.Append(pBuffer);
				CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);
				Raise_ItemGetInfoTipText(pLBItem, INFOTIPSIZE - 1, &text, &cancelToolTip);
				if(cancelToolTip == VARIANT_FALSE) {
					int bufferSize = text.Length() + 1;
					if(bufferSize > INFOTIPSIZE) {
						bufferSize = INFOTIPSIZE;
					}
					lstrcpynA(pBuffer, CW2A(text), bufferSize);
				}

				if(pBuffer[0] == '\0') {
					pText = pLabelTip;
				} else if(lstrcmpA(pBuffer, pLabelTip)) {
					infoTipMode = TRUE;
				}
			}
			if(infoTipMode) {
				static const RECT margins = {4, 4, 4, 4};
				SendMessage(toolTipStatus.hWndToolTip, TTM_SETMARGIN, 0, reinterpret_cast<LPARAM>(&margins));
				HDC hDC = GetDC();
				int width = MulDiv(GetDeviceCaps(hDC, LOGPIXELSX), 300, 72);
				int maxWidth = GetDeviceCaps(hDC, HORZRES) * 3 / 4;
				ReleaseDC(hDC);
				SendMessage(toolTipStatus.hWndToolTip, TTM_SETMAXTIPWIDTH, 0, min(width, maxWidth));
			} else {
				static const RECT margins = {0, 0, 0, 0};
				SendMessage(toolTipStatus.hWndToolTip, TTM_SETMARGIN, 0, reinterpret_cast<LPARAM>(&margins));
				toolTipStatus.inplaceToolTip = TRUE;
				SendMessage(toolTipStatus.hWndToolTip, TTM_SETMAXTIPWIDTH, 0, -1);
			}

			if(pfnStr_SetPtrA) {
				pfnStr_SetPtrA(&toolTipStatus.pCurrentToolTipA, pText);
			}
		}
	//}
	if(cancelToolTip == VARIANT_FALSE) {
		pDetails->lpszText = toolTipStatus.pCurrentToolTipA;
	}
	return 0;
}

LRESULT ListBox::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);

	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	VARIANT_BOOL cancelToolTip = VARIANT_FALSE;
	//if(itemIndex != toolTipStatus.toolTipItem) {
		Str_SetPtrW(&toolTipStatus.pCurrentToolTipW, NULL);
		toolTipStatus.toolTipItem = itemIndex;
		toolTipStatus.inplaceToolTip = FALSE;
		if(itemIndex >= 0) {
			WCHAR pBuffer[INFOTIPSIZE];
			WCHAR pLabelTip[INFOTIPSIZE];
			pBuffer[0] = 0;
			pLabelTip[0] = 0;
			LPWSTR pText = pBuffer;

			if(properties.toolTips & ttLabelTips) {
				TCHAR p[INFOTIPSIZE];
				if(IsItemTruncated(itemIndex, p, INFOTIPSIZE)) {
					CT2W converter(p);
					LPWSTR p2 = converter;
					lstrcpynW(pLabelTip, p2, INFOTIPSIZE);
					lstrcpynW(pBuffer, pLabelTip, INFOTIPSIZE);
				}
			}

			BOOL infoTipMode = FALSE;
			if(properties.toolTips & ttInfoTips) {
				CComBSTR text = L"";
				text.Append(pBuffer);
				CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);
				Raise_ItemGetInfoTipText(pLBItem, INFOTIPSIZE - 1, &text, &cancelToolTip);
				if(cancelToolTip == VARIANT_FALSE) {
					int bufferSize = text.Length() + 1;
					if(bufferSize > INFOTIPSIZE) {
						bufferSize = INFOTIPSIZE;
					}
					lstrcpynW(pBuffer, text, bufferSize);
				}

				if(pBuffer[0] == L'\0') {
					pText = pLabelTip;
				} else if(lstrcmpW(pBuffer, pLabelTip)) {
					infoTipMode = TRUE;
				}
			}
			if(infoTipMode) {
				static const RECT margins = {4, 4, 4, 4};
				SendMessage(toolTipStatus.hWndToolTip, TTM_SETMARGIN, 0, reinterpret_cast<LPARAM>(&margins));
				HDC hDC = GetDC();
				int width = MulDiv(GetDeviceCaps(hDC, LOGPIXELSX), 300, 72);
				int maxWidth = GetDeviceCaps(hDC, HORZRES) * 3 / 4;
				ReleaseDC(hDC);
				SendMessage(toolTipStatus.hWndToolTip, TTM_SETMAXTIPWIDTH, 0, min(width, maxWidth));
			} else {
				static const RECT margins = {0, 0, 0, 0};
				SendMessage(toolTipStatus.hWndToolTip, TTM_SETMARGIN, 0, reinterpret_cast<LPARAM>(&margins));
				toolTipStatus.inplaceToolTip = TRUE;
				SendMessage(toolTipStatus.hWndToolTip, TTM_SETMAXTIPWIDTH, 0, -1);
			}

			Str_SetPtrW(&toolTipStatus.pCurrentToolTipW, pText);
		}
	//}
	if(cancelToolTip == VARIANT_FALSE) {
		pDetails->lpszText = toolTipStatus.pCurrentToolTipW;
	}

	return 0;
}

LRESULT ListBox::OnToolTipShowNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled)
{
	if(toolTipStatus.toolTipItem != -1) {
		if(toolTipStatus.inplaceToolTip) {
			//LPDWORD pStyle = reinterpret_cast<LPDWORD>(pNotificationDetails + 1);
			RECT labelRectangle = {0};
			if(SendMessage(LB_GETITEMRECT, toolTipStatus.toolTipItem, reinterpret_cast<LPARAM>(&labelRectangle)) != LB_ERR) {
				SendMessage(toolTipStatus.hWndToolTip, TTM_ADJUSTRECT, TRUE, reinterpret_cast<LPARAM>(&labelRectangle));
				if(!(GetStyle() & LBS_MULTICOLUMN)) {
					OffsetRect(&labelRectangle, 2, 0);
				}
				MapWindowPoints(HWND_DESKTOP, &labelRectangle);
				::SetWindowPos(toolTipStatus.hWndToolTip, HWND_TOP, labelRectangle.left, labelRectangle.top, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE | SWP_HIDEWINDOW);
				//*pStyle |= TTS_NOANIMATE;
				return TRUE;
			}
		}
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ListBox::OnSetTabStops(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr && !IsInDesignMode()) {
		#ifdef USE_STL
			properties.tabStops.clear();
		#else
			properties.tabStops.RemoveAll();
		#endif
		if(wParam == 0) {
			properties.tabWidthInDTUs = 32;

			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			CDC memoryDC;
			memoryDC.CreateCompatibleDC();
			HFONT hPreviousFont = memoryDC.SelectFont(hFont);
			SIZE textExtent = {0};
			memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
			int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
			memoryDC.SelectFont(hPreviousFont);
			int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

			// convert to pixels
			properties.tabWidthInPixels = 2 * properties.tabWidthInDTUs * averageCharWidthOfUsedFont / averageCharWidthOfSystemFont;
		} else if(wParam == 1) {
			properties.tabWidthInDTUs = *reinterpret_cast<PUINT>(lParam);
			#ifdef USE_STL
				properties.tabStops.push_back(properties.tabWidthInDTUs);
			#else
				properties.tabStops.Add(properties.tabWidthInDTUs);
			#endif

			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			CDC memoryDC;
			memoryDC.CreateCompatibleDC();
			HFONT hPreviousFont = memoryDC.SelectFont(hFont);
			SIZE textExtent = {0};
			memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
			int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
			memoryDC.SelectFont(hPreviousFont);
			int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

			// convert to pixels
			properties.tabWidthInPixels = 2 * properties.tabWidthInDTUs * averageCharWidthOfUsedFont / averageCharWidthOfSystemFont;
		} else {
			for(UINT i = 0; i < wParam; ++i) {
				#ifdef USE_STL
					properties.tabStops.push_back(reinterpret_cast<PUINT>(lParam)[i]);
				#else
					properties.tabStops.Add(reinterpret_cast<PUINT>(lParam)[i]);
				#endif
			}
		}
	}
	return lr;
}

LRESULT ListBox::OnReflectedErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_OutOfMemory();
	wasHandled = FALSE;
	return 0;
}
//
//LRESULT ListBox::OnReflectedSelCancel(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
//{
//	ATLASSERT(FALSE);
//	wasHandled = FALSE;
//	return 0;
//}

LRESULT ListBox::OnReflectedSelChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	int newCaretItem = static_cast<int>(SendMessage(LB_GETCARETINDEX, 0, 0));
	if(newCaretItem != currentCaretItem) {
		Raise_CaretChanged(currentCaretItem, newCaretItem);
	}
	Raise_SelectionChanged();

	wasHandled = FALSE;
	return 0;
}

LRESULT ListBox::OnReflectedSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(*this);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(*this, ime);
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}


void ListBox::ApplyFont(void)
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


inline HRESULT ListBox::Raise_AbortedDrag(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_AbortedDrag();
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_CaretChanged(int previousCaretItem, int newCaretItem)
{
	currentCaretItem = newCaretItem;
	//if(m_nFreezeEvents == 0) {
		CComPtr<IListBoxItem> pLBPreviousCaretItem = ClassFactory::InitListItem(previousCaretItem, this);
		CComPtr<IListBoxItem> pLBNewCaretItem = ClassFactory::InitListItem(newCaretItem, this);
		return Fire_CaretChanged(pLBPreviousCaretItem, pLBNewCaretItem);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	mouseStatus.lastClickedItem = HitTest(x, y, &hitTestDetails);
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(mouseStatus.lastClickedItem, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_Click(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_CompareItems(IListBoxItem* pFirstItem, IListBoxItem* pSecondItem, LONG locale, CompareResultConstants* pResult)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CompareItems(pFirstItem, pSecondItem, locale, pResult);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		BOOL dontUsePosition = FALSE;
		LONG itemIndex = -1;
		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		if(x == -1 && y == -1) {
			// the event was caused by the keyboard
			dontUsePosition = TRUE;
			if(properties.processContextMenuKeys) {
				// retrieve the caret item and propose its rectangle's middle as the menu's position
				CRect clientRectangle;
				GetClientRect(&clientRectangle);
				CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
				hitTestDetails = htNotOverItem;

				itemIndex = static_cast<LONG>(SendMessage(LB_GETCARETINDEX, 0, 0));
				if(itemIndex >= 0) {
					CRect itemRectangle;
					if(SendMessage(LB_GETITEMRECT, itemIndex, reinterpret_cast<LPARAM>(&itemRectangle)) != LB_ERR) {
						centerPoint = itemRectangle.CenterPoint();
						x = centerPoint.x;
						y = centerPoint.y;
						dontUsePosition = FALSE;
						hitTestDetails = static_cast<HitTestConstants>(0);
					}
				}
			} else {
				return S_OK;
			}
		}

		if(!dontUsePosition) {
			itemIndex = HitTest(x, y, &hitTestDetails);
		}
		CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

		return Fire_ContextMenu(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(x, y, &hitTestDetails);
	if(itemIndex != mouseStatus.lastClickedItem) {
		itemIndex = -1;
	}
	mouseStatus.lastClickedItem = -1;
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_DestroyedControlWindow(HWND hWnd)
{
	KillTimer(timers.ID_REDRAW);
	RemoveItemFromItemContainers(-1);
	#ifdef USE_STL
		itemIDs.clear();
	#else
		itemIDs.RemoveAll();
	#endif
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(dragDropStatus.hDragImageList) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = HitTest(x, y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IListBoxItem* pDropTarget = NULL;
	ClassFactory::InitListItem(dropTarget, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtSelection, NULL, &yItem);
		yToItemTop = y - yItem;
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePosition(x, y);
		CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePosition);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePosition);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePosition.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePosition.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePosition.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePosition.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_DragMouseMove(&pDropTarget, button, shift, x, y, yToItemTop, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
	//}

	if(pDropTarget) {
		dropTarget = -1;
		pDropTarget->get_Index(reinterpret_cast<LONG*>(&dropTarget));
		// we don't want to produce a mem-leak...
		pDropTarget->Release();
	} else {
		dropTarget = -1;
	}
	dragDropStatus.lastDropTarget = dropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
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
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ListBox::Raise_Drop(void)
{
	//if(m_nFreezeEvents == 0) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);

		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
		/*TODO: if(dropTarget == -1) {
			if(hitTestDetails & htBelowLastItem) {
				dropTarget = <last item>;
			}
		}*/
		CComPtr<IListBoxItem> pDropTarget = ClassFactory::InitListItem(dropTarget, this);

		LONG yToItemTop = 0;
		if(pDropTarget) {
			LONG yItem = 0;
			pDropTarget->get_Index(reinterpret_cast<LONG*>(&dropTarget));
			yToItemTop = mousePosition.y - yItem;
		}

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		return Fire_Drop(pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_FreeItemData(IListBoxItem* pListItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeItemData(pListItem);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_InsertedItem(IListBoxItem* pListItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedItem(pListItem);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_InsertingItem(IVirtualListBoxItem* pListItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingItem(pListItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ItemBeginDrag(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemBeginDrag(pListItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ItemBeginRDrag(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemBeginRDrag(pListItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ItemGetDisplayInfo(IListBoxItem* pListItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pOverlayIndex)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetDisplayInfo(pListItem, requestedInfo, pIconIndex, pOverlayIndex);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ItemGetInfoTipText(IListBoxItem* pListItem, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetInfoTipText(pListItem, maxInfoTipLength, pInfoTipText, pAbortToolTip);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ItemMouseEnter(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_ItemMouseEnter(pListItem, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ListBox::Raise_ItemMouseLeave(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_ItemMouseLeave(pListItem, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ListBox::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	mouseStatus.lastClickedItem = HitTest(x, y, &hitTestDetails);
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(mouseStatus.lastClickedItem, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(x, y, &hitTestDetails);
	if(itemIndex != mouseStatus.lastClickedItem) {
		itemIndex = -1;
	}
	mouseStatus.lastClickedItem = -1;
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_MeasureItem(IListBoxItem* pListItem, OLE_YSIZE_PIXELS* pItemHeight)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MeasureItem(pListItem, pItemHeight);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		mouseStatus.StoreClickCandidate(button);
		SetCapture();

		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = HitTest(x, y, &hitTestDetails);
		CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

		return Fire_MouseDown(pLBItem, button, shift, x, y, hitTestDetails);
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
		mouseStatus.StoreClickCandidate(button);
		SetCapture();
		return S_OK;
	}
}

inline HRESULT ListBox::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(x, y, &hitTestDetails);
	itemUnderMouse = itemIndex;
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_MouseEnter(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	if(pLBItem) {
		Raise_ItemMouseEnter(pLBItem, button, shift, x, y, hitTestDetails);
	}
	return hr;
}

inline HRESULT ListBox::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = HitTest(x, y, &hitTestDetails);
		CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

		return Fire_MouseHover(pLBItem, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ListBox::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(x, y, &hitTestDetails);

	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemUnderMouse, this);
	if(pLBItem) {
		Raise_ItemMouseLeave(pLBItem, button, shift, x, y, hitTestDetails);
	}
	itemUnderMouse = -1;

	mouseStatus.LeaveControl();

	//if(m_nFreezeEvents == 0) {
		pLBItem = ClassFactory::InitListItem(itemIndex, this);
		return Fire_MouseLeave(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(x, y, &hitTestDetails);
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

	// Do we move over another item than before?
	if(itemIndex != itemUnderMouse) {
		CComPtr<IListBoxItem> pPrevLBItem = ClassFactory::InitListItem(itemUnderMouse, this);
		if(pPrevLBItem) {
			Raise_ItemMouseLeave(pPrevLBItem, button, shift, x, y, hitTestDetails);
		}
		itemUnderMouse = itemIndex;
		if(pLBItem) {
			Raise_ItemMouseEnter(pLBItem, button, shift, x, y, hitTestDetails);
		}
	}
	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = HitTest(x, y, &hitTestDetails);
		CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

		hr = Fire_MouseUp(pLBItem, button, shift, x, y, hitTestDetails);
	}

	if(mouseStatus.IsClickCandidate(button)) {
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
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(GetCapture() == *this) {
			ReleaseCapture();
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT ListBox::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(x, y, &hitTestDetails);
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(pLBItem, button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IListBoxItem* pDropTarget = NULL;
	ClassFactory::InitListItem(dropTarget, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

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
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails);
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT ListBox::Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IListBoxItem* pDropTarget = NULL;
	ClassFactory::InitListItem(dropTarget, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtSelection, NULL, &yItem);
		yToItemTop = mousePosition.y - yItem;
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone;
		GetClientRect(&noScrollZone);
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
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
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

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
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
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ListBox::Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_OLEDragLeave(void)
{
	KillTimer(timers.ID_DRAGSCROLL);
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	CComPtr<IListBoxItem> pDropTarget = ClassFactory::InitListItem(dropTarget, this);

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtSelection, NULL, &yItem);
		yToItemTop = dragDropStatus.lastMousePosition.y - yItem;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, yToItemTop, hitTestDetails);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();
	return hr;
}

inline HRESULT ListBox::Raise_OLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IListBoxItem* pDropTarget = NULL;
	ClassFactory::InitListItem(dropTarget, this, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG yToItemTop = 0;
	if(pDropTarget) {
		LONG yItem = 0;
		pDropTarget->GetRectangle(irtSelection, NULL, &yItem);
		yToItemTop = mousePosition.y - yItem;
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone;
		GetClientRect(&noScrollZone);
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
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
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

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
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
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ListBox::Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
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
HRESULT ListBox::Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT ListBox::Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLESetData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_OLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_OutOfMemory(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OutOfMemory();
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_OwnerDrawItem(IListBoxItem* pListItem, OwnerDrawActionConstants requiredAction, OwnerDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OwnerDrawItem(pListItem, requiredAction, itemState, hDC, pDrawingRectangle);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ProcessCharacterInput(SHORT keyAscii, SHORT shift, LONG* pItemToSelect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ProcessCharacterInput(keyAscii, shift, pItemToSelect);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ProcessKeyStroke(SHORT keyCode, SHORT shift, LONG* pItemToSelect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ProcessKeyStroke(keyCode, shift, pItemToSelect);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	mouseStatus.lastClickedItem = HitTest(x, y, &hitTestDetails);
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(mouseStatus.lastClickedItem, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(x, y, &hitTestDetails);
	if(itemIndex != mouseStatus.lastClickedItem) {
		itemIndex = -1;
	}
	mouseStatus.lastClickedItem = -1;
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_RecreatedControlWindow(HWND hWnd)
{
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

inline HRESULT ListBox::Raise_RemovedItem(IVirtualListBoxItem* pListItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedItem(pListItem);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_RemovingItem(IListBoxItem* pListItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingItem(pListItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_SelectionChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectionChanged();
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	mouseStatus.lastClickedItem = HitTest(x, y, &hitTestDetails);
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(mouseStatus.lastClickedItem, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ListBox::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = HitTest(x, y, &hitTestDetails);
	if(itemIndex != mouseStatus.lastClickedItem) {
		itemIndex = -1;
	}
	mouseStatus.lastClickedItem = -1;
	CComPtr<IListBoxItem> pLBItem = ClassFactory::InitListItem(itemIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(pLBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}


void ListBox::RecreateControlWindow(void)
{
	if(m_bInPlaceActive && !flags.dontRecreate) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate(isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE);
		// if we don't invalidate, we'll have drawing glitches on Vista
		Invalidate();
	}
}

IMEModeConstants ListBox::GetCurrentIMEContextMode(HWND hWnd)
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

void ListBox::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
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

IMEModeConstants ListBox::GetEffectiveIMEMode(void)
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

DWORD ListBox::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD ListBox::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_HSCROLL | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}
	if(!properties.allowItemSelection) {
		style |= LBS_NOSEL;
	}
	if(properties.alwaysShowVerticalScrollBar) {
		style |= LBS_DISABLENOSCROLL;
	}
	if(!(properties.disabledEvents & deProcessKeyboardInput)) {
		style |= LBS_WANTKEYBOARDINPUT;
	}
	if(properties.hasStrings) {
		style |= LBS_HASSTRINGS;
	}
	if(!properties.integralHeight) {
		style |= LBS_NOINTEGRALHEIGHT;
	}
	if(properties.multiColumn) {
		style |= LBS_MULTICOLUMN;
	}
	switch(properties.multiSelect) {
		case msNormal:
			style |= LBS_EXTENDEDSEL;
			break;
		case msSelectByClick:
			style |= LBS_MULTIPLESEL;
			break;
	}
	switch(properties.ownerDrawItems) {
		case odiOwnerDrawFixedHeight:
			style |= LBS_OWNERDRAWFIXED;
			break;
		case odiOwnerDrawVariableHeight:
			style |= LBS_OWNERDRAWVARIABLE;
			break;
	}
	if(properties.processTabs) {
		style |= LBS_USETABSTOPS;
	}
	if(properties.sorted) {
		style |= LBS_SORT;
	}
	if(properties.virtualMode) {
		style |= LBS_NODATA;
	}
	return style;
}

void ListBox::SendConfigurationMessages(void)
{
	SendMessage(LB_SETCOLUMNWIDTH, properties.columnWidth, 0);
	SendMessage(LB_SETLOCALE, properties.locale, 0);
	SendMessage(LB_SETHORIZONTALEXTENT, properties.scrollableWidth, 0);

	ApplyFont();

	#ifdef USE_STL
		if(properties.tabStops.empty()) {
	#else
		if(properties.tabStops.IsEmpty()) {
	#endif
		if(properties.tabWidthInPixels == -1) {
			SendMessage(LB_SETTABSTOPS, 0, 0);
		} else {
			UINT distance = properties.tabWidthInDTUs;
			SendMessage(LB_SETTABSTOPS, 1, reinterpret_cast<LPARAM>(&distance));
		}
	} else {
		#ifdef USE_STL
			size_t numberOfStops = properties.tabStops.size();
		#else
			size_t numberOfStops = properties.tabStops.GetCount();
		#endif
		PUINT pTabStops = new UINT[numberOfStops];
		for(size_t i = 0; i < numberOfStops; ++i) {
			pTabStops[i] = properties.tabStops[i];
		}
		SendMessage(LB_SETTABSTOPS, numberOfStops, reinterpret_cast<LPARAM>(pTabStops));
		delete[] pTabStops;
	}
	if(properties.setItemHeight != -1) {
		if(!(GetStyle() & LBS_OWNERDRAWVARIABLE)) {
			SendMessage(LB_SETITEMHEIGHT, 0, properties.setItemHeight);
		}
	}

	if(properties.toolTips) {
		toolTipStatus.hWndToolTip = CreateWindowEx(0, TOOLTIPS_CLASS, NULL, WS_POPUP | TTS_NOPREFIX, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, *this, NULL, NULL, NULL);
		if(toolTipStatus.hWndToolTip) {
			TOOLINFO toolInfo = {0};
			toolInfo.cbSize = sizeof(TOOLINFO);
			toolInfo.uFlags = TTF_TRANSPARENT;
			toolInfo.hwnd = *this;
			toolInfo.lpszText = LPSTR_TEXTCALLBACK;
			GetClientRect(&toolInfo.rect);
			SendMessage(toolTipStatus.hWndToolTip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&toolInfo));
			SendMessage(toolTipStatus.hWndToolTip, WM_SETFONT, reinterpret_cast<WPARAM>(GetFont()), MAKELPARAM(TRUE, 0));
		}
	}
}

HCURSOR ListBox::MousePointerConst2hCursor(MousePointerConstants mousePointer)
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

int ListBox::IDToItemIndex(LONG ID)
{
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
	return -1;
}

LONG ListBox::ItemIndexToID(int itemIndex)
{
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

LONG ListBox::GetNewItemID(void)
{
	return ++lastItemID;
}

void ListBox::RegisterItemContainer(IItemContainer* pContainer)
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

void ListBox::DeregisterItemContainer(DWORD containerID)
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

void ListBox::RemoveItemFromItemContainers(LONG itemIdentifier)
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

void ListBox::UpdateItemIDInItemContainers(LONG oldItemID, LONG newItemID)
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

int ListBox::HitTest(LONG x, LONG y, HitTestConstants* pFlags, BOOL autoScroll/* = FALSE*/)
{
	ATLASSERT(IsWindow());

	int itemIndex = -1;

	POINT pt = {x, y};
	ClientToScreen(&pt);
	CRect rc;
	GetWindowRect(&rc);
	if(rc.PtInRect(pt)) {
		*pFlags = htNotOverItem;
		itemIndex = LBItemFromPt(*this, pt, autoScroll);
		if(itemIndex != -1) {
			*pFlags = htItem;
		}
	} else {
		if(autoScroll) {
			LBItemFromPt(*this, pt, TRUE);
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

BOOL ListBox::IsItemTruncated(int itemIndex, LPTSTR pLabel, int bufferSize)
{
	BOOL containsStrings;
	DWORD style = GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		containsStrings = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		containsStrings = TRUE;
	}
	if(!containsStrings) {
		return FALSE;
	}

	int textLength = static_cast<int>(SendMessage(LB_GETTEXTLEN, itemIndex, 0));
	if(textLength == LB_ERR) {
		return FALSE;
	}
	LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
	if(!pBuffer) {
		return FALSE;
	}
	SendMessage(LB_GETTEXT, itemIndex, reinterpret_cast<LPARAM>(pBuffer));

	BOOL isTruncated = FALSE;

	HDC hCompatibleDC = GetDC();
	CDC memoryDC;
	memoryDC.CreateCompatibleDC(hCompatibleDC);
	HFONT hPreviousFont = memoryDC.SelectFont(GetFont());

	CRect textRectangle;
	if(flags.usingThemes) {
		CTheme themingEngine;
		themingEngine.OpenThemeData(NULL, VSCLASS_LISTBOX);
		if(!themingEngine.IsThemeNull()) {
			CT2W converter(pBuffer);
			LPWSTR p = converter;
			themingEngine.GetThemeTextExtent(memoryDC, LBCP_ITEM, 0, p, -1, DT_CALCRECT | DT_SINGLELINE | DT_NOPREFIX, NULL, &textRectangle);
			MARGINS margins = {0};
			themingEngine.GetThemeMargins(memoryDC, LBCP_ITEM, 0, TMT_CONTENTMARGINS, &textRectangle, &margins);
			textRectangle.left -= margins.cxLeftWidth;
			textRectangle.right += margins.cxRightWidth;
			textRectangle.top -= margins.cyTopHeight;
			textRectangle.bottom += margins.cyBottomHeight;
		} else {
			memoryDC.DrawText(pBuffer, -1, &textRectangle, DT_CALCRECT | DT_SINGLELINE | DT_NOPREFIX);
		}
	} else {
		memoryDC.DrawText(pBuffer, -1, &textRectangle, DT_CALCRECT | DT_SINGLELINE | DT_NOPREFIX);
	}
	if(!(GetStyle() & LBS_MULTICOLUMN)) {
		OffsetRect(&textRectangle, 2, 0);
	}

	memoryDC.SelectFont(hPreviousFont);
	ReleaseDC(hCompatibleDC);

	SCROLLINFO scrollInfo = {0};
	scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_POS;
	GetScrollInfo(SB_HORZ, &scrollInfo);
	textRectangle.OffsetRect(-scrollInfo.nPos, 0);

	CRect clientRectangle;
	GetClientRect(&clientRectangle);
	MapWindowPoints(HWND_DESKTOP, &clientRectangle);
	MapWindowPoints(HWND_DESKTOP, &textRectangle);
	if(textRectangle.left < clientRectangle.left || textRectangle.right > clientRectangle.right) {
		isTruncated = TRUE;
	}

	if(isTruncated) {
		lstrcpyn(pLabel, pBuffer, bufferSize);
	}

	HeapFree(GetProcessHeap(), 0, pBuffer);
	return isTruncated;
}

BOOL ListBox::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void ListBox::AutoScroll(void)
{
	LONG realScrollTimeBase = properties.dragScrollTimeBase;
	if(realScrollTimeBase == -1) {
		realScrollTimeBase = GetDoubleClickTime() / 4;
	}

	if((dragDropStatus.autoScrolling.currentHScrollVelocity != 0) && ((GetStyle() & WS_HSCROLL) == WS_HSCROLL)) {
		if(dragDropStatus.autoScrolling.currentHScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll to the left?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Left) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll left
					dragDropStatus.autoScrolling.lastScroll_Left = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINELEFT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll to the right?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Right) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll right
					dragDropStatus.autoScrolling.lastScroll_Right = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINERIGHT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}

	if((dragDropStatus.autoScrolling.currentVScrollVelocity != 0) && ((GetStyle() & WS_VSCROLL) == WS_VSCROLL)) {
		if(dragDropStatus.autoScrolling.currentVScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll upwardly?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Up) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll up
					dragDropStatus.autoScrolling.lastScroll_Up = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEUP, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll downwards?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Down) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll down
					dragDropStatus.autoScrolling.lastScroll_Down = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEDOWN, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}
}

void ListBox::DrawInsertionMark(CDCHandle targetDC)
{
	if(insertMark.itemIndex != -1 && !insertMark.hidden) {
		CPen pen;
		pen.CreatePen(PS_SOLID, 1, insertMark.color);
		HPEN hPreviousPen = targetDC.SelectPen(pen);

		RECT itemBoundingRectangle = {0};
		SendMessage(LB_GETITEMRECT, insertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
		RECT insertMarkRect = {0};
		if(insertMark.afterItem) {
			insertMarkRect.top = itemBoundingRectangle.bottom - 3;
			insertMarkRect.bottom = itemBoundingRectangle.bottom + 3;
		} else {
			insertMarkRect.top = itemBoundingRectangle.top - 3;
			insertMarkRect.bottom = itemBoundingRectangle.top + 3;
		}
		insertMarkRect.left = itemBoundingRectangle.left;
		insertMarkRect.right = itemBoundingRectangle.right - 1;

		// draw the main line
		targetDC.MoveTo(insertMarkRect.left, insertMarkRect.top + 2);
		targetDC.LineTo(insertMarkRect.right, insertMarkRect.top + 2);
		targetDC.MoveTo(insertMarkRect.left, insertMarkRect.top + 3);
		targetDC.LineTo(insertMarkRect.right, insertMarkRect.top + 3);

		// draw the ends
		targetDC.MoveTo(insertMarkRect.left, insertMarkRect.top);
		targetDC.LineTo(insertMarkRect.left, insertMarkRect.bottom);
		targetDC.MoveTo(insertMarkRect.left + 1, insertMarkRect.top + 1);
		targetDC.LineTo(insertMarkRect.left + 1, insertMarkRect.bottom - 1);
		targetDC.MoveTo(insertMarkRect.right - 1, insertMarkRect.top + 1);
		targetDC.LineTo(insertMarkRect.right - 1, insertMarkRect.bottom - 1);
		targetDC.MoveTo(insertMarkRect.right, insertMarkRect.top);
		targetDC.LineTo(insertMarkRect.right, insertMarkRect.bottom);

		targetDC.SelectPen(hPreviousPen);
	}
}

BOOL ListBox::IsLeftMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	}
}

BOOL ListBox::IsRightMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	}
}


void ListBox::BufferItemData(LPARAM itemData)
{
	#ifdef USE_STL
		properties.itemDataOfInsertedItems.push(itemData);
	#else
		properties.itemDataOfInsertedItems.AddTail(itemData);
	#endif
}

void ListBox::EnterSilentItemInsertionSection(void)
{
	++flags.silentItemInsertions;
}

void ListBox::LeaveSilentItemInsertionSection(void)
{
	--flags.silentItemInsertions;
}

void ListBox::EnterSilentItemDeletionSection(void)
{
	++flags.silentItemDeletions;
}

void ListBox::LeaveSilentItemDeletionSection(void)
{
	--flags.silentItemDeletions;
}

void ListBox::SetItemID(int itemIndex, LONG itemID)
{
	LONG currentID = ItemIndexToID(itemIndex);
	#ifdef USE_STL
		if(itemIndex >= 0 && itemIndex < static_cast<int>(itemIDs.size())) {
			itemIDs[itemIndex] = itemID;
		}
	#else
		if(itemIndex >= 0 && itemIndex < static_cast<int>(itemIDs.GetCount())) {
			itemIDs[itemIndex] = itemID;
		}
	#endif
	UpdateItemIDInItemContainers(currentID, itemID);
}

void ListBox::DecrementNextItemID(void)
{
	--lastItemID;
}


HRESULT ListBox::CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing)
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