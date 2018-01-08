// ComboBox.cpp: Superclasses ComboBox.

#include "stdafx.h"
#include "ComboBox.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
IMEModeConstants ComboBox::IMEFlags::chineseIMETable[10] = {
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

IMEModeConstants ComboBox::IMEFlags::japaneseIMETable[10] = {
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

IMEModeConstants ComboBox::IMEFlags::koreanIMETable[10] = {
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

FONTDESC ComboBox::Properties::FontProperty::defaultFont = {
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
ComboBox::ComboBox() :
    containedListBox(WC_LISTBOX, this, 1),
    containedEdit(WC_EDIT, this, 2)
{
	properties.font.InitializePropertyWatcher(this, DISPID_CB_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_CB_MOUSEICON);

	SIZEL size = {100, 20};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
	lastItemID = 0;
	currentSelectedItem = -1;
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
}
#pragma warning(pop)


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ComboBox::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IComboBox, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ComboBox::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<ComboBox>::Load(pPropertyBag, pErrorLog);
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

		bstr.Empty();
		hr = pPropertyBag->Read(OLESTR("Text"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_Text(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP ComboBox::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<ComboBox>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.cueBanner.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("CueBanner"), &propertyValue);
		VariantClear(&propertyValue);

		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.text.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("Text"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP ComboBox::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 26 VT_I4 properties...
	pSize->QuadPart += 26 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 13 VT_BOOL properties...
	pSize->QuadPart += 13 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 2 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.cueBanner.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.text.ByteLength() + sizeof(OLECHAR);

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

STDMETHODIMP ComboBox::Load(LPSTREAM pStream)
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
	if(subSignature != 0x01010101/*4x 0x01 (-> ComboBox)*/) {
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
	VARIANT_BOOL acceptNumbersOnly = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL autoHorizontalScrolling = propertyValue.boolVal;
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
	CharacterConversionConstants characterConversion = static_cast<CharacterConversionConstants>(propertyValue.lVal);
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL integralHeight = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS itemHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL listAlwaysShowVerticalScrollBar = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR listBackColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG listDragScrollTimeBase = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR listForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS listHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR listInsertMarkColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS listScrollableWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS listWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG locale = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maxTextLength = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG minVisibleItems = propertyValue.lVal;

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
	OwnerDrawItemsConstants ownerDrawItems = static_cast<OwnerDrawItemsConstants>(propertyValue.lVal);
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
	VARIANT_BOOL sorted = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	StyleConstants style = static_cast<StyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;


	BOOL b = flags.dontRecreate;
	flags.dontRecreate = TRUE;
	hr = put_AcceptNumbersOnly(acceptNumbersOnly);
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
	hr = put_CueBanner(cueBanner);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
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
	hr = put_HoverTime(hoverTime);
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
	hr = put_ListBackColor(listBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ListDragScrollTimeBase(listDragScrollTimeBase);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ListForeColor(listForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ListHeight(listHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ListInsertMarkColor(listInsertMarkColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ListScrollableWidth(listScrollableWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ListWidth(listWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Locale(locale);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxTextLength(maxTextLength);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MinVisibleItems(minVisibleItems);
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
	hr = put_SelectionFieldHeight(selectionFieldHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoHorizontalScrolling(autoHorizontalScrolling);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CharacterConversion(characterConversion);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DoOEMConversion(doOEMConversion);
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
	hr = put_ListAlwaysShowVerticalScrollBar(listAlwaysShowVerticalScrollBar);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OwnerDrawItems(ownerDrawItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Sorted(sorted);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Style(style);
	if(FAILED(hr)) {
		return hr;
	}
	flags.dontRecreate = b;
	RecreateControlWindow();

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ComboBox::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
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
	LONG subSignature = 0x01010101/*4x 0x01 (-> ComboBox)*/;
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.acceptNumbersOnly);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoHorizontalScrolling);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.characterConversion;
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
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
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
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.listAlwaysShowVerticalScrollBar);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.listBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.listDragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.listForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.listHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.listInsertMarkColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.listScrollableWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.listWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.locale;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maxTextLength;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.minVisibleItems;
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
	propertyValue.lVal = properties.ownerDrawItems;
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.sorted);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.style;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.text.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND ComboBox::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_STANDARD_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<ComboBox>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT ComboBox::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("ComboBox ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void ComboBox::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP ComboBox::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<ComboBox>::IOleObject_SetClientSite(pClientSite);

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

STDMETHODIMP ComboBox::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL ComboBox::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
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
	return CComControl<ComboBox>::PreTranslateAccelerator(pMessage, hReturnValue);
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP ComboBox::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	HWND hWnd = WindowFromPoint(buffer);
	if(hWnd == containedListBox) {
		Raise_ListOLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
		if(dragDropStatus.pDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragEnter(containedListBox, pDataObject, &buffer, *pEffect);
		}
	} else {
		Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
		if(dragDropStatus.pDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::DragLeave(void)
{
	if(dragDropStatus.isOverListBox) {
		Raise_ListOLEDragLeave(FALSE);
	} else {
		Raise_OLEDragLeave(FALSE);
	}
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP ComboBox::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	HWND hWnd = WindowFromPoint(buffer);
	if(hWnd == containedListBox) {
		if(dragDropStatus.isOverListBox) {
			Raise_ListOLEDragMouseMove(pEffect, keyState, mousePosition);
		} else {
			// we've left the combo box and entered the list box
			Raise_OLEDragLeave(TRUE);
			if(dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragLeave();
			}

			DROPDESCRIPTION newDropDescription;
			ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
			IDataObject_GetDropDescription(pDataObject, newDropDescription);
			if(memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
				InvalidateDragWindow(pDataObject);
				oldDropDescription = newDropDescription;
			}

			Raise_ListOLEDragEnter(NULL, pEffect, keyState, mousePosition);
			if(dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragEnter(containedListBox, pDataObject, &buffer, *pEffect);
			}
		}
	} else {
		if(dragDropStatus.isOverListBox) {
			// we've left the list box and entered the combo box
			Raise_ListOLEDragLeave(TRUE);
			if(dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragLeave();
			}

			DROPDESCRIPTION newDropDescription;
			ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
			IDataObject_GetDropDescription(pDataObject, newDropDescription);
			if(memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
				InvalidateDragWindow(pDataObject);
				oldDropDescription = newDropDescription;
			}

			Raise_OLEDragEnter(NULL, pEffect, keyState, mousePosition);
			if(dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			}
		} else {
			Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
		}
	}
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

STDMETHODIMP ComboBox::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	HWND hWnd = WindowFromPoint(buffer);
	if(hWnd == containedListBox) {
		Raise_ListOLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	} else {
		Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	}
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
// implementation of ICategorizeProperties
STDMETHODIMP ComboBox::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
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

STDMETHODIMP ComboBox::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_CB_DROPDOWNBUTTONOBJECTSTATE:
			*pCategory = PROPCAT_Accessibility;
			return S_OK;
			break;
		case DISPID_CB_APPEARANCE:
		case DISPID_CB_BORDERSTYLE:
		case DISPID_CB_CUEBANNER:
		case DISPID_CB_ISDROPPEDDOWN:
		case DISPID_CB_ITEMHEIGHT:
		case DISPID_CB_MOUSEICON:
		case DISPID_CB_MOUSEPOINTER:
		case DISPID_CB_SELECTIONFIELDHEIGHT:
		case DISPID_CB_STYLE:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_CB_ACCEPTNUMBERSONLY:
		case DISPID_CB_AUTOHORIZONTALSCROLLING:
		case DISPID_CB_CHARACTERCONVERSION:
		case DISPID_CB_DISABLEDEVENTS:
		case DISPID_CB_DONTREDRAW:
		case DISPID_CB_DOOEMCONVERSION:
		case DISPID_CB_DROPDOWNKEY:
		case DISPID_CB_HOVERTIME:
		case DISPID_CB_IMEMODE:
		case DISPID_CB_LISTALWAYSSHOWVERTICALSCROLLBAR:
		case DISPID_CB_LISTHEIGHT:
		case DISPID_CB_LISTSCROLLABLEWIDTH:
		case DISPID_CB_LISTWIDTH:
		case DISPID_CB_MAXTEXTLENGTH:
		case DISPID_CB_MINVISIBLEITEMS:
		case DISPID_CB_OWNERDRAWITEMS:
		case DISPID_CB_PROCESSCONTEXTMENUKEYS:
		case DISPID_CB_RIGHTTOLEFT:
		case DISPID_CB_SORTED:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_CB_BACKCOLOR:
		case DISPID_CB_FORECOLOR:
		case DISPID_CB_LISTBACKCOLOR:
		case DISPID_CB_LISTFORECOLOR:
		case DISPID_CB_LISTINSERTMARKCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_CB_APPID:
		case DISPID_CB_APPNAME:
		case DISPID_CB_APPSHORTNAME:
		case DISPID_CB_BUILD:
		case DISPID_CB_CHARSET:
		case DISPID_CB_HWND:
		case DISPID_CB_HWNDEDIT:
		case DISPID_CB_HWNDLISTBOX:
		case DISPID_CB_ISRELEASE:
		case DISPID_CB_PROGRAMMER:
		case DISPID_CB_TESTER:
		case DISPID_CB_TEXT:
		case DISPID_CB_TEXTLENGTH:
		case DISPID_CB_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_CB_DRAGDROPDOWNTIME:
		case DISPID_CB_LISTDRAGSCROLLTIMEBASE:
		case DISPID_CB_REGISTERFOROLEDRAGDROP:
		case DISPID_CB_SHOWDRAGIMAGE:
		case DISPID_CB_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_CB_FONT:
		case DISPID_CB_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_CB_COMBOITEMS:
		case DISPID_CB_FIRSTVISIBLEITEM:
		case DISPID_CB_HASSTRINGS:
		case DISPID_CB_INTEGRALHEIGHT:
		case DISPID_CB_SELECTEDITEM:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_CB_ENABLED:
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
CAtlString ComboBox::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ComboBox::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ComboBox::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ComboListBoxControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ComboBox::GetSpecialThanks(void)
{
	return TEXT("Wine Headquarters");
}

CAtlString ComboBox::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ComboBox::GetVersion(void)
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
STDMETHODIMP ComboBox::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_CB_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_CB_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_CB_CHARACTERCONVERSION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.characterConversion), description);
			break;
		case DISPID_CB_CUEBANNER:
		case DISPID_CB_TEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_CB_DROPDOWNKEY:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.dropDownKey), description);
			break;
		case DISPID_CB_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_CB_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_CB_OWNERDRAWITEMS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.ownerDrawItems), description);
			break;
		case DISPID_CB_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_CB_STYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.style), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<ComboBox>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ComboBox::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_CB_BORDERSTYLE:
		case DISPID_CB_DROPDOWNKEY:
			c = 2;
			break;
		case DISPID_CB_CHARACTERCONVERSION:
		case DISPID_CB_OWNERDRAWITEMS:
		case DISPID_CB_STYLE:
			c = 3;
			break;
		case DISPID_CB_APPEARANCE:
		case DISPID_CB_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_CB_IMEMODE:
			c = 12;
			break;
		case DISPID_CB_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<ComboBox>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_CB_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
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

STDMETHODIMP ComboBox::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_CB_APPEARANCE:
		case DISPID_CB_BORDERSTYLE:
		case DISPID_CB_CHARACTERCONVERSION:
		case DISPID_CB_DROPDOWNKEY:
		case DISPID_CB_IMEMODE:
		case DISPID_CB_MOUSEPOINTER:
		case DISPID_CB_OWNERDRAWITEMS:
		case DISPID_CB_RIGHTTOLEFT:
		case DISPID_CB_STYLE:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<ComboBox>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ComboBox::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_CB_CUEBANNER:
		case DISPID_CB_TEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<ComboBox>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ComboBox::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_CB_APPEARANCE:
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
				case aDefault:
					description = GetResStringWithNumber(IDP_APPEARANCEDEFAULT, aDefault);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_CB_BORDERSTYLE:
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
		case DISPID_CB_CHARACTERCONVERSION:
			switch(cookie) {
				case ccNone:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONNONE, ccNone);
					break;
				case ccLowerCase:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONLOWERCASE, ccLowerCase);
					break;
				case ccUpperCase:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONUPPERCASE, ccUpperCase);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_CB_DROPDOWNKEY:
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
		case DISPID_CB_IMEMODE:
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
		case DISPID_CB_MOUSEPOINTER:
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
		case DISPID_CB_OWNERDRAWITEMS:
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
		case DISPID_CB_RIGHTTOLEFT:
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
		case DISPID_CB_STYLE:
			switch(cookie) {
				case sComboDropDownList:
					description = GetResStringWithNumber(IDP_STYLECOMBODROPDOWNLIST, sComboDropDownList);
					break;
				case sDropDownList:
					description = GetResStringWithNumber(IDP_STYLEDROPDOWNLIST, sDropDownList);
					break;
				case sComboField:
					description = GetResStringWithNumber(IDP_STYLECOMBOFIELD, sComboField);
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
STDMETHODIMP ComboBox::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 5;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_StockColorPage;
		pPropertyPages->pElems[3] = CLSID_StockFontPage;
		pPropertyPages->pElems[4] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP ComboBox::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ComboBox>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ComboBox::EnumVerbs(IEnumOLEVERB** ppEnumerator)
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
STDMETHODIMP ComboBox::GetControlInfo(LPCONTROLINFO pControlInfo)
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

HRESULT ComboBox::DoVerbAbout(HWND hWndParent)
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

HRESULT ComboBox::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_CB_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT ComboBox::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP ComboBox::get_AcceptNumbersOnly(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.acceptNumbersOnly = ((containedEdit.GetStyle() & ES_NUMBER) == ES_NUMBER);
	}

	*pValue = BOOL2VARIANTBOOL(properties.acceptNumbersOnly);
	return S_OK;
}

STDMETHODIMP ComboBox::put_AcceptNumbersOnly(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_ACCEPTNUMBERSONLY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.acceptNumbersOnly != b) {
		properties.acceptNumbersOnly = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			if(properties.acceptNumbersOnly) {
				containedEdit.ModifyStyle(0, ES_NUMBER);
			} else {
				containedEdit.ModifyStyle(ES_NUMBER, 0);
			}
		}
		FireOnChanged(DISPID_CB_ACCEPTNUMBERSONLY);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_Appearance(AppearanceConstants* pValue)
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

STDMETHODIMP ComboBox::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_APPEARANCE);
	if(newValue < 0 || newValue > aDefault) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(newValue == aDefault && !IsInDesignMode() && IsWindow()) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

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
		FireOnChanged(DISPID_CB_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 16;
	return S_OK;
}

STDMETHODIMP ComboBox::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ComboListBoxControls");
	return S_OK;
}

STDMETHODIMP ComboBox::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"CBLCtls");
	return S_OK;
}

STDMETHODIMP ComboBox::get_AutoHorizontalScrolling(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.autoHorizontalScrolling = ((GetStyle() & CBS_AUTOHSCROLL) == CBS_AUTOHSCROLL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoHorizontalScrolling);
	return S_OK;
}

STDMETHODIMP ComboBox::put_AutoHorizontalScrolling(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_AUTOHORIZONTALSCROLLING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoHorizontalScrolling != b) {
		properties.autoHorizontalScrolling = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_AUTOHORIZONTALSCROLLING);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP ComboBox::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CB_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.Invalidate();
		}
		FireViewChange();
		FireOnChanged(DISPID_CB_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_BorderStyle(BorderStyleConstants* pValue)
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

STDMETHODIMP ComboBox::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_BORDERSTYLE);
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
		FireOnChanged(DISPID_CB_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ComboBox::get_CharacterConversion(CharacterConversionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, CharacterConversionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & CBS_LOWERCASE) {
			properties.characterConversion = ccLowerCase;
		} else if(style & CBS_UPPERCASE) {
			properties.characterConversion = ccUpperCase;
		} else {
			properties.characterConversion = ccNone;
		}
	}

	*pValue = properties.characterConversion;
	return S_OK;
}

STDMETHODIMP ComboBox::put_CharacterConversion(CharacterConversionConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_CHARACTERCONVERSION);
	if(properties.characterConversion != newValue) {
		properties.characterConversion = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_CHARACTERCONVERSION);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_CharSet(BSTR* pValue)
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

STDMETHODIMP ComboBox::get_ComboItems(IComboBoxItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IComboBoxItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	ClassFactory::InitComboItems(this, IID_IComboBoxItems, reinterpret_cast<LPUNKNOWN*>(ppItems));
	return S_OK;
}

STDMETHODIMP ComboBox::get_CueBanner(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		WCHAR pBuffer[1025];
		if(SendMessage(CB_GETCUEBANNER, reinterpret_cast<WPARAM>(pBuffer), 1024)) {
			properties.cueBanner = pBuffer;
		}
	}

	*pValue = properties.cueBanner.Copy();
	return S_OK;
}

STDMETHODIMP ComboBox::put_CueBanner(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_CB_CUEBANNER);
	if(properties.cueBanner != newValue) {
		properties.cueBanner.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(CB_SETCUEBANNER, 0, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
		}
		FireOnChanged(DISPID_CB_CUEBANNER);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ComboBox::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_DISABLEDEVENTS);
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
					if(containedEdit.IsWindow()) {
						trackingOptions.hwndTrack = containedEdit;
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
		FireOnChanged(DISPID_CB_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP ComboBox::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_CB_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_DoOEMConversion(VARIANT_BOOL* pValue)
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

STDMETHODIMP ComboBox::put_DoOEMConversion(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_DOOEMCONVERSION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.doOEMConversion != b) {
		properties.doOEMConversion = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_DOOEMCONVERSION);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_DragDropDownTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragDropDownTime;
	return S_OK;
}

STDMETHODIMP ComboBox::put_DragDropDownTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CB_DRAGDROPDOWNTIME);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragDropDownTime != newValue) {
		properties.dragDropDownTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CB_DRAGDROPDOWNTIME);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_DropDownButtonObjectState(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(*this, &controlInfo);
		*pValue = controlInfo.stateButton;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBox::get_DropDownKey(DropDownKeyConstants* pValue)
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

STDMETHODIMP ComboBox::put_DropDownKey(DropDownKeyConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_DROPDOWNKEY);
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
		FireOnChanged(DISPID_CB_DROPDOWNKEY);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_Enabled(VARIANT_BOOL* pValue)
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

STDMETHODIMP ComboBox::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_ENABLED);
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

		FireOnChanged(DISPID_CB_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_FirstVisibleItem(IComboBoxItem** ppFirstItem)
{
	ATLASSERT_POINTER(ppFirstItem, IComboBoxItem*);
	if(!ppFirstItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		ClassFactory::InitComboItem(static_cast<int>(SendMessage(CB_GETTOPINDEX, 0, 0)), this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFirstItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBox::putref_FirstVisibleItem(IComboBoxItem* pNewFirstItem)
{
	PUTPROPPROLOG(DISPID_CB_FIRSTVISIBLEITEM);
	HRESULT hr = E_FAIL;

	int newFirstItem = -1;
	if(pNewFirstItem) {
		LONG l = -1;
		pNewFirstItem->get_Index(&l);
		newFirstItem = l;
		// TODO: Shouldn't we AddRef' pNewFirstItem?
	}

	if(IsWindow()) {
		SendMessage(CB_SETTOPINDEX, newFirstItem, 0);
		hr = S_OK;
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_CB_FIRSTVISIBLEITEM);
	FireViewChange();
	return hr;
}

STDMETHODIMP ComboBox::get_Font(IFontDisp** ppFont)
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

STDMETHODIMP ComboBox::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_CB_FONT);
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
	FireOnChanged(DISPID_CB_FONT);
	return S_OK;
}

STDMETHODIMP ComboBox::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_CB_FONT);
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
	FireOnChanged(DISPID_CB_FONT);
	return S_OK;
}

STDMETHODIMP ComboBox::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP ComboBox::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CB_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.Invalidate();
		}
		FireViewChange();
		FireOnChanged(DISPID_CB_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_HasStrings(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.hasStrings = ((GetStyle() & CBS_HASSTRINGS) == CBS_HASSTRINGS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.hasStrings);
	return S_OK;
}

STDMETHODIMP ComboBox::put_HasStrings(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_HASSTRINGS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.hasStrings != b) {
		properties.hasStrings = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_HASSTRINGS);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP ComboBox::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CB_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CB_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP ComboBox::get_hWndEdit(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(containedEdit.m_hWnd);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_hWndListBox(OLE_HANDLE* pValue)
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

STDMETHODIMP ComboBox::get_IMEMode(IMEModeConstants* pValue)
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

STDMETHODIMP ComboBox::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			HWND h = GetFocus();
			if(h == *this || h == containedEdit || h == containedListBox) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(h, ime);
				}
			}
		}
		FireOnChanged(DISPID_CB_IMEMODE);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_IntegralHeight(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.integralHeight = !(GetStyle() & CBS_NOINTEGRALHEIGHT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.integralHeight);
	return S_OK;
}

STDMETHODIMP ComboBox::put_IntegralHeight(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_INTEGRALHEIGHT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.integralHeight != b) {
		properties.integralHeight = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_INTEGRALHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_IsDroppedDown(VARIANT_BOOL* pValue)
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

STDMETHODIMP ComboBox::get_IsRelease(VARIANT_BOOL* pValue)
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

STDMETHODIMP ComboBox::get_ItemHeight(OLE_YSIZE_PIXELS* pValue)
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

STDMETHODIMP ComboBox::put_ItemHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_CB_ITEMHEIGHT);
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
				if(!(GetStyle() & CBS_OWNERDRAWVARIABLE)) {
					SendMessage(CB_SETITEMHEIGHT, 0, properties.itemHeight);
				}
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_CB_ITEMHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ListAlwaysShowVerticalScrollBar(VARIANT_BOOL* pValue)
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

STDMETHODIMP ComboBox::put_ListAlwaysShowVerticalScrollBar(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_LISTALWAYSSHOWVERTICALSCROLLBAR);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.listAlwaysShowVerticalScrollBar != b) {
		properties.listAlwaysShowVerticalScrollBar = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_LISTALWAYSSHOWVERTICALSCROLLBAR);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ListBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.listBackColor;
	return S_OK;
}

STDMETHODIMP ComboBox::put_ListBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CB_LISTBACKCOLOR);
	if(properties.listBackColor != newValue) {
		properties.listBackColor = newValue;
		SetDirty(TRUE);

		if(containedListBox.IsWindow()) {
			containedListBox.Invalidate();
		}
		FireViewChange();
		FireOnChanged(DISPID_CB_LISTBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ListDragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.listDragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP ComboBox::put_ListDragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CB_LISTDRAGSCROLLTIMEBASE);

	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.listDragScrollTimeBase != newValue) {
		properties.listDragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CB_LISTDRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ListForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.listForeColor;
	return S_OK;
}

STDMETHODIMP ComboBox::put_ListForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CB_LISTFORECOLOR);
	if(properties.listForeColor != newValue) {
		properties.listForeColor = newValue;
		SetDirty(TRUE);

		if(containedListBox.IsWindow()) {
			containedListBox.Invalidate();
		}
		FireViewChange();
		FireOnChanged(DISPID_CB_LISTFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ListHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedListBox.IsWindow()) {
		WTL::CRect rc;
		containedListBox.GetWindowRect(&rc);
		properties.listHeight = rc.Height();
	}

	*pValue = properties.listHeight;
	return S_OK;
}

STDMETHODIMP ComboBox::put_ListHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_CB_LISTHEIGHT);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.listHeight != newValue) {
		properties.listHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			WTL::CRect rc;
			GetWindowRect(&rc);
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
			GetParent().ScreenToClient(&rc);
			MoveWindow(&rc, FALSE);
		}
		FireOnChanged(DISPID_CB_LISTHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ListInsertMarkColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedListBox.IsWindow()) {
		if(listInsertMark.color != OLECOLOR2COLORREF(properties.listInsertMarkColor)) {
			properties.listInsertMarkColor = listInsertMark.color;
		}
	}

	*pValue = properties.listInsertMarkColor;
	return S_OK;
}

STDMETHODIMP ComboBox::put_ListInsertMarkColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CB_LISTINSERTMARKCOLOR);
	if(properties.listInsertMarkColor != newValue) {
		properties.listInsertMarkColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			listInsertMark.color = OLECOLOR2COLORREF(properties.listInsertMarkColor);
			if(!listInsertMark.hidden) {
				RECT itemBoundingRectangle = {0};
				WTL::CRect insertMarkRect;
				if(listInsertMark.itemIndex != -1) {
					containedListBox.SendMessage(LB_GETITEMRECT, listInsertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
					if(listInsertMark.afterItem) {
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
					containedListBox.InvalidateRect(&insertMarkRect);
				}
			}
		}
		FireOnChanged(DISPID_CB_LISTINSERTMARKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ListScrollableWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.listScrollableWidth = static_cast<LONG>(SendMessage(CB_GETHORIZONTALEXTENT, 0, 0));
	}

	*pValue = properties.listScrollableWidth;
	return S_OK;
}

STDMETHODIMP ComboBox::put_ListScrollableWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_CB_LISTSCROLLABLEWIDTH);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.listScrollableWidth != newValue) {
		properties.listScrollableWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(CB_SETHORIZONTALEXTENT, properties.listScrollableWidth, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_CB_LISTSCROLLABLEWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ListWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.listWidth = static_cast<LONG>(SendMessage(CB_GETDROPPEDWIDTH, 0, 0));
	}

	*pValue = properties.listWidth;
	return S_OK;
}

STDMETHODIMP ComboBox::put_ListWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_CB_LISTWIDTH);
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
		FireOnChanged(DISPID_CB_LISTWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_Locale(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.locale = SendMessage(CB_GETLOCALE, 0, 0);
	}

	*pValue = properties.locale;
	return S_OK;
}

STDMETHODIMP ComboBox::put_Locale(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CB_LOCALE);
	if(properties.locale != newValue) {
		properties.locale = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(CB_SETLOCALE, properties.locale, 0);
		}
		FireOnChanged(DISPID_CB_LOCALE);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_MaxTextLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.maxTextLength = static_cast<LONG>(containedEdit.SendMessage(EM_GETLIMITTEXT, 0, 0));
	}

	*pValue = properties.maxTextLength;
	return S_OK;
}

STDMETHODIMP ComboBox::put_MaxTextLength(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CB_MAXTEXTLENGTH);
	if(properties.maxTextLength != newValue) {
		properties.maxTextLength = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(CB_LIMITTEXT, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength), 0);
		}
		FireOnChanged(DISPID_CB_MAXTEXTLENGTH);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_MinVisibleItems(LONG* pValue)
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

STDMETHODIMP ComboBox::put_MinVisibleItems(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CB_MINVISIBLEITEMS);
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
		FireOnChanged(DISPID_CB_MINVISIBLEITEMS);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_MouseIcon(IPictureDisp** ppMouseIcon)
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

STDMETHODIMP ComboBox::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_CB_MOUSEICON);
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
	FireOnChanged(DISPID_CB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ComboBox::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_CB_MOUSEICON);
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
	FireOnChanged(DISPID_CB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ComboBox::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP ComboBox::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CB_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_OwnerDrawItems(OwnerDrawItemsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OwnerDrawItemsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
			case CBS_OWNERDRAWFIXED:
				properties.ownerDrawItems = odiOwnerDrawFixedHeight;
				break;
			case CBS_OWNERDRAWVARIABLE:
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

STDMETHODIMP ComboBox::put_OwnerDrawItems(OwnerDrawItemsConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_OWNERDRAWITEMS);
	if(properties.ownerDrawItems != newValue) {
		properties.ownerDrawItems = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_OWNERDRAWITEMS);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP ComboBox::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CB_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ComboBox::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP ComboBox::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_REGISTERFOROLEDRAGDROP);
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
		FireOnChanged(DISPID_CB_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_RightToLeft(RightToLeftConstants* pValue)
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

STDMETHODIMP ComboBox::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_RIGHTTOLEFT);
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
		FireOnChanged(DISPID_CB_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_SelectedItem(IComboBoxItem** ppSelectedItem)
{
	ATLASSERT_POINTER(ppSelectedItem, IComboBoxItem*);
	if(!ppSelectedItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		int itemIndex = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		if(itemIndex != CB_ERR) {
			ClassFactory::InitComboItem(itemIndex, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppSelectedItem));
			return S_OK;
		} else {
			*ppSelectedItem = NULL;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ComboBox::putref_SelectedItem(IComboBoxItem* pNewSelectedItem)
{
	PUTPROPPROLOG(DISPID_CB_SELECTEDITEM);
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
	FireOnChanged(DISPID_CB_SELECTEDITEM);
	return S_OK;
}

STDMETHODIMP ComboBox::get_SelectionFieldHeight(OLE_YSIZE_PIXELS* pValue)
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

STDMETHODIMP ComboBox::put_SelectionFieldHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_CB_SELECTIONFIELDHEIGHT);
	if(properties.selectionFieldHeight != newValue) {
		properties.selectionFieldHeight = newValue;
		properties.setSelectionFieldHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.selectionFieldHeight == -1) {
				ApplyFont();
				/*Called in OnSetFont: if(properties.setItemHeight != -1) {
					if(!(GetStyle() & CBS_OWNERDRAWVARIABLE)) {
						SendMessage(CB_SETITEMHEIGHT, 0, properties.setItemHeight);
					}
				}*/
			} else {
				SendMessage(CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), properties.selectionFieldHeight);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_CB_SELECTIONFIELDHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP ComboBox::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_SHOWDRAGIMAGE);
	if(!dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage();
	} else {
		dragDropStatus.ShowDragImage();
	}

	FireOnChanged(DISPID_CB_SHOWDRAGIMAGE);
	return S_OK;
}

STDMETHODIMP ComboBox::get_Sorted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.sorted = ((GetStyle() & CBS_SORT) == CBS_SORT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.sorted);
	return S_OK;
}

STDMETHODIMP ComboBox::put_Sorted(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_SORTED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.sorted != b) {
		properties.sorted = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_SORTED);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_Style(StyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, StyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) {
			case CBS_DROPDOWN:
				properties.style = sComboDropDownList;
				break;
			case CBS_DROPDOWNLIST:
				properties.style = sDropDownList;
				break;
			case CBS_SIMPLE:
				properties.style = sComboField;
				break;
		}
	}

	*pValue = properties.style;
	return S_OK;
}

STDMETHODIMP ComboBox::put_Style(StyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_CB_STYLE);
	if(properties.style != newValue) {
		properties.style = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CB_STYLE);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP ComboBox::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CB_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ComboBox::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		BOOL useGetWindowText = TRUE;
		if(flags.windowTextIsOutdated) {
			int selectedItem = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
			CComPtr<IComboBoxItem> pSelectedItem = ClassFactory::InitComboItem(selectedItem, this);
			if(pSelectedItem) {
				if(SUCCEEDED(pSelectedItem->get_Text(&properties.text))) {
					useGetWindowText = FALSE;
				}
			}
		}

		if(useGetWindowText) {
			GetWindowText(&properties.text);
		}
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP ComboBox::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_CB_TEXT);
	properties.text.AssignBSTR(newValue);
	if(IsWindow()) {
		SetWindowText(COLE2CT(properties.text));
		flags.windowTextIsOutdated = FALSE;
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_CB_TEXT);
	SendOnDataChange();
	return S_OK;
}

STDMETHODIMP ComboBox::get_TextLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = GetWindowTextLength();
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP ComboBox::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CB_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_CB_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::get_Version(BSTR* pValue)
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

STDMETHODIMP ComboBox::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ComboBox::CloseDropDownWindow(void)
{
	if(IsWindow()) {
		SendMessage(CB_SHOWDROPDOWN, FALSE, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBox::FinishOLEDragDrop(void)
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

STDMETHODIMP ComboBox::FindItemByItemData(LONG itemData, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IComboBoxItem** ppFoundItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppFoundItem, IComboBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IComboBoxItem> pItem = startAfterItem.pdispVal;
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
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}

	if(expectsString) {
		int numberOfItems = static_cast<int>(SendMessage(CB_GETCOUNT, 0, 0));
		for(int i = itemIndex + 1; i < numberOfItems; i++) {
			if(static_cast<LONG>(SendMessage(CB_GETITEMDATA, i, 0)) == itemData) {
				ClassFactory::InitComboItem(i, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
				return S_OK;
			}
		}
		for(int i = 0; i <= itemIndex; i++) {
			if(static_cast<LONG>(SendMessage(CB_GETITEMDATA, i, 0)) == itemData) {
				ClassFactory::InitComboItem(i, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
				return S_OK;
			}
		}
		*ppFoundItem = NULL;
		return S_OK;
	} else {
		itemIndex = static_cast<LONG>(SendMessage(CB_FINDSTRING, itemIndex, itemData));
		ClassFactory::InitComboItem(itemIndex, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
		return S_OK;
	}
}

STDMETHODIMP ComboBox::FindItemByText(BSTR searchString, VARIANT_BOOL exactMatch/* = VARIANT_TRUE*/, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IComboBoxItem** ppFoundItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppFoundItem, IComboBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IComboBoxItem> pItem = startAfterItem.pdispVal;
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
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	if(!expectsString) {
		return E_FAIL;
	}

	COLE2T converter(searchString);
	LPTSTR pBuffer = converter;
	itemIndex = static_cast<LONG>(SendMessage((exactMatch == VARIANT_FALSE ? CB_FINDSTRING : CB_FINDSTRINGEXACT), itemIndex, reinterpret_cast<LPARAM>(pBuffer)));
	ClassFactory::InitComboItem(itemIndex, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
	return S_OK;
}

STDMETHODIMP ComboBox::GetClosestListInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, IComboBoxItem** ppComboItem)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppComboItem, IComboBoxItem*);
	if(!ppComboItem) {
		return E_POINTER;
	}

	int proposedItemIndex = -1;
	InsertMarkPositionConstants proposedRelativePosition = impNowhere;

	POINT pt = {x, y};
	int numberOfItems = containedListBox.SendMessage(LB_GETCOUNT, 0, 0);
	int firstVisibleItem = static_cast<int>(containedListBox.SendMessage(LB_GETTOPINDEX, 0, 0));
	RECT clientRectangle = {0};
	containedListBox.GetClientRect(&clientRectangle);
	BOOL abort = FALSE;
	for(int itemIndex = max(firstVisibleItem - 1, 0); itemIndex < numberOfItems; ++itemIndex) {
		WTL::CRect itemBoundingRectangle;
		containedListBox.SendMessage(LB_GETITEMRECT, itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
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
		*ppComboItem = NULL;
		*pRelativePosition = impNowhere;
	} else {
		*pRelativePosition = proposedRelativePosition;
		ClassFactory::InitComboItem(proposedItemIndex, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppComboItem), FALSE);
	}
	return S_OK;
}

STDMETHODIMP ComboBox::GetDropDownButtonRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(IsWindow()) {
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(*this, &controlInfo);
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

STDMETHODIMP ComboBox::GetDroppedStateRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
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

STDMETHODIMP ComboBox::GetListInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, IComboBoxItem** ppComboItem)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppComboItem, IComboBoxItem*);
	if(!ppComboItem) {
		return E_POINTER;
	}

	if(containedListBox.IsWindow()) {
		if(listInsertMark.hidden) {
			*ppComboItem = NULL;
			*pRelativePosition = impNowhere;
		} else {
			if(listInsertMark.afterItem) {
				*pRelativePosition = impAfter;
			} else {
				*pRelativePosition = impBefore;
			}
			ClassFactory::InitComboItem(listInsertMark.itemIndex, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppComboItem), FALSE);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBox::GetListInsertMarkRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(containedListBox.IsWindow()) {
		RECT itemBoundingRectangle = {0};
		RECT insertMarkRect = {0};
		if(listInsertMark.itemIndex != -1) {
			containedListBox.SendMessage(LB_GETITEMRECT, listInsertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
			if(listInsertMark.afterItem) {
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
	}
	return hr;
}

STDMETHODIMP ComboBox::GetSelection(LONG* pSelectionStart/* = NULL*/, LONG* pSelectionEnd/* = NULL*/)
{
	if(IsWindow()) {
		SendMessage(CB_GETEDITSEL, reinterpret_cast<WPARAM>(pSelectionStart), reinterpret_cast<LPARAM>(pSelectionEnd));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBox::GetSelectionFieldRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(IsWindow()) {
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(*this, &controlInfo);
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

STDMETHODIMP ComboBox::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails)
{
	ATLASSERT_POINTER(pHitTestDetails, HitTestConstants);
	if(!pHitTestDetails) {
		return E_POINTER;
	}

	if(IsWindow()) {
		POINT pt = {x, y};
		WTL::CRect rc;
		GetWindowRect(&rc);
		ScreenToClient(&rc);
		if(rc.PtInRect(pt)) {
			*pHitTestDetails = htComboBoxPortion;
			if(containedEdit.IsWindow()) {
				containedEdit.GetWindowRect(&rc);
				ScreenToClient(&rc);
				if(rc.PtInRect(pt)) {
					*pHitTestDetails = htTextBoxPortion;
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

STDMETHODIMP ComboBox::ListHitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IComboBoxItem** ppHitItem)
{
	ATLASSERT_POINTER(pHitTestDetails, HitTestConstants);
	ATLASSERT_POINTER(ppHitItem, IComboBoxItem*);
	if(!pHitTestDetails || !ppHitItem) {
		return E_POINTER;
	}

	if(containedListBox.IsWindow()) {
		ClassFactory::InitComboItem(ListBoxHitTest(x, y, pHitTestDetails), this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppHitItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBox::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP ComboBox::OpenDropDownWindow(void)
{
	if(IsWindow()) {
		SendMessage(CB_SHOWDROPDOWN, TRUE, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ComboBox::PrepareForItemInsertions(LONG numberOfItems, LONG averageStringWidth, LONG* pAllocatedItems)
{
	if(IsWindow()) {
		*pAllocatedItems = SendMessage(CB_INITSTORAGE, numberOfItems, numberOfItems * averageStringWidth * sizeof(TCHAR));
		if(*pAllocatedItems == CB_ERRSPACE) {
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

STDMETHODIMP ComboBox::Refresh(void)
{
	if(IsWindow()) {
		dragDropStatus.HideDragImage();
		Invalidate();
		UpdateWindow();
		dragDropStatus.ShowDragImage();
	}
	return S_OK;
}

STDMETHODIMP ComboBox::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP ComboBox::SelectItemByItemData(LONG itemData, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IComboBoxItem** ppFoundItem/* = NULL*/)
{
	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}

	if(expectsString) {
		// cannot use CB_SELECTSTRING, so use emulation
		HRESULT hr = FindItemByItemData(itemData, startAfterItem, ppFoundItem);
		if(SUCCEEDED(hr)) {
			hr = putref_SelectedItem(*ppFoundItem);
		}
		return hr;
	}

	ATLASSERT_POINTER(ppFoundItem, IComboBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IComboBoxItem> pItem = startAfterItem.pdispVal;
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

	itemIndex = static_cast<LONG>(SendMessage(CB_SELECTSTRING, itemIndex, itemData));
	ClassFactory::InitComboItem(itemIndex, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
	return S_OK;
}

STDMETHODIMP ComboBox::SelectItemByText(BSTR searchString, VARIANT_BOOL exactMatch/* = VARIANT_TRUE*/, VARIANT startAfterItem/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IComboBoxItem** ppFoundItem/* = NULL*/)
{
	if(exactMatch != VARIANT_FALSE) {
		// cannot use CB_SELECTSTRING, so use emulation
		HRESULT hr = FindItemByText(searchString, exactMatch, startAfterItem, ppFoundItem);
		if(SUCCEEDED(hr)) {
			hr = putref_SelectedItem(*ppFoundItem);
		}
		return hr;
	}

	ATLASSERT_POINTER(ppFoundItem, IComboBoxItem*);
	if(!ppFoundItem) {
		return E_POINTER;
	}

	LONG itemIndex = -1;
	if(startAfterItem.vt != VT_ERROR) {
		itemIndex = -1;
		if(startAfterItem.vt == VT_DISPATCH) {
			CComQIPtr<IComboBoxItem> pItem = startAfterItem.pdispVal;
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
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}
	if(!expectsString) {
		return E_FAIL;
	}

	COLE2T converter(searchString);
	LPTSTR pBuffer = converter;
	itemIndex = static_cast<LONG>(SendMessage(CB_SELECTSTRING, itemIndex, reinterpret_cast<LPARAM>(pBuffer)));
	ClassFactory::InitComboItem(itemIndex, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
	return S_OK;
}

STDMETHODIMP ComboBox::SetListInsertMarkPosition(InsertMarkPositionConstants relativePosition, IComboBoxItem* pComboItem)
{
	HRESULT hr = E_FAIL;
	if(containedListBox.IsWindow()) {
		int itemIndex = -1;
		if(pComboItem) {
			LONG l = -1;
			pComboItem->get_Index(&l);
			itemIndex = l;
		}

		int oldItemIndex = listInsertMark.itemIndex;
		BOOL oldAfterItem = listInsertMark.afterItem;
		switch(relativePosition) {
			case impNowhere:
				listInsertMark.itemIndex = -1;
				listInsertMark.afterItem = FALSE;
				listInsertMark.hidden = TRUE;
				break;
			case impBefore:
				listInsertMark.itemIndex = itemIndex;
				listInsertMark.afterItem = FALSE;
				listInsertMark.hidden = FALSE;
				break;
			case impAfter:
				listInsertMark.itemIndex = itemIndex;
				listInsertMark.afterItem = TRUE;
				listInsertMark.hidden = FALSE;
				break;
		}

		RECT itemBoundingRectangle = {0};
		WTL::CRect oldInsertMarkRect;
		if(oldItemIndex != -1) {
			containedListBox.SendMessage(LB_GETITEMRECT, oldItemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
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

		WTL::CRect newInsertMarkRect;
		if(listInsertMark.itemIndex != -1) {
			containedListBox.SendMessage(LB_GETITEMRECT, listInsertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
			if(listInsertMark.afterItem) {
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
				containedListBox.InvalidateRect(&oldInsertMarkRect);
			}
			if(newInsertMarkRect.Width() > 0) {
				containedListBox.InvalidateRect(&newInsertMarkRect);
			}
		}

		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP ComboBox::SetSelection(LONG selectionStart, LONG selectionEnd)
{
	if(IsWindow()) {
		// TODO: Maybe use EM_SETSEL? CB_SETEDITSEL is limited to 16 bit per value.
		SendMessage(CB_SETEDITSEL, 0, MAKELPARAM(selectionStart, selectionEnd));
		return S_OK;
	}
	return E_FAIL;
}


LRESULT ComboBox::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT ComboBox::OnContextMenu(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
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
		case 0:
			Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion, &showDefaultMenu);
			break;
		case 2:
			Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion, &showDefaultMenu);
			break;
	}
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}

	return 0;
}

LRESULT ComboBox::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(*this);
	}
	return lr;
}

LRESULT ComboBox::OnCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
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

LRESULT ComboBox::OnCtlColorListBox(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.listBackColor));
	SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.listForeColor));
	if(!(properties.listBackColor & 0x80000000)) {
		SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.listBackColor));
		return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
	} else {
		return reinterpret_cast<LRESULT>(GetSysColorBrush(properties.listBackColor & 0x0FFFFFFF));
	}
}

LRESULT ComboBox::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(*this);

	wasHandled = FALSE;
	return 0;
}

LRESULT ComboBox::OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if(ime != imeNoControl) {
		HWND h = GetFocus();
		if(h == *this || h == containedEdit || h == containedListBox) {
			// we've the focus, so configure the IME
			SetCurrentIMEContextMode(h, ime);
		}
	}
	return lr;
}

LRESULT ComboBox::OnKeyDown(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		/* The native combo box forwards RETURN and ESCAPE that are meant for the edit control to the container
		   control (the combo box itself). But we want the event to be raised only once. */
		if(!(index != 2 && (wParam == VK_RETURN || wParam == VK_ESCAPE) && containedEdit.IsWindow())) {
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
	}
	switch(index) {
		case 2:
			return containedEdit.DefWindowProc(message, wParam, lParam);
		default:
			return DefWindowProc(message, wParam, lParam);
	}
}

LRESULT ComboBox::OnKeyUp(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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
		case 2:
			return containedEdit.DefWindowProc(message, wParam, lParam);
		default:
			return DefWindowProc(message, wParam, lParam);
	}
}

LRESULT ComboBox::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<ComboBox>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	return lr;
}

LRESULT ComboBox::OnLButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		switch(index) {
			case 0:
				Raise_DblClick(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
				break;
			case 2:
				containedEdit.MapWindowPoints(*this, &mousePosition, 1);
				Raise_DblClick(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
				break;
		}
	}

	return 0;
}

LRESULT ComboBox::OnLButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnLButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnMButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		switch(index) {
			case 0:
				Raise_MDblClick(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
				break;
			case 2:
				containedEdit.MapWindowPoints(*this, &mousePosition, 1);
				Raise_MDblClick(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
				break;
		}
	}

	return 0;
}

LRESULT ComboBox::OnMButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnMButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT ComboBox::OnMouseHover(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseHover(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseHover(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnMouseLeave(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	switch(index) {
		case 0:
			Raise_MouseLeave(button, shift, mouseStatus_ComboBox.lastPosition.x, mouseStatus_ComboBox.lastPosition.y, htComboBoxPortion);
			break;
		case 2:
			Raise_MouseLeave(button, shift, mouseStatus_Edit.lastPosition.x, mouseStatus_Edit.lastPosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnMouseMove(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		switch(index) {
			case 0:
				Raise_MouseMove(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
				break;
			case 2:
				containedEdit.MapWindowPoints(*this, &mousePosition, 1);
				Raise_MouseMove(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
				break;
		}
	} else {
		switch(index) {
			case 0:
				if(!mouseStatus_ComboBox.enteredControl) {
					mouseStatus_ComboBox.EnterControl();
				}
				break;
			case 2:
				if(!mouseStatus_Edit.enteredControl) {
					mouseStatus_Edit.EnterControl();
				}
				break;
		}
	}

	return 0;
}

LRESULT ComboBox::OnMouseWheel(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
			case 0:
				Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam), htComboBoxPortion);
				break;
			case 2:
				Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam), htTextBoxPortion);
				break;
		}
	} else {
		switch(index) {
			case 0:
				if(!mouseStatus_ComboBox.enteredControl) {
					mouseStatus_ComboBox.EnterControl();
				}
				break;
			case 2:
				if(!mouseStatus_Edit.enteredControl) {
					mouseStatus_Edit.EnterControl();
				}
				break;
		}
	}

	return 0;
}

LRESULT ComboBox::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ComboBox::OnRButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		switch(index) {
			case 0:
				Raise_RDblClick(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
				break;
			case 2:
				containedEdit.MapWindowPoints(*this, &mousePosition, 1);
				Raise_RDblClick(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
				break;
		}
	}

	return 0;
}

LRESULT ComboBox::OnRButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnRButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnSetCursor(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	WTL::CRect clientArea;
	switch(index) {
		case 0:
			GetClientRect(&clientArea);
			ClientToScreen(&clientArea);
			break;
		case 2:
			containedEdit.GetClientRect(&clientArea);
			containedEdit.ClientToScreen(&clientArea);
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
			case 2:
				setCursor = (WindowFromPoint(mousePosition) == containedEdit);
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

LRESULT ComboBox::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<ComboBox>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<ComboBox>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT ComboBox::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_CB_FONT) == S_FALSE) {
		return 0;
	}

	if(!IsInDesignMode()) {
		if(properties.setSelectionFieldHeight != -1) {
			properties.selectionFieldHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(CB_GETITEMHEIGHT, static_cast<WPARAM>(-1), 0));
		}
		if(properties.setItemHeight != -1) {
			if(!(GetStyle() & CBS_OWNERDRAWVARIABLE)) {
				properties.itemHeight = static_cast<OLE_YSIZE_PIXELS>(SendMessage(CB_GETITEMHEIGHT, 0, 0));
			}
		}
	}
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(!IsInDesignMode()) {
		if(properties.setItemHeight != -1) {
			if(!(GetStyle() & CBS_OWNERDRAWVARIABLE)) {
				SendMessage(CB_SETITEMHEIGHT, 0, properties.itemHeight);
			}
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
		FireOnChanged(DISPID_CB_FONT);
	}

	return lr;
}

LRESULT ComboBox::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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

LRESULT ComboBox::OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_CB_TEXT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	SetDirty(TRUE);
	FireOnChanged(DISPID_CB_TEXT);
	SendOnDataChange();
	return lr;
}

LRESULT ComboBox::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT ComboBox::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT ComboBox::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT ComboBox::OnWindowPosChanged(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	WTL::CRect windowRectangle = m_rcPos;
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

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		if(!(pDetails->flags & SWP_NOSIZE) && !RunTimeHelper::IsCommCtrl6() && containedListBox.IsWindow()) {
			WTL::CRect dropDownRectangle;
			containedListBox.GetWindowRect(&dropDownRectangle);
			flags.lastListBoxHeight = dropDownRectangle.Height();
		}
		LRESULT lr = DefWindowProc(message, wParam, lParam);

		if((GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_SIMPLE && containedListBox.IsWindow()) {
			RECT dropDownRectangle = {0};
			containedListBox.GetWindowRect(&dropDownRectangle);
			GetWindowRect(&windowRectangle);
			windowRectangle.bottom = dropDownRectangle.top;
			GetParent().ScreenToClient(&windowRectangle);
			MoveWindow(&windowRectangle);
		}
		return lr;
	} else {
		wasHandled = FALSE;
		return 0;
	}
}

LRESULT ComboBox::OnXButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		switch(index) {
			case 0:
				Raise_XDblClick(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
				break;
			case 2:
				containedEdit.MapWindowPoints(*this, &mousePosition, 1);
				Raise_XDblClick(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
				break;
		}
	}

	return 0;
}

LRESULT ComboBox::OnXButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnXButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htComboBoxPortion);
			break;
		case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htTextBoxPortion);
			break;
	}

	return 0;
}

LRESULT ComboBox::OnReflectedCompareItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPCOMPAREITEMSTRUCT pDetails = reinterpret_cast<LPCOMPAREITEMSTRUCT>(lParam);
	ATLASSERT(pDetails->CtlType == ODT_COMBOBOX);
	ATLASSERT(pDetails->hwndItem == *this);

	// NOTE: The indexes may be -1.
	CComPtr<IComboBoxItem> pFirstItem = ClassFactory::InitComboItem(pDetails->itemID1, this, FALSE, pDetails->itemData1);
	CComPtr<IComboBoxItem> pSecondItem = ClassFactory::InitComboItem(pDetails->itemID2, this, FALSE, pDetails->itemData2);
	CompareResultConstants result = crEqual;
	Raise_CompareItems(pFirstItem, pSecondItem, pDetails->dwLocaleId, &result);
	return result;
}

LRESULT ComboBox::OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPDRAWITEMSTRUCT pDetails = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);

	ATLASSERT(pDetails->CtlType == ODT_COMBOBOX);
	ATLASSERT((pDetails->itemState & (ODS_GRAYED | ODS_CHECKED | ODS_DEFAULT | ODS_HOTLIGHT | ODS_INACTIVE)) == 0);

	if(pDetails->hwndItem == *this) {
		int itemIndex = pDetails->itemID;
		CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(itemIndex, this);
		Raise_OwnerDrawItem(pCBItem, static_cast<OwnerDrawActionConstants>(pDetails->itemAction), static_cast<OwnerDrawItemStateConstants>(pDetails->itemState), HandleToLong(pDetails->hDC), reinterpret_cast<RECTANGLE*>(&pDetails->rcItem));
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ComboBox::OnReflectedMeasureItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPMEASUREITEMSTRUCT pDetails = reinterpret_cast<LPMEASUREITEMSTRUCT>(lParam);
	ATLASSERT(pDetails->CtlType == ODT_COMBOBOX);

	if(pDetails->CtlType == ODT_COMBOBOX) {
		int itemIndex = pDetails->itemID;
		CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(itemIndex, this, FALSE);
		Raise_MeasureItem(pCBItem, reinterpret_cast<OLE_YSIZE_PIXELS*>(&pDetails->itemHeight));
		return TRUE;
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT ComboBox::OnAddString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	//TODO: toolTipStatus.InvalidateToolTipItem();

	if(flags.silentItemInsertions > 0) {
		LONG itemID = GetNewItemID();
		int insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem != CB_ERR) {
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

	int insertedItem = CB_ERR;
	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
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
		// TODO: For CBS_SORT, this might be the wrong itemIndex.
		int itemIndex = static_cast<int>(SendMessage(CB_GETCOUNT, 0, 0));

		CComObject<VirtualComboBoxItem>* pVCBoxItemObj = NULL;
		CComPtr<IVirtualComboBoxItem> pVCBoxItemItf = NULL;
		CComObject<VirtualComboBoxItem>::CreateInstance(&pVCBoxItemObj);
		pVCBoxItemObj->AddRef();
		pVCBoxItemObj->SetOwner(this);
		pVCBoxItemObj->LoadState(itemIndex, lParam, (expectsString ? itemData : lParam));
		pVCBoxItemObj->QueryInterface(IID_IVirtualComboBoxItem, reinterpret_cast<LPVOID*>(&pVCBoxItemItf));
		pVCBoxItemObj->Release();
		Raise_InsertingItem(pVCBoxItemItf, &cancel);
		pVCBoxItemObj = NULL;
	}

	if(cancel == VARIANT_FALSE) {
		LONG itemID = GetNewItemID();
		// finally pass the message to the combo box
		insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem != CB_ERR) {
			if(expectsString) {
				SendMessage(CB_SETITEMDATA, insertedItem, itemData);
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
				CComPtr<IComboBoxItem> pCBoxItem = ClassFactory::InitComboItem(insertedItem, this);
				if(pCBoxItem) {
					Raise_InsertedItem(pCBoxItem);
				}
			}
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
					CComPtr<IComboBoxItem> pCBoxItem = ClassFactory::InitComboItem(itemUnderMouse, this);
					Raise_ItemMouseLeave(pCBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				if(itemUnderMouse >= 0) {
					CComPtr<IComboBoxItem> pCBoxItem = ClassFactory::InitComboItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pCBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return insertedItem;
}

LRESULT ComboBox::OnDeleteString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	//TODO: toolTipStatus.InvalidateToolTipItem();

	if(flags.silentItemDeletions > 0) {
		LRESULT lr = DefWindowProc(message, wParam, lParam);
		if(lr != CB_ERR) {
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
	CComPtr<IComboBoxItem> pCBoxItemItf = NULL;
	CComObject<ComboBoxItem>* pCBoxItemObj = NULL;
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		CComObject<ComboBoxItem>::CreateInstance(&pCBoxItemObj);
		pCBoxItemObj->AddRef();
		pCBoxItemObj->SetOwner(this);
		pCBoxItemObj->Attach(static_cast<int>(wParam));
		pCBoxItemObj->QueryInterface(IID_IComboBoxItem, reinterpret_cast<LPVOID*>(&pCBoxItemItf));
		pCBoxItemObj->Release();
		Raise_RemovingItem(pCBoxItemItf, &cancel);
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
				Raise_ItemMouseLeave(pCBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				itemUnderMouse = -1;
			}
		}

		if(!(properties.disabledEvents & deFreeItemData)) {
			Raise_FreeItemData(pCBoxItemItf);
		}
		BOOL textChanged = FALSE;
		if(static_cast<int>(wParam) == currentSelectedItem) {
			if(!(properties.disabledEvents & deTextChangedEvents)) {
				CComBSTR selectedItemText;
				CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(currentSelectedItem, this);
				if(pCBItem) {
					pCBItem->get_Text(&selectedItemText);
					GetWindowText(&properties.text);

					textChanged = (properties.text == selectedItemText);
				}
			}
		}

		CComPtr<IVirtualComboBoxItem> pVCBoxItemItf = NULL;
		if(!(properties.disabledEvents & deItemDeletionEvents)) {
			CComObject<VirtualComboBoxItem>* pVCBoxItemObj = NULL;
			CComObject<VirtualComboBoxItem>::CreateInstance(&pVCBoxItemObj);
			pVCBoxItemObj->AddRef();
			pVCBoxItemObj->SetOwner(this);
			if(pCBoxItemObj) {
				pCBoxItemObj->SaveState(pVCBoxItemObj);
			}

			pVCBoxItemObj->QueryInterface(IID_IVirtualComboBoxItem, reinterpret_cast<LPVOID*>(&pVCBoxItemItf));
			pVCBoxItemObj->Release();
		}
		lr = DefWindowProc(message, wParam, lParam);
		if(lr != CB_ERR) {
			#ifdef USE_STL
				if(static_cast<int>(wParam) >= 0 && static_cast<int>(wParam) < static_cast<int>(itemIDs.size())) {
					itemIDs.erase(itemIDs.begin() + static_cast<int>(wParam));
				}
			#else
				if(static_cast<int>(wParam) >= 0 && static_cast<int>(wParam) < static_cast<int>(itemIDs.GetCount())) {
					itemIDs.RemoveAt(static_cast<int>(wParam));
				}
			#endif
			if(!(properties.disabledEvents & deItemDeletionEvents)) {
				Raise_RemovedItem(pVCBoxItemItf);
			}
			if(static_cast<int>(wParam) == currentSelectedItem) {
				// NOTE: currentSelectedItem has already been deleted, so the event is not quite right
				Raise_SelectionChanged(currentSelectedItem, -1);
				if(textChanged && !(properties.disabledEvents & deTextChangedEvents)) {
					Raise_TextChanged();
				}
			}

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
						pCBoxItemItf = ClassFactory::InitComboItem(itemUnderMouse, this);
						Raise_ItemMouseLeave(pCBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
					itemUnderMouse = newItemUnderMouse;
					if(itemUnderMouse >= 0) {
						pCBoxItemItf = ClassFactory::InitComboItem(itemUnderMouse, this);
						Raise_ItemMouseEnter(pCBoxItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}
	return lr;
}

LRESULT ComboBox::OnInsertString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	//TODO: toolTipStatus.InvalidateToolTipItem();

	if(flags.silentItemInsertions > 0) {
		LONG itemID = GetNewItemID();
		int insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem != CB_ERR) {
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

	int insertedItem = CB_ERR;
	BOOL expectsString;
	DWORD style = GetStyle();
	if(style & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & CBS_HASSTRINGS) == CBS_HASSTRINGS);
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
		CComObject<VirtualComboBoxItem>* pVCBoxItemObj = NULL;
		CComPtr<IVirtualComboBoxItem> pVCBoxItemItf = NULL;
		CComObject<VirtualComboBoxItem>::CreateInstance(&pVCBoxItemObj);
		pVCBoxItemObj->AddRef();
		pVCBoxItemObj->SetOwner(this);
		pVCBoxItemObj->LoadState(static_cast<int>(wParam), lParam, (expectsString ? itemData : lParam));
		pVCBoxItemObj->QueryInterface(IID_IVirtualComboBoxItem, reinterpret_cast<LPVOID*>(&pVCBoxItemItf));
		pVCBoxItemObj->Release();
		Raise_InsertingItem(pVCBoxItemItf, &cancel);
		pVCBoxItemObj = NULL;
	}

	if(cancel == VARIANT_FALSE) {
		LONG itemID = GetNewItemID();
		// finally pass the message to the combo box
		insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem >= 0) {
			if(expectsString) {
				SendMessage(CB_SETITEMDATA, insertedItem, itemData);
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
				CComPtr<IComboBoxItem> pCBoxItem = ClassFactory::InitComboItem(insertedItem, this);
				if(pCBoxItem) {
					Raise_InsertedItem(pCBoxItem);
				}
			}
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
					CComPtr<IComboBoxItem> pCBoxItem = ClassFactory::InitComboItem(itemUnderMouse, this);
					Raise_ItemMouseLeave(pCBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				if(itemUnderMouse >= 0) {
					CComPtr<IComboBoxItem> pCBoxItem = ClassFactory::InitComboItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pCBoxItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return insertedItem;
}

LRESULT ComboBox::OnResetContent(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	//TODO: toolTipStatus.InvalidateToolTipItem();

	/* NOTE: If CB_DELETESTRING is called for the last item in the control, the control does *not* send
	         CB_RESETCONTENT to itself. */
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
			CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(itemUnderMouse, this);
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			containedListBox.ScreenToClient(&mousePosition);
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
			ListBoxHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			Raise_ItemMouseLeave(pCBItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
			itemUnderMouse = -1;
		}
	}
	if(!(properties.disabledEvents & deFreeItemData)) {
		Raise_FreeItemData(NULL);
	}
	BOOL textChanged = FALSE;
	if(currentSelectedItem != -1) {
		if(!(properties.disabledEvents & deTextChangedEvents)) {
			CComBSTR selectedItemText;
			CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(currentSelectedItem, this);
			if(pCBItem) {
				pCBItem->get_Text(&selectedItemText);
				GetWindowText(&properties.text);

				textChanged = (properties.text == selectedItemText);
			}
		}
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
	if(currentSelectedItem != -1) {
		// NOTE: currentSelectedItem has already been deleted, so the event is not quite right
		Raise_SelectionChanged(currentSelectedItem, -1);
		if(textChanged && !(properties.disabledEvents & deTextChangedEvents)) {
			Raise_TextChanged();
		}
	}

	return lr;
}

LRESULT ComboBox::OnSetCueBanner(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_CB_CUEBANNER) == S_FALSE) {
		return FALSE;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		properties.cueBanner = reinterpret_cast<LPWSTR>(lParam);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_CB_CUEBANNER);
	SendOnDataChange();
	return lr;
}

LRESULT ComboBox::OnSetCurSel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_CB_SELECTEDITEM) == S_FALSE) {
		return FALSE;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	Raise_SelectionChanged(currentSelectedItem, wParam);

	SetDirty(TRUE);
	FireOnChanged(DISPID_CB_SELECTEDITEM);
	SendOnDataChange();
	return lr;
}

LRESULT ComboBox::OnListBoxDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_DestroyedListBoxControlWindow(containedListBox);
	return 0;
}

LRESULT ComboBox::OnListBoxLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_ListMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ComboBox::OnListBoxLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_ListMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ComboBox::OnListBoxMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_ListMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ComboBox::OnListBoxMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_ListMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ComboBox::OnListBoxMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deListBoxMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_ListMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ComboBox::OnListBoxMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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

LRESULT ComboBox::OnListBoxPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = containedListBox.DefWindowProc(message, wParam, lParam);
	if(message == WM_PAINT) {
		HDC hDC = containedListBox.GetDC();
		ListBoxDrawInsertionMark(hDC);
		containedListBox.ReleaseDC(hDC);
	} else {
		ListBoxDrawInsertionMark(reinterpret_cast<HDC>(wParam));
	}
	return lr;
}

LRESULT ComboBox::OnListBoxRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_ListMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ComboBox::OnListBoxRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_ListMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ComboBox::OnListBoxScroll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL tmp = listInsertMark.hidden;
	if(!listInsertMark.hidden) {
		listInsertMark.hidden = TRUE;
		// remove the insertion mark - we might get drawing glitches otherwise
		WTL::CRect oldInsertMarkRect;
		// calculate the current insertion mark's rectangle
		RECT itemBoundingRectangle = {0};
		containedListBox.SendMessage(LB_GETITEMRECT, listInsertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
		if(listInsertMark.afterItem) {
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
			containedListBox.InvalidateRect(&oldInsertMarkRect);
		}
	}

	LRESULT lr = containedListBox.DefWindowProc(message, wParam, lParam);

	if(!tmp) {
		// now draw the insertion mark again
		listInsertMark.hidden = FALSE;
		HDC hDC = containedListBox.GetDC();
		ListBoxDrawInsertionMark(hDC);
		containedListBox.ReleaseDC(hDC);
	}

	return lr;
}

LRESULT ComboBox::OnEditDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_DestroyedEditControlWindow(containedEdit);
	return 0;
}


LRESULT ComboBox::OnReflectedCloseUp(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	Raise_ListCloseUp();
	return 0;
}

LRESULT ComboBox::OnReflectedDropDown(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	Raise_ListDropDown();

	if(GetAsyncKeyState(VK_LBUTTON) & 0x8000) {
		// HACK: We capture mouse input on mouse-down and this interferes with drop-down handling of ComboBox.
		if(mouseStatus_ComboBox.IsClickCandidate(1/*MouseButtonConstants.vbLeftButton*/)) {
			if(GetCapture() == *this) {
				ReleaseCapture();
			}
			mouseStatus_ComboBox.RemoveClickCandidate(1/*MouseButtonConstants.vbLeftButton*/);
		}
	}
	return 0;
}

LRESULT ComboBox::OnReflectedEditChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	int selectedItem = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
	if(selectedItem != currentSelectedItem) {
		Raise_SelectionChanged(currentSelectedItem, selectedItem);
	}
	if(!(properties.disabledEvents & deTextChangedEvents)) {
		Raise_TextChanged();
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_CB_TEXT);
	SendOnDataChange();

	wasHandled = FALSE;
	return 0;
}

LRESULT ComboBox::OnReflectedEditUpdate(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deBeforeDrawText)) {
		Raise_BeforeDrawText();
	}

	return 0;
}

LRESULT ComboBox::OnReflectedErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_OutOfMemory();
	wasHandled = FALSE;
	return 0;
}

LRESULT ComboBox::OnReflectedSelChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	int newSelectedItem = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
	if(newSelectedItem != currentSelectedItem) {
		Raise_SelectionChanged(currentSelectedItem, newSelectedItem);
		if(containedEdit.IsWindow()) {
			if(!(properties.disabledEvents & deTextChangedEvents)) {
				flags.windowTextIsOutdated = TRUE;
				Raise_TextChanged();
				flags.windowTextIsOutdated = FALSE;
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ComboBox::OnReflectedSelEndCancel(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deSelectionChangingEvents)) {
		Raise_SelectionCanceled();
	}
	return 0;
}

LRESULT ComboBox::OnReflectedSelEndOk(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deSelectionChangingEvents)) {
		Raise_SelectionChanging();
	}
	return 0;
}

LRESULT ComboBox::OnReflectedSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
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

	// let ComboBox handle this notification
	wasHandled = FALSE;
	return 0;
}

LRESULT ComboBox::OnAlign(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	switch(notifyCode) {
		case EN_ALIGN_LTR_EC:
			Raise_WritingDirectionChanged(wdLeftToRight);
			break;
		case EN_ALIGN_RTL_EC:
			Raise_WritingDirectionChanged(wdRightToLeft);
			break;
	}

	// let ComboBox handle this notification
	wasHandled = FALSE;
	return 0;
}

LRESULT ComboBox::OnMaxText(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_TruncatedText();

	// let ComboBox handle this notification
	wasHandled = FALSE;
	return 0;
}


void ComboBox::ApplyFont(void)
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


inline HRESULT ComboBox::Raise_BeforeDrawText(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeforeDrawText();
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_CompareItems(IComboBoxItem* pFirstItem, IComboBoxItem* pSecondItem, LONG locale, CompareResultConstants* pResult)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CompareItems(pFirstItem, pSecondItem, locale, pResult);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu)
{
	//if(m_nFreezeEvents == 0) {
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// propose the middle of the control's client rectangle as the menu's position
				WTL::CRect clientRectangle;
				GetClientRect(&clientRectangle);
				WTL::CPoint centerPoint = clientRectangle.CenterPoint();
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

inline HRESULT ComboBox::Raise_CreatedEditControlWindow(HWND hWndEdit)
{
	// subclass the edit control
	containedEdit.SubclassWindow(hWndEdit);

	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedEditControlWindow(HandleToLong(hWndEdit));
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_CreatedListBoxControlWindow(HWND hWndListBox)
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

inline HRESULT ComboBox::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_DestroyedControlWindow(HWND hWnd)
{
	KillTimer(timers.ID_REDRAW);
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

inline HRESULT ComboBox::Raise_DestroyedEditControlWindow(HWND hWndEdit)
{
	if(!(properties.disabledEvents & deMouseEvents)) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = containedEdit;
		trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		if(mouseStatus_Edit.enteredControl) {
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			Raise_MouseLeave(button, shift, mouseStatus_Edit.lastPosition.x, mouseStatus_Edit.lastPosition.y, htTextBoxPortion);
		}
	}

	//containedEdit.UnsubclassWindow();     // done automatically
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedEditControlWindow(HandleToLong(hWndEdit));
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_DestroyedListBoxControlWindow(HWND hWndListBox)
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

inline HRESULT ComboBox::Raise_FreeItemData(IComboBoxItem* pComboItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeItemData(pComboItem);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_InsertedItem(IComboBoxItem* pComboItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedItem(pComboItem);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_InsertingItem(IVirtualComboBoxItem* pComboItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingItem(pComboItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_ItemMouseEnter(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemMouseEnter(pComboItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_ItemMouseLeave(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemMouseLeave(pComboItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_ListCloseUp(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ListCloseUp();
	//}
}

inline HRESULT ComboBox::Raise_ListDropDown(void)
{
	if(!RunTimeHelper::IsCommCtrl6()) {
		if(properties.listHeight == -1 && flags.lastListBoxHeight >= 3) {
			WTL::CRect dropDownRectangle;
			containedListBox.GetWindowRect(&dropDownRectangle);
			dropDownRectangle.bottom = dropDownRectangle.top + flags.lastListBoxHeight;
			containedListBox.MoveWindow(&dropDownRectangle);
		} else if(properties.listHeight != -1) {
			WTL::CRect dropDownRectangle;
			containedListBox.GetWindowRect(&dropDownRectangle);
			dropDownRectangle.bottom = dropDownRectangle.top + properties.listHeight;
			containedListBox.MoveWindow(&dropDownRectangle);
		}
	}
	//if(m_nFreezeEvents == 0) {
		return Fire_ListDropDown();
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_ListMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deListBoxMouseEvents)) {
		mouseStatus_ListBox.StoreClickCandidate(button);
		//if(button != 1/*MouseButtonConstants.vbLeftButton*/) {
		//	containedListBox.SetCapture();
		//}

		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
		CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(itemIndex, this);
		return Fire_ListMouseDown(pCBItem, button, shift, x, y, hitTestDetails);
	} else {
		mouseStatus_ListBox.StoreClickCandidate(button);
		//if(button != 1/*MouseButtonConstants.vbLeftButton*/) {
		//	containedListBox.SetCapture();
		//}
		return S_OK;
	}
}

inline HRESULT ComboBox::Raise_ListMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus_ListBox.lastPosition.x = x;
	mouseStatus_ListBox.lastPosition.y = y;

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
	CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(itemIndex, this);

	// Do we move over another date than before?
	if(itemIndex != itemUnderMouse) {
		if(itemUnderMouse >= 0) {
			CComPtr<IComboBoxItem> pPrevCBItem = ClassFactory::InitComboItem(itemUnderMouse, this);
			Raise_ItemMouseLeave(pPrevCBItem, button, shift, x, y, hitTestDetails);
		}
		itemUnderMouse = itemIndex;
		if(pCBItem) {
			Raise_ItemMouseEnter(pCBItem, button, shift, x, y, hitTestDetails);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_ListMouseMove(pCBItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_ListMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deListBoxMouseEvents)) {
		HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
		int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
		CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(itemIndex, this);
		hr = Fire_ListMouseUp(pCBItem, button, shift, x, y, hitTestDetails);
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

		mouseStatus_ListBox.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT ComboBox::Raise_ListMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int itemIndex = ListBoxHitTest(x, y, &hitTestDetails);
	CComPtr<IComboBoxItem> pCBItem = ClassFactory::InitComboItem(itemIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_ListMouseWheel(pCBItem, button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_ListOLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

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
	IComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitComboItem(dropTarget, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

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
			hr = Fire_ListOLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, yToItemTop, hitTestDetails);
		}
	//}

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

inline HRESULT ComboBox::Raise_ListOLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

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
	IComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitComboItem(dropTarget, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

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
		WTL::CPoint mousePos(mousePosition.x, mousePosition.y);
		WTL::CRect noScrollZone;
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

inline HRESULT ComboBox::Raise_ListOLEDragLeave(BOOL fakedLeave)
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
	CComPtr<IComboBoxItem> pDropTarget = ClassFactory::InitComboItem(dropTarget, this);

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

inline HRESULT ComboBox::Raise_ListOLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
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
	IComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitComboItem(dropTarget, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

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
		WTL::CPoint mousePos(mousePosition.x, mousePosition.y);
		WTL::CRect noScrollZone;
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

inline HRESULT ComboBox::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_MeasureItem(IComboBoxItem* pComboItem, OLE_YSIZE_PIXELS* pItemHeight)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MeasureItem(pComboItem, pItemHeight);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htTextBoxPortion) {
		pMouseStatusToUse = &mouseStatus_Edit;
	} else if(hitTestDetails & htComboBoxPortion) {
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT((hitTestDetails & htTextBoxPortion) || (hitTestDetails & htComboBoxPortion));
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
			if(containedEdit.IsWindow()) {
				trackingOptions.hwndTrack = containedEdit;
				TrackMouseEvent(&trackingOptions);
			}

			Raise_MouseHover(button, shift, x, y, hitTestDetails);
		}
		pMouseStatusToUse->StoreClickCandidate(button);
		if(hitTestDetails & htTextBoxPortion) {
			containedEdit.SetCapture();
		} else if(hitTestDetails & htComboBoxPortion) {
			SetCapture();
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
		if(hitTestDetails & htTextBoxPortion) {
			containedEdit.SetCapture();
		} else if(hitTestDetails & htComboBoxPortion) {
			SetCapture();
		}
		return S_OK;
	}
}

inline HRESULT ComboBox::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htTextBoxPortion) {
		if(mouseStatus_ComboBox.enteredControl) {
			return S_OK;
		}
		trackingOptions.hwndTrack = containedEdit;
		pMouseStatusToUse = &mouseStatus_Edit;
	} else if(hitTestDetails & htComboBoxPortion) {
		if(mouseStatus_Edit.enteredControl) {
			return S_OK;
		}
		trackingOptions.hwndTrack = *this;
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT((hitTestDetails & htTextBoxPortion) || (hitTestDetails & htComboBoxPortion));
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

inline HRESULT ComboBox::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(hitTestDetails & htTextBoxPortion) {
		mouseStatus_Edit.HoverControl();
	} else if(hitTestDetails & htComboBoxPortion) {
		mouseStatus_ComboBox.HoverControl();
	} else {
		ATLASSERT((hitTestDetails & htTextBoxPortion) || (hitTestDetails & htComboBoxPortion));
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ComboBox::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(hitTestDetails & htTextBoxPortion) {
		mouseStatus_Edit.LeaveControl();
		if(mouseStatus_ComboBox.enteredControl) {
			return S_OK;
		}
	} else if(hitTestDetails & htComboBoxPortion) {
		mouseStatus_ComboBox.LeaveControl();
		if(mouseStatus_Edit.hoveredControl) {
			return S_OK;
		}
	} else {
		ATLASSERT((hitTestDetails & htTextBoxPortion) || (hitTestDetails & htComboBoxPortion));
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ComboBox::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htTextBoxPortion) {
		pMouseStatusToUse = &mouseStatus_Edit;
	} else if(hitTestDetails & htComboBoxPortion) {
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT((hitTestDetails & htTextBoxPortion) || (hitTestDetails & htComboBoxPortion));
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(!pMouseStatusToUse->enteredControl) {
		Raise_MouseEnter(button, shift, x, y, hitTestDetails);
	}
	mouseStatus_Edit.lastPosition.x = x;
	mouseStatus_Edit.lastPosition.y = y;
	mouseStatus_ComboBox.lastPosition.x = x;
	mouseStatus_ComboBox.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y, hitTestDetails);
	}

	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htTextBoxPortion) {
		pMouseStatusToUse = &mouseStatus_Edit;
	} else if(hitTestDetails & htComboBoxPortion) {
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT((hitTestDetails & htTextBoxPortion) || (hitTestDetails & htComboBoxPortion));
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
		if(hitTestDetails & htTextBoxPortion) {
			if(GetCapture() == containedEdit) {
				ReleaseCapture();
			}
		} else if(hitTestDetails & htComboBoxPortion) {
			if(GetCapture() == *this) {
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

inline HRESULT ComboBox::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta, HitTestConstants hitTestDetails)
{
	MouseStatus* pMouseStatusToUse = NULL;
	if(hitTestDetails & htTextBoxPortion) {
		pMouseStatusToUse = &mouseStatus_Edit;
	} else if(hitTestDetails & htComboBoxPortion) {
		pMouseStatusToUse = &mouseStatus_ComboBox;
	} else {
		ATLASSERT((hitTestDetails & htTextBoxPortion) || (hitTestDetails & htComboBoxPortion));
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

inline HRESULT ComboBox::Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGDROPDOWN);
	dragDropStatus.dropDownTimerIsRunning = FALSE;

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = -1;
	if(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails))) {
		if(hitTestDetails == htComboBoxPortion || hitTestDetails == htTextBoxPortion) {
			dropTarget = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
	}
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitComboItem(dropTarget, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

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
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
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

inline HRESULT ComboBox::Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	VARIANT_BOOL autoDropDown = VARIANT_FALSE;
	if(GetStyle() & (CBS_DROPDOWN | CBS_DROPDOWNLIST)) {
		POINT pt = {mousePosition.x, mousePosition.y};
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(*this, &controlInfo);
		autoDropDown = BOOL2VARIANTBOOL(PtInRect(&controlInfo.rcButton, pt));
	}

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = -1;
	if(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails))) {
		if(hitTestDetails == htComboBoxPortion || hitTestDetails == htTextBoxPortion) {
			dropTarget = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
	}
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitComboItem(dropTarget, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	HRESULT hr = S_OK;
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
	return hr;
}

inline HRESULT ComboBox::Raise_OLEDragLeave(BOOL fakedLeave)
{
	KillTimer(timers.ID_DRAGDROPDOWN);
	dragDropStatus.dropDownTimerIsRunning = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = -1;
	if(SUCCEEDED(HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails))) {
		if(hitTestDetails == htComboBoxPortion || hitTestDetails == htTextBoxPortion) {
			dropTarget = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
	}
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	CComPtr<IComboBoxItem> pDropTarget = ClassFactory::InitComboItem(dropTarget, this);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, hitTestDetails);
		}
	//}

	if(!fakedLeave) {
		dragDropStatus.pActiveDataObject = NULL;
		dragDropStatus.OLEDragLeaveOrDrop();
		Invalidate();
	}
	return hr;
}

inline HRESULT ComboBox::Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	VARIANT_BOOL autoDropDown = VARIANT_FALSE;
	if(GetStyle() & (CBS_DROPDOWN | CBS_DROPDOWNLIST)) {
		POINT pt = {mousePosition.x, mousePosition.y};
		COMBOBOXINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(COMBOBOXINFO);
		GetComboBoxInfo(*this, &controlInfo);
		autoDropDown = BOOL2VARIANTBOOL(PtInRect(&controlInfo.rcButton, pt));
	}

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	int dropTarget = -1;
	if(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails))) {
		if(hitTestDetails == htComboBoxPortion || hitTestDetails == htTextBoxPortion) {
			dropTarget = static_cast<int>(SendMessage(CB_GETCURSEL, 0, 0));
		}
	}
	/*TODO: if(dropTarget == -1) {
		if(hitTestDetails & htBelowLastItem) {
			dropTarget = <last item>;
		}
	}*/
	IComboBoxItem* pDropTarget = NULL;
	ClassFactory::InitComboItem(dropTarget, this, IID_IComboBoxItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	HRESULT hr = S_OK;
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
	return hr;
}

inline HRESULT ComboBox::Raise_OutOfMemory(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OutOfMemory();
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_OwnerDrawItem(IComboBoxItem* pComboItem, OwnerDrawActionConstants requiredAction, OwnerDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OwnerDrawItem(pComboItem, requiredAction, itemState, hDC, pDrawingRectangle);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_RecreatedControlWindow(HWND hWnd)
{
	HWND hWndEdit = FindWindowEx(*this, NULL, WC_EDIT, NULL);
	if(hWndEdit) {
		Raise_CreatedEditControlWindow(hWndEdit);
	}
	COMBOBOXINFO comboBoxInfo = {0};
	comboBoxInfo.cbSize = sizeof(COMBOBOXINFO);
	GetComboBoxInfo(*this, &comboBoxInfo);
	if(comboBoxInfo.hwndList) {
		Raise_CreatedListBoxControlWindow(comboBoxInfo.hwndList);
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

inline HRESULT ComboBox::Raise_RemovedItem(IVirtualComboBoxItem* pComboItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedItem(pComboItem);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_RemovingItem(IComboBoxItem* pComboItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingItem(pComboItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_SelectionCanceled(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectionCanceled();
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_SelectionChanged(int previousSelectedItem, int newSelectedItem)
{
	currentSelectedItem = newSelectedItem;
	//if(m_nFreezeEvents == 0) {
		if(previousSelectedItem != newSelectedItem) {
			CComPtr<IComboBoxItem> pCBPreviousSelectedItem = ClassFactory::InitComboItem(previousSelectedItem, this);
			CComPtr<IComboBoxItem> pCBNewSelectedItem = ClassFactory::InitComboItem(newSelectedItem, this);
			return Fire_SelectionChanged(pCBPreviousSelectedItem, pCBNewSelectedItem);
		}
	//}
	return S_OK;
}

inline HRESULT ComboBox::Raise_SelectionChanging(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectionChanging();
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_TextChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TextChanged();
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_TruncatedText(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TruncatedText();
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_WritingDirectionChanged(WritingDirectionConstants newWritingDirection)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_WritingDirectionChanged(newWritingDirection);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ComboBox::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}


void ComboBox::RecreateControlWindow(void)
{
	if(m_bInPlaceActive && !flags.dontRecreate) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

IMEModeConstants ComboBox::GetCurrentIMEContextMode(HWND hWnd)
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

void ComboBox::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
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

IMEModeConstants ComboBox::GetEffectiveIMEMode(void)
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

DWORD ComboBox::GetExStyleBits(void)
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

DWORD ComboBox::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_HSCROLL | WS_VISIBLE | WS_VSCROLL;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(properties.autoHorizontalScrolling) {
		style |= CBS_AUTOHSCROLL;
	}
	switch(properties.characterConversion) {
		case ccLowerCase:
			style |= CBS_LOWERCASE;
			break;
		case ccUpperCase:
			style |= CBS_UPPERCASE;
			break;
	}
	if(properties.doOEMConversion) {
		style |= CBS_OEMCONVERT;
	}
	if(properties.hasStrings) {
		style |= CBS_HASSTRINGS;
	}
	if(!properties.integralHeight) {
		style |= CBS_NOINTEGRALHEIGHT;
	}
	if(properties.listAlwaysShowVerticalScrollBar) {
		style |= CBS_DISABLENOSCROLL;
	}
	switch(properties.ownerDrawItems) {
		case odiOwnerDrawFixedHeight:
			style |= CBS_OWNERDRAWFIXED;
			break;
		case odiOwnerDrawVariableHeight:
			style |= CBS_OWNERDRAWVARIABLE;
			break;
	}
	if(properties.sorted) {
		style |= CBS_SORT;
	}
	switch(properties.style) {
		case sComboDropDownList:
			style |= CBS_DROPDOWN;
			break;
		case sDropDownList:
			style |= CBS_DROPDOWNLIST;
			break;
		case sComboField:
			style |= CBS_SIMPLE;
			break;
	}
	return style;
}

void ComboBox::SendConfigurationMessages(void)
{
	if(properties.acceptNumbersOnly && containedEdit.IsWindow()) {
		containedEdit.ModifyStyle(0, ES_NUMBER);
	}
	// ComboBox fiddles around with the borders and edges, so enforce our own styles here
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

	SendMessage(CB_LIMITTEXT, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength), 0);
	SetWindowText(COLE2CT(properties.text));
	if(RunTimeHelper::IsCommCtrl6()) {
		SendMessage(CB_SETCUEBANNER, 0, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
	}
	SendMessage(CB_SETLOCALE, properties.locale, 0);
	SendMessage(CB_SETHORIZONTALEXTENT, properties.listScrollableWidth, 0);
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

	WTL::CRect rc;
	GetWindowRect(&rc);
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
	GetParent().ScreenToClient(&rc);
	MoveWindow(&rc, FALSE);
	if(properties.setItemHeight != -1) {
		if(!(GetStyle() & CBS_OWNERDRAWVARIABLE)) {
			PostMessage(CB_SETITEMHEIGHT, 0, properties.setItemHeight);
		}
	}
	if(properties.setSelectionFieldHeight != -1) {
		PostMessage(CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), properties.setSelectionFieldHeight);
	}
	if(RunTimeHelper::IsCommCtrl6()) {
		PostMessage(CB_SETMINVISIBLE, properties.minVisibleItems, 0);
	}

	if((GetStyle() & (CBS_SIMPLE | CBS_DROPDOWN | CBS_DROPDOWNLIST)) == CBS_SIMPLE) {
		containedListBox.ShowWindow(SW_HIDE);
	}
}

HCURSOR ComboBox::MousePointerConst2hCursor(MousePointerConstants mousePointer)
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

int ComboBox::IDToItemIndex(LONG ID)
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
	return -2;
}

LONG ComboBox::ItemIndexToID(int itemIndex)
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

LONG ComboBox::GetNewItemID(void)
{
	return ++lastItemID;
}

int ComboBox::ListBoxHitTest(LONG x, LONG y, HitTestConstants* pFlags, BOOL autoScroll/* = FALSE*/)
{
	ATLASSERT(containedListBox.IsWindow());

	int itemIndex = -1;

	POINT pt = {x, y};
	containedListBox.ClientToScreen(&pt);
	WTL::CRect rc;
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

BOOL ComboBox::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void ComboBox::ListBoxAutoScroll(void)
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
					dragDropStatus.HideDragImage();
					containedListBox.SendMessage(WM_HSCROLL, SB_LINELEFT, 0);
					dragDropStatus.ShowDragImage();
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
					dragDropStatus.HideDragImage();
					containedListBox.SendMessage(WM_HSCROLL, SB_LINERIGHT, 0);
					dragDropStatus.ShowDragImage();
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
					dragDropStatus.HideDragImage();
					containedListBox.SendMessage(WM_VSCROLL, SB_LINEUP, 0);
					dragDropStatus.ShowDragImage();
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
					dragDropStatus.HideDragImage();
					containedListBox.SendMessage(WM_VSCROLL, SB_LINEDOWN, 0);
					dragDropStatus.ShowDragImage();
				}
			}
		}
	}
}

void ComboBox::ListBoxDrawInsertionMark(CDCHandle targetDC)
{
	if(listInsertMark.itemIndex != -1 && !listInsertMark.hidden) {
		CPen pen;
		pen.CreatePen(PS_SOLID, 1, listInsertMark.color);
		HPEN hPreviousPen = targetDC.SelectPen(pen);

		RECT itemBoundingRectangle = {0};
		containedListBox.SendMessage(LB_GETITEMRECT, listInsertMark.itemIndex, reinterpret_cast<LPARAM>(&itemBoundingRectangle));
		RECT insertMarkRect = {0};
		if(listInsertMark.afterItem) {
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


void ComboBox::BufferItemData(LPARAM itemData)
{
	#ifdef USE_STL
		properties.itemDataOfInsertedItems.push(itemData);
	#else
		properties.itemDataOfInsertedItems.AddTail(itemData);
	#endif
}

void ComboBox::EnterSilentItemInsertionSection(void)
{
	++flags.silentItemInsertions;
}

void ComboBox::LeaveSilentItemInsertionSection(void)
{
	--flags.silentItemInsertions;
}

void ComboBox::EnterSilentItemDeletionSection(void)
{
	++flags.silentItemDeletions;
}

void ComboBox::LeaveSilentItemDeletionSection(void)
{
	--flags.silentItemDeletions;
}

void ComboBox::SetItemID(int itemIndex, LONG itemID)
{
	#ifdef USE_STL
		if(itemIndex >= 0 && itemIndex < static_cast<int>(itemIDs.size())) {
			itemIDs[itemIndex] = itemID;
		}
	#else
		if(itemIndex >= 0 && itemIndex < static_cast<int>(itemIDs.GetCount())) {
			itemIDs[itemIndex] = itemID;
		}
	#endif
}

void ComboBox::DecrementNextItemID(void)
{
	--lastItemID;
}