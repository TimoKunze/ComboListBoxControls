// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	CComPtr<IComboBoxItems> pComboItems = controls.cmbFont->GetComboItems();
	CComQIPtr<IEnumVARIANT> pEnum = pComboItems;
	pEnum->Reset();     // NOTE: It's important to call Reset()!
	ULONG ul = 0;
	_variant_t v;
	v.Clear();
	while(pEnum->Next(1, &v, &ul) == S_OK) {
		CComQIPtr<IComboBoxItem> pItem = v.pdispVal;
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->ItemData));
		if(GetObjectType(h) == OBJ_FONT) {
			DeleteObject(h);
		}
		v.Clear();
	}
	CComPtr<IListBoxItems> pListItems = controls.lstFont->GetListItems();
	pEnum = pListItems;
	pEnum->Reset();     // NOTE: It's important to call Reset()!
	ul = 0;
	v.Clear();
	while(pEnum->Next(1, &v, &ul) == S_OK) {
		CComQIPtr<IListBoxItem> pItem = v.pdispVal;
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->ItemData));
		if(GetObjectType(h) == OBJ_FONT) {
			DeleteObject(h);
		}
		v.Clear();
	}

	if(controls.cmbFont) {
		IDispEventImpl<IDC_CMBFONT, CMainDlg, &__uuidof(_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventUnadvise(controls.cmbFont);
		controls.cmbFont.Release();
	}
	if(controls.lstFont) {
		IDispEventImpl<IDC_LSTFONT, CMainDlg, &__uuidof(_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventUnadvise(controls.lstFont);
		controls.lstFont.Release();
	}
	if(controls.cmbColor) {
		IDispEventImpl<IDC_CMBCOLOR, CMainDlg, &__uuidof(_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventUnadvise(controls.cmbColor);
		controls.cmbColor.Release();
	}
	if(controls.lstColor) {
		IDispEventImpl<IDC_LSTCOLOR, CMainDlg, &__uuidof(_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventUnadvise(controls.lstColor);
		controls.lstColor.Release();
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	cmbFontWnd.SubclassWindow(GetDlgItem(IDC_CMBFONT));
	cmbFontWnd.QueryControl(__uuidof(IComboBox), reinterpret_cast<LPVOID*>(&controls.cmbFont));
	if(controls.cmbFont) {
		IDispEventImpl<IDC_CMBFONT, CMainDlg, &__uuidof(_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventAdvise(controls.cmbFont);
		InsertComboItems_Font();
	}
	lstFontWnd.SubclassWindow(GetDlgItem(IDC_LSTFONT));
	lstFontWnd.QueryControl(__uuidof(IListBox), reinterpret_cast<LPVOID*>(&controls.lstFont));
	if(controls.lstFont) {
		IDispEventImpl<IDC_LSTFONT, CMainDlg, &__uuidof(_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventAdvise(controls.lstFont);
		InsertListItems_Font();
	}
	cmbColorWnd.SubclassWindow(GetDlgItem(IDC_CMBCOLOR));
	cmbColorWnd.QueryControl(__uuidof(IComboBox), reinterpret_cast<LPVOID*>(&controls.cmbColor));
	if(controls.cmbColor) {
		IDispEventImpl<IDC_CMBCOLOR, CMainDlg, &__uuidof(_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventAdvise(controls.cmbColor);
		InsertComboItems_Color();
	}
	lstColorWnd.SubclassWindow(GetDlgItem(IDC_LSTCOLOR));
	lstColorWnd.QueryControl(__uuidof(IListBox), reinterpret_cast<LPVOID*>(&controls.lstColor));
	if(controls.lstColor) {
		IDispEventImpl<IDC_LSTCOLOR, CMainDlg, &__uuidof(_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventAdvise(controls.lstColor);
		InsertListItems_Color();
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.cmbFont) {
		controls.cmbFont->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

int CALLBACK CMainDlg::EnumFontFamExProc(const LPENUMLOGFONTEX lpElfe, const NEWTEXTMETRICEX* /*lpntme*/, DWORD /*FontType*/, LPARAM lParam)
{
	static CAtlString lastFontName = TEXT("");

	if(lastFontName != CAtlString(lpElfe->elfFullName)) {
		LPENUMFONTPARAM pParam = reinterpret_cast<LPENUMFONTPARAM>(lParam);
		lpElfe->elfLogFont.lfWidth = pParam->lf.lfWidth;
		lpElfe->elfLogFont.lfHeight = pParam->lf.lfHeight;

		HFONT hFont = CreateFontIndirect(&lpElfe->elfLogFont);
		if(pParam->fillList) {
			pParam->pListItemsToAddTo->Add(_bstr_t(lpElfe->elfFullName), -1, HandleToLong(hFont));
		} else {
			pParam->pComboItemsToAddTo->Add(_bstr_t(lpElfe->elfFullName), -1, HandleToLong(hFont));
		}
	}
	lastFontName = lpElfe->elfFullName;

	return 1;
}

void CMainDlg::InsertComboItems_Font(void)
{
	// insert items
	CComPtr<IComboBoxItems> pItems = controls.cmbFont->GetComboItems();
	if(pItems) {
		hDefaultFont = reinterpret_cast<HFONT>(SendMessage(static_cast<HWND>(LongToHandle(controls.cmbFont->GethWnd())), WM_GETFONT, 0, 0));
		ENUMFONTPARAM param = {0};
		ZeroMemory(&param.lf, sizeof(LOGFONT));
		GetObject(hDefaultFont, sizeof(LOGFONT), &param.lf);
		param.lf.lfFaceName[0] = '\0';
		HDC hDC = ::GetDC(NULL);

		param.fillList = FALSE;
		param.pComboItemsToAddTo = controls.cmbFont->GetComboItems();

		//param.lf.lfCharSet = ANSI_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = BALTIC_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = CHINESEBIG5_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		param.lf.lfCharSet = DEFAULT_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = EASTEUROPE_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = GB2312_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = GREEK_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = HANGUL_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = MAC_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = OEM_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = RUSSIAN_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = SHIFTJIS_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = SYMBOL_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = TURKISH_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		::ReleaseDC(NULL, hDC);
	}
}

void CMainDlg::InsertComboItems_Color(void)
{
	// insert items
	CComPtr<IComboBoxItems> pItems = controls.cmbColor->GetComboItems();
	if(pItems) {
		pItems->Add(OLESTR("ActiveBorder"), -1, COLOR_ACTIVEBORDER);
		pItems->Add(OLESTR("ActiveCaption"), -1, COLOR_ACTIVECAPTION);
		pItems->Add(OLESTR("ActiveCaptionText"), -1, COLOR_CAPTIONTEXT);
		pItems->Add(OLESTR("ActiveCaptionGradient"), -1, COLOR_GRADIENTACTIVECAPTION);
		pItems->Add(OLESTR("AppWorkSpace"), -1, COLOR_APPWORKSPACE);
		pItems->Add(OLESTR("Control"), -1, COLOR_3DFACE);
		pItems->Add(OLESTR("ControlDark"), -1, COLOR_3DSHADOW);
		pItems->Add(OLESTR("ControlDarkDark"), -1, COLOR_3DDKSHADOW);
		pItems->Add(OLESTR("ControlLight"), -1, COLOR_3DLIGHT);
		pItems->Add(OLESTR("ControlLightLight"), -1, COLOR_3DHIGHLIGHT);
		pItems->Add(OLESTR("ControlText"), -1, COLOR_BTNTEXT);
		pItems->Add(OLESTR("Desktop"), -1, COLOR_BACKGROUND);
		pItems->Add(OLESTR("GrayText"), -1, COLOR_GRAYTEXT);
		pItems->Add(OLESTR("Highlight"), -1, COLOR_HIGHLIGHT);
		pItems->Add(OLESTR("HighlightText"), -1, COLOR_HIGHLIGHTTEXT);
		pItems->Add(OLESTR("Hot"), -1, COLOR_HOTLIGHT);
		pItems->Add(OLESTR("InactiveBorder"), -1, COLOR_INACTIVEBORDER);
		pItems->Add(OLESTR("InactiveCaption"), -1, COLOR_INACTIVECAPTION);
		pItems->Add(OLESTR("InactiveCaptionGradient"), -1, COLOR_GRADIENTINACTIVECAPTION);
		pItems->Add(OLESTR("InactiveCaptionText"), -1, COLOR_INACTIVECAPTIONTEXT);
		pItems->Add(OLESTR("Info"), -1, COLOR_INFOBK);
		pItems->Add(OLESTR("InfoText"), -1, COLOR_INFOTEXT);
		pItems->Add(OLESTR("Menu"), -1, COLOR_MENU);
		pItems->Add(OLESTR("MenuBar"), -1, COLOR_MENUBAR);
		pItems->Add(OLESTR("MenuHighlight"), -1, COLOR_MENUHILIGHT);
		pItems->Add(OLESTR("MenuText"), -1, COLOR_MENUTEXT);
		pItems->Add(OLESTR("ScrollBar"), -1, COLOR_SCROLLBAR);
		pItems->Add(OLESTR("Window"), -1, COLOR_WINDOW);
		pItems->Add(OLESTR("WindowFrame"), -1, COLOR_WINDOWFRAME);
		pItems->Add(OLESTR("WindowText"), -1, COLOR_WINDOWTEXT);
	}
}

void CMainDlg::InsertListItems_Font(void)
{
	// setup tooltip
	CWindow wnd = static_cast<HWND>(LongToHandle(controls.lstFont->GethWndToolTip()));
	wnd.ModifyStyle(WS_BORDER, TTS_BALLOON);
	CToolTipCtrl tooltip = wnd;
	tooltip.SetTitle(1, TEXT("Font name:"));

	// insert items
	CComPtr<IListBoxItems> pItems = controls.lstFont->GetListItems();
	if(pItems) {
		hDefaultFont = reinterpret_cast<HFONT>(SendMessage(static_cast<HWND>(LongToHandle(controls.lstFont->GethWnd())), WM_GETFONT, 0, 0));
		ENUMFONTPARAM param = {0};
		ZeroMemory(&param.lf, sizeof(LOGFONT));
		GetObject(hDefaultFont, sizeof(LOGFONT), &param.lf);
		param.lf.lfFaceName[0] = '\0';
		HDC hDC = ::GetDC(NULL);

		param.fillList = TRUE;
		param.pListItemsToAddTo = controls.lstFont->GetListItems();

		//param.lf.lfCharSet = ANSI_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = BALTIC_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = CHINESEBIG5_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		param.lf.lfCharSet = DEFAULT_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = EASTEUROPE_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = GB2312_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = GREEK_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = HANGUL_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = MAC_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = OEM_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = RUSSIAN_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = SHIFTJIS_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = SYMBOL_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);
		//param.lf.lfCharSet = TURKISH_CHARSET;
		//EnumFontFamiliesEx(hDC, &param.lf, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		::ReleaseDC(NULL, hDC);
	}
}

void CMainDlg::InsertListItems_Color(void)
{
	// insert items
	CComPtr<IListBoxItems> pItems = controls.lstColor->GetListItems();
	if(pItems) {
		pItems->Add(OLESTR("ActiveBorder"), -1, COLOR_ACTIVEBORDER);
		pItems->Add(OLESTR("ActiveCaption"), -1, COLOR_ACTIVECAPTION);
		pItems->Add(OLESTR("ActiveCaptionText"), -1, COLOR_CAPTIONTEXT);
		pItems->Add(OLESTR("ActiveCaptionGradient"), -1, COLOR_GRADIENTACTIVECAPTION);
		pItems->Add(OLESTR("AppWorkSpace"), -1, COLOR_APPWORKSPACE);
		pItems->Add(OLESTR("Control"), -1, COLOR_3DFACE);
		pItems->Add(OLESTR("ControlDark"), -1, COLOR_3DSHADOW);
		pItems->Add(OLESTR("ControlDarkDark"), -1, COLOR_3DDKSHADOW);
		pItems->Add(OLESTR("ControlLight"), -1, COLOR_3DLIGHT);
		pItems->Add(OLESTR("ControlLightLight"), -1, COLOR_3DHIGHLIGHT);
		pItems->Add(OLESTR("ControlText"), -1, COLOR_BTNTEXT);
		pItems->Add(OLESTR("Desktop"), -1, COLOR_BACKGROUND);
		pItems->Add(OLESTR("GrayText"), -1, COLOR_GRAYTEXT);
		pItems->Add(OLESTR("Highlight"), -1, COLOR_HIGHLIGHT);
		pItems->Add(OLESTR("HighlightText"), -1, COLOR_HIGHLIGHTTEXT);
		pItems->Add(OLESTR("Hot"), -1, COLOR_HOTLIGHT);
		pItems->Add(OLESTR("InactiveBorder"), -1, COLOR_INACTIVEBORDER);
		pItems->Add(OLESTR("InactiveCaption"), -1, COLOR_INACTIVECAPTION);
		pItems->Add(OLESTR("InactiveCaptionGradient"), -1, COLOR_GRADIENTINACTIVECAPTION);
		pItems->Add(OLESTR("InactiveCaptionText"), -1, COLOR_INACTIVECAPTIONTEXT);
		pItems->Add(OLESTR("Info"), -1, COLOR_INFOBK);
		pItems->Add(OLESTR("InfoText"), -1, COLOR_INFOTEXT);
		pItems->Add(OLESTR("Menu"), -1, COLOR_MENU);
		pItems->Add(OLESTR("MenuBar"), -1, COLOR_MENUBAR);
		pItems->Add(OLESTR("MenuHighlight"), -1, COLOR_MENUHILIGHT);
		pItems->Add(OLESTR("MenuText"), -1, COLOR_MENUTEXT);
		pItems->Add(OLESTR("ScrollBar"), -1, COLOR_SCROLLBAR);
		pItems->Add(OLESTR("Window"), -1, COLOR_WINDOW);
		pItems->Add(OLESTR("WindowFrame"), -1, COLOR_WINDOWFRAME);
		pItems->Add(OLESTR("WindowText"), -1, COLOR_WINDOWTEXT);
	}
}

void __stdcall CMainDlg::FreeItemDataCmbFont(LPDISPATCH comboItem)
{
	CComQIPtr<IComboBoxItem> pItem = comboItem;
	if(pItem) {
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->ItemData));
		if(GetObjectType(h) == OBJ_FONT) {
			DeleteObject(h);
		}
	}
}

void __stdcall CMainDlg::MeasureItemCmbFont(LPDISPATCH comboItem, OLE_YSIZE_PIXELS* itemHeight)
{
	CComQIPtr<IComboBoxItem> pItem = comboItem;
	if(pItem) {
		CDC dc;
		dc.CreateCompatibleDC();
		// measure item text
		HFONT hOldFont = NULL;
		HFONT hFont = static_cast<HFONT>(LongToHandle(pItem->ItemData));
		if(GetObjectType(hFont) == OBJ_FONT) {
			hOldFont = dc.SelectFont(hFont);
		}
		UINT flags = DT_LEFT | DT_SINGLELINE;
		if(controls.cmbFont->RightToLeft & rtlText) {
			flags |= DT_RTLREADING;
		}
		LPWSTR pText = OLE2W(pItem->Text);
		CRect rc;
		dc.DrawTextW(pText, -1, &rc, flags | DT_CALCRECT);
		*itemHeight = rc.Height();

		if(hOldFont) {
			dc.SelectFont(hOldFont);
		}
	}
}

void __stdcall CMainDlg::OwnerDrawItemCmbFont(LPDISPATCH comboItem, OwnerDrawActionConstants /*requiredAction*/, OwnerDrawItemStateConstants itemState, long hDC, RECTANGLE* drawingRectangle)
{
	CComQIPtr<IComboBoxItem> pItem = comboItem;
	if(pItem) {
		CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
		COLORREF backClr, foreClr;
		if(itemState & odisSelected) {
			if(itemState & odisFocus) {
				backClr = GetSysColor(COLOR_HIGHLIGHT);
				foreClr = GetSysColor(COLOR_HIGHLIGHTTEXT);
			} else {
				backClr = GetSysColor(COLOR_3DFACE);
				foreClr = GetSysColor(COLOR_BTNTEXT);
			}
		} else {
			OleTranslateColor(controls.cmbFont->BackColor, NULL, &backClr);
			OleTranslateColor(controls.cmbFont->ForeColor, NULL, &foreClr);
		}

		// draw item background
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		CBrush brush;
		brush.CreateSolidBrush(backClr);
		dc.FillRect(&rc, brush);

		// draw item text
		HFONT hOldFont = NULL;
		HFONT hFont = static_cast<HFONT>(LongToHandle(pItem->ItemData));
		if(GetObjectType(hFont) == OBJ_FONT) {
			hOldFont = dc.SelectFont(hFont);
		}
		COLORREF oldBkColor = dc.SetBkColor(backClr);
		COLORREF oldTextColor = dc.SetTextColor(foreClr);
		UINT flags = DT_LEFT | DT_VCENTER | DT_SINGLELINE;
		if(controls.cmbFont->RightToLeft & rtlText) {
			flags |= DT_RTLREADING;
		}
		LPWSTR pText = OLE2W(pItem->Text);
		dc.DrawTextW(pText, -1, &rc, flags);

		dc.SetTextColor(oldTextColor);
		dc.SetBkColor(oldBkColor);
		if(hOldFont) {
			dc.SelectFont(hOldFont);
		}

		// draw the focus rectangle
		if((itemState & (odisSelected | odisFocus | odisNoFocusRectangle)) == (odisSelected | odisFocus)) {
			dc.DrawFocusRect(&rc);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowCmbFont(long /*hWnd*/)
{
	InsertComboItems_Font();
}

void __stdcall CMainDlg::FreeItemDataLstFont(LPDISPATCH listItem)
{
	CComQIPtr<IListBoxItem> pItem = listItem;
	if(pItem) {
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->ItemData));
		if(GetObjectType(h) == OBJ_FONT) {
			DeleteObject(h);
		}
	}
}

void __stdcall CMainDlg::ItemGetInfoTipTextLstFont(LPDISPATCH listItem, long /*maxInfoTipLength*/, BSTR* infoTipText, VARIANT_BOOL* /*abortToolTip*/)
{
	CComQIPtr<IListBoxItem> pItem = listItem;
	if(pItem) {
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->ItemData));
		if(GetObjectType(h) == OBJ_FONT) {
			*infoTipText = _bstr_t(pItem->Text).Detach();
		}
	}
}

void __stdcall CMainDlg::MeasureItemLstFont(LPDISPATCH listItem, OLE_YSIZE_PIXELS* itemHeight)
{
	CComQIPtr<IListBoxItem> pItem = listItem;
	if(pItem) {
		CDC dc;
		dc.CreateCompatibleDC();
		// measure item text
		HFONT hOldFont = NULL;
		HFONT hFont = static_cast<HFONT>(LongToHandle(pItem->ItemData));
		if(GetObjectType(hFont) == OBJ_FONT) {
			hOldFont = dc.SelectFont(hFont);
		}
		UINT flags = DT_LEFT | DT_SINGLELINE;
		if(controls.cmbFont->RightToLeft & rtlText) {
			flags |= DT_RTLREADING;
		}
		LPWSTR pText = OLE2W(pItem->Text);
		CRect rc;
		dc.DrawTextW(pText, -1, &rc, flags | DT_CALCRECT);
		*itemHeight = rc.Height();

		if(hOldFont) {
			dc.SelectFont(hOldFont);
		}
	}
}

void __stdcall CMainDlg::OwnerDrawItemLstFont(LPDISPATCH listItem, OwnerDrawActionConstants /*requiredAction*/, OwnerDrawItemStateConstants itemState, long hDC, RECTANGLE* drawingRectangle)
{
	CComQIPtr<IListBoxItem> pItem = listItem;
	if(pItem) {
		CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
		COLORREF backClr, foreClr;
		if(itemState & odisSelected) {
			if(itemState & odisFocus) {
				backClr = GetSysColor(COLOR_HIGHLIGHT);
				foreClr = GetSysColor(COLOR_HIGHLIGHTTEXT);
			} else {
				backClr = GetSysColor(COLOR_3DFACE);
				foreClr = GetSysColor(COLOR_BTNTEXT);
			}
		} else {
			OleTranslateColor(controls.lstFont->BackColor, NULL, &backClr);
			OleTranslateColor(controls.lstFont->ForeColor, NULL, &foreClr);
		}

		// draw item background
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		CBrush brush;
		brush.CreateSolidBrush(backClr);
		dc.FillRect(&rc, brush);

		// draw item text
		HFONT hOldFont = NULL;
		HFONT hFont = static_cast<HFONT>(LongToHandle(pItem->ItemData));
		if(GetObjectType(hFont) == OBJ_FONT) {
			hOldFont = dc.SelectFont(hFont);
		}
		COLORREF oldBkColor = dc.SetBkColor(backClr);
		COLORREF oldTextColor = dc.SetTextColor(foreClr);
		UINT flags = DT_LEFT | DT_VCENTER | DT_SINGLELINE;
		if(controls.lstFont->RightToLeft & rtlText) {
			flags |= DT_RTLREADING;
		}
		LPWSTR pText = OLE2W(pItem->Text);
		dc.DrawTextW(pText, -1, &rc, flags);

		dc.SetTextColor(oldTextColor);
		dc.SetBkColor(oldBkColor);
		if(hOldFont) {
			dc.SelectFont(hOldFont);
		}

		// draw the focus rectangle
		if((itemState & (odisSelected | odisFocus | odisNoFocusRectangle)) == (odisSelected | odisFocus)) {
			dc.DrawFocusRect(&rc);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowLstFont(long /*hWnd*/)
{
	InsertListItems_Font();
}

void __stdcall CMainDlg::OwnerDrawItemCmbColor(LPDISPATCH comboItem, OwnerDrawActionConstants /*requiredAction*/, OwnerDrawItemStateConstants itemState, long hDC, RECTANGLE* drawingRectangle)
{
	CComQIPtr<IComboBoxItem> pItem = comboItem;
	if(pItem) {
		CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
		WTL::CRect rcItem(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		RECT rcText = rcItem;
		rcText.left += 30;

		COLORREF backClr, foreClr;
		if(itemState & odisSelected) {
			if(itemState & odisFocus) {
				backClr = GetSysColor(COLOR_HIGHLIGHT);
				foreClr = GetSysColor(COLOR_HIGHLIGHTTEXT);
			} else {
				backClr = GetSysColor(COLOR_3DFACE);
				foreClr = GetSysColor(COLOR_BTNTEXT);
			}
		} else {
			OleTranslateColor(controls.cmbColor->BackColor, NULL, &backClr);
			OleTranslateColor(controls.cmbColor->ForeColor, NULL, &foreClr);
		}

		// draw item background
		CBrush brush;
		brush.CreateSolidBrush(backClr);
		dc.FillRect(&rcItem, brush);

		// now calculate the dimensions of the color rectangle
		RECT rcColor;
		rcColor.left = rcItem.left + 3;
		rcColor.top = rcItem.top + 2;
		rcColor.bottom = rcItem.bottom - 2;
		rcColor.right = rcColor.left + 24;

		// draw the color rectangle
		HBRUSH hPreviousBrush = dc.SelectBrush(GetSysColorBrush(pItem->ItemData));
		CPen borderPen;
		borderPen.CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
		HPEN hPreviousPen = dc.SelectPen(borderPen);
		dc.Rectangle(&rcColor);
		dc.SelectPen(hPreviousPen);
		dc.SelectBrush(hPreviousBrush);

		// draw item text
		COLORREF oldBkColor = dc.SetBkColor(backClr);
		COLORREF oldTextColor = dc.SetTextColor(foreClr);
		UINT flags = DT_LEFT | DT_VCENTER | DT_SINGLELINE;
		if(controls.cmbColor->RightToLeft & rtlText) {
			flags |= DT_RTLREADING;
		}
		LPWSTR pText = OLE2W(pItem->Text);
		dc.DrawTextW(pText, -1, &rcText, flags);

		dc.SetTextColor(oldTextColor);
		dc.SetBkColor(oldBkColor);

		// draw the focus rectangle
		if((itemState & (odisSelected | odisFocus | odisNoFocusRectangle)) == (odisSelected | odisFocus)) {
			dc.DrawFocusRect(&rcItem);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowCmbColor(long /*hWnd*/)
{
	InsertComboItems_Color();
}

void __stdcall CMainDlg::OwnerDrawItemLstColor(LPDISPATCH listItem, OwnerDrawActionConstants /*requiredAction*/, OwnerDrawItemStateConstants itemState, long hDC, RECTANGLE* drawingRectangle)
{
	CComQIPtr<IListBoxItem> pItem = listItem;
	if(pItem) {
		CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
		WTL::CRect rcItem(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		RECT rcText = rcItem;
		rcText.left += 30;

		COLORREF backClr, foreClr;
		if(itemState & odisSelected) {
			if(itemState & odisFocus) {
				backClr = GetSysColor(COLOR_HIGHLIGHT);
				foreClr = GetSysColor(COLOR_HIGHLIGHTTEXT);
			} else {
				backClr = GetSysColor(COLOR_3DFACE);
				foreClr = GetSysColor(COLOR_BTNTEXT);
			}
		} else {
			OleTranslateColor(controls.lstColor->BackColor, NULL, &backClr);
			OleTranslateColor(controls.lstColor->ForeColor, NULL, &foreClr);
		}

		// draw item background
		CBrush brush;
		brush.CreateSolidBrush(backClr);
		dc.FillRect(&rcItem, brush);

		// now calculate the dimensions of the color rectangle
		RECT rcColor;
		rcColor.left = rcItem.left + 3;
		rcColor.top = rcItem.top + 2;
		rcColor.bottom = rcItem.bottom - 2;
		rcColor.right = rcColor.left + 24;

		// draw the color rectangle
		HBRUSH hPreviousBrush = dc.SelectBrush(GetSysColorBrush(pItem->ItemData));
		CPen borderPen;
		borderPen.CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
		HPEN hPreviousPen = dc.SelectPen(borderPen);
		dc.Rectangle(&rcColor);
		dc.SelectPen(hPreviousPen);
		dc.SelectBrush(hPreviousBrush);

		// draw item text
		COLORREF oldBkColor = dc.SetBkColor(backClr);
		COLORREF oldTextColor = dc.SetTextColor(foreClr);
		UINT flags = DT_LEFT | DT_VCENTER | DT_SINGLELINE;
		if(controls.lstColor->RightToLeft & rtlText) {
			flags |= DT_RTLREADING;
		}
		LPWSTR pText = OLE2W(pItem->Text);
		dc.DrawTextW(pText, -1, &rcText, flags);

		dc.SetTextColor(oldTextColor);
		dc.SetBkColor(oldBkColor);

		// draw the focus rectangle
		if((itemState & (odisSelected | odisFocus | odisNoFocusRectangle)) == (odisSelected | odisFocus)) {
			dc.DrawFocusRect(&rcItem);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowLstColor(long /*hWnd*/)
{
	InsertListItems_Color();
}
