// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"


BOOL CMainDlg::IsComctl32Version600OrNewer(void)
{
	BOOL ret = FALSE;

	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		DLLGETVERSIONPROC pDllGetVersion = reinterpret_cast<DLLGETVERSIONPROC>(GetProcAddress(hMod, "DllGetVersion"));
		if(pDllGetVersion) {
			DLLVERSIONINFO versionInfo = {0};
			versionInfo.cbSize = sizeof(versionInfo);
			HRESULT hr = (*pDllGetVersion)(&versionInfo);
			if(SUCCEEDED(hr)) {
				ret = (versionInfo.dwMajorVersion >= 6);
			}
		}
		FreeLibrary(hMod);
	}

	return ret;
}

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
	if(controls.lstU) {
		IDispEventImpl<IDC_LSTU, CMainDlg, &__uuidof(CBLCtlsLibU::_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventUnadvise(controls.lstU);
		controls.lstU.Release();
	}
	if(controls.cmbU) {
		IDispEventImpl<IDC_CMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventUnadvise(controls.cmbU);
		controls.cmbU.Release();
	}
	if(controls.drvcmbU) {
		IDispEventImpl<IDC_DRVCMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventUnadvise(controls.drvcmbU);
		controls.drvcmbU.Release();
	}
	if(controls.imgcmbU) {
		IDispEventImpl<IDC_IMGCMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IImageComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventUnadvise(controls.imgcmbU);
		controls.imgcmbU.Release();
	}
	if(controls.lstA) {
		IDispEventImpl<IDC_LSTA, CMainDlg, &__uuidof(CBLCtlsLibA::_IListBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>::DispEventUnadvise(controls.lstA);
		controls.lstA.Release();
	}
	if(controls.cmbA) {
		IDispEventImpl<IDC_CMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>::DispEventUnadvise(controls.cmbA);
		controls.cmbA.Release();
	}
	if(controls.drvcmbA) {
		IDispEventImpl<IDC_DRVCMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>::DispEventUnadvise(controls.drvcmbA);
		controls.drvcmbA.Release();
	}
	if(controls.imgcmbA) {
		IDispEventImpl<IDC_IMGCMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IImageComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>::DispEventUnadvise(controls.imgcmbA);
		controls.imgcmbA.Release();
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
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	HMODULE hMod = LoadLibrary(TEXT("shell32.dll"));
	if(hMod) {
		if(IsComctl32Version600OrNewer()) {
			controls.imageList.Create(16, 16, ILC_COLOR32 | ILC_MASK, 10, 0);
			controls.imageList.SetBkColor(CLR_NONE);

			int resIDs[10] = {1489, 516, 26, 185, 266, 51, 275, 126, 37, 1029};
			if(RunTimeHelper::IsVista()) {
				resIDs[0] = 60;
				resIDs[1] = 75;
				resIDs[2] = 117;
				resIDs[3] = 133;
				resIDs[4] = 180;
				resIDs[5] = 189;
				resIDs[6] = 258;
				resIDs[7] = 320;
				resIDs[9] = 384;
			}

			for(int i = 0; i < 10; i++) {
				HRSRC hResource = FindResource(hMod, MAKEINTRESOURCE(resIDs[i]), RT_ICON);
				if(hResource) {
					HGLOBAL hMem = LoadResource(hMod, hResource);
					if(hMem) {
						LPVOID pIconData = LockResource(hMem);
						if(pIconData) {
							HICON hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
							if(hIcon) {
								controls.imageList.AddIcon(hIcon);
								DestroyIcon(hIcon);
							}
						}
					}
				}
			}
		} else {
			controls.imageList.Create(16, 16, ILC_COLOR24 | ILC_MASK, 10, 0);
			controls.imageList.SetBkColor(CLR_NONE);

			int resIDs[10] = {1486, 513, 22, 181, 263, 48, 272, 122, 34, 1026};
			if(RunTimeHelper::IsVista()) {
				resIDs[0] = 57;
				resIDs[1] = 72;
				resIDs[2] = 114;
				resIDs[3] = 129;
				resIDs[4] = 177;
				resIDs[5] = 186;
				resIDs[6] = 255;
				resIDs[7] = 317;
				resIDs[9] = 381;
			}

			for(int i = 0; i < 10; i++) {
				HRSRC hResource = FindResource(hMod, MAKEINTRESOURCE(resIDs[i]), RT_ICON);
				if(hResource) {
					HGLOBAL hMem = LoadResource(hMod, hResource);
					if(hMem) {
						LPVOID pIconData = LockResource(hMem);
						if(pIconData) {
							HICON hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
							if(hIcon) {
								controls.imageList.AddIcon(hIcon);
								DestroyIcon(hIcon);
							}
						}
					}
				}
			}
		}
		FreeLibrary(hMod);
	}

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.logEdit = GetDlgItem(IDC_EDITLOG);
	controls.aboutButton = GetDlgItem(ID_APP_ABOUT);

	lstUWnd.SubclassWindow(GetDlgItem(IDC_LSTU));
	lstUWnd.QueryControl(__uuidof(CBLCtlsLibU::IListBox), reinterpret_cast<LPVOID*>(&controls.lstU));
	if(controls.lstU) {
		IDispEventImpl<IDC_LSTU, CMainDlg, &__uuidof(CBLCtlsLibU::_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventAdvise(controls.lstU);
		InsertListItemsU();
	}
	cmbUWnd.SubclassWindow(GetDlgItem(IDC_CMBU));
	cmbUWnd.QueryControl(__uuidof(CBLCtlsLibU::IComboBox), reinterpret_cast<LPVOID*>(&controls.cmbU));
	if(controls.cmbU) {
		IDispEventImpl<IDC_CMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventAdvise(controls.cmbU);
		InsertComboItemsU();
	}
	drvcmbUWnd.SubclassWindow(GetDlgItem(IDC_DRVCMBU));
	drvcmbUWnd.QueryControl(__uuidof(CBLCtlsLibU::IDriveComboBox), reinterpret_cast<LPVOID*>(&controls.drvcmbU));
	if(controls.drvcmbU) {
		IDispEventImpl<IDC_DRVCMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventAdvise(controls.drvcmbU);
	}
	imgcmbUWnd.SubclassWindow(GetDlgItem(IDC_IMGCMBU));
	imgcmbUWnd.QueryControl(__uuidof(CBLCtlsLibU::IImageComboBox), reinterpret_cast<LPVOID*>(&controls.imgcmbU));
	if(controls.imgcmbU) {
		IDispEventImpl<IDC_IMGCMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IImageComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>::DispEventAdvise(controls.imgcmbU);
		InsertImageComboItemsU();
	}

	lstAWnd.SubclassWindow(GetDlgItem(IDC_LSTA));
	lstAWnd.QueryControl(__uuidof(CBLCtlsLibA::IListBox), reinterpret_cast<LPVOID*>(&controls.lstA));
	if(controls.lstA) {
		IDispEventImpl<IDC_LSTA, CMainDlg, &__uuidof(CBLCtlsLibA::_IListBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>::DispEventAdvise(controls.lstA);
		InsertListItemsA();
	}
	cmbAWnd.SubclassWindow(GetDlgItem(IDC_CMBA));
	cmbAWnd.QueryControl(__uuidof(CBLCtlsLibA::IComboBox), reinterpret_cast<LPVOID*>(&controls.cmbA));
	if(controls.cmbA) {
		IDispEventImpl<IDC_CMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>::DispEventAdvise(controls.cmbA);
		InsertComboItemsA();
	}
	drvcmbAWnd.SubclassWindow(GetDlgItem(IDC_DRVCMBA));
	drvcmbAWnd.QueryControl(__uuidof(CBLCtlsLibA::IDriveComboBox), reinterpret_cast<LPVOID*>(&controls.drvcmbA));
	if(controls.drvcmbA) {
		IDispEventImpl<IDC_DRVCMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>::DispEventAdvise(controls.drvcmbA);
	}
	imgcmbAWnd.SubclassWindow(GetDlgItem(IDC_IMGCMBA));
	imgcmbAWnd.QueryControl(__uuidof(CBLCtlsLibA::IImageComboBox), reinterpret_cast<LPVOID*>(&controls.imgcmbA));
	if(controls.imgcmbA) {
		IDispEventImpl<IDC_IMGCMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IImageComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>::DispEventAdvise(controls.imgcmbA);
		InsertImageComboItemsA();
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.logEdit && controls.aboutButton) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			WTL::CRect rc;
			GetClientRect(&rc);

			lstUWnd.SetWindowPos(NULL, 0, 0, 193, (rc.Height() - 5) / 2, SWP_NOMOVE);
			int y = 3;
			WTL::CRect rect;
			cmbUWnd.GetWindowRect(&rect);
			cmbUWnd.SetWindowPos(NULL, 198, y, rect.Width(), rect.Height(), 0);
			y += rect.Height() + 5;
			drvcmbUWnd.GetWindowRect(&rect);
			drvcmbUWnd.SetWindowPos(NULL, 198, y, rect.Width(), rect.Height(), 0);
			y += rect.Height() + 5;
			imgcmbUWnd.GetWindowRect(&rect);
			imgcmbUWnd.SetWindowPos(NULL, 198, y, rect.Width(), rect.Height(), 0);
			controls.logEdit.SetWindowPos(NULL, 198 + rect.Width() + 5, 0, rc.Width() - (198 + rect.Width() + 5), rc.Height() - 42, 0);
			controls.aboutButton.SetWindowPos(NULL, 198 + rect.Width() + 5, rc.Height() - 37, 0, 0, SWP_NOSIZE);

			y = (rc.Height() - 5) / 2 + 5;
			lstAWnd.SetWindowPos(NULL, 0, y, 193, (rc.Height() - 5) / 2, 0);
			y += 3;
			cmbAWnd.GetWindowRect(&rect);
			cmbAWnd.SetWindowPos(NULL, 198, y, rect.Width(), rect.Height(), 0);
			y += rect.Height() + 5;
			drvcmbAWnd.GetWindowRect(&rect);
			drvcmbAWnd.SetWindowPos(NULL, 198, y, rect.Width(), rect.Height(), 0);
			y += rect.Height() + 5;
			imgcmbAWnd.GetWindowRect(&rect);
			imgcmbAWnd.SetWindowPos(NULL, 198, y, rect.Width(), rect.Height(), 0);
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	switch(focusedControl) {
		case 1:
			if(controls.lstU) {
				controls.lstU->About();
			}
			break;
		case 2:
			if(controls.cmbU) {
				controls.cmbU->About();
			}
			break;
		case 3:
			if(controls.drvcmbU) {
				controls.drvcmbU->About();
			}
			break;
		case 4:
			if(controls.imgcmbU) {
				controls.imgcmbU->About();
			}
			break;
		case 5:
			if(controls.lstA) {
				controls.lstA->About();
			}
			break;
		case 6:
			if(controls.cmbA) {
				controls.cmbA->About();
			}
			break;
		case 7:
			if(controls.drvcmbA) {
				controls.drvcmbA->About();
			}
			break;
		case 8:
			if(controls.imgcmbA) {
				controls.imgcmbA->About();
			}
			break;
	}
	return 0;
}

LRESULT CMainDlg::OnSetFocusLstU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 1;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusCmbU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 2;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusDrvCmbU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 3;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusImgCmbU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 4;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusLstA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 5;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusCmbA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 6;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusDrvCmbA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 7;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusImgCmbA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 8;
	bHandled = FALSE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertComboItemsA(void)
{
	CComPtr<CBLCtlsLibA::IComboBoxItems> pItems = controls.cmbA->GetComboItems();
	for(int i = 1; i <= 20; i++) {
		WTL::CString s;
		s.Format(TEXT("Item %i"), i);
		pItems->Add(_bstr_t(s), -1, i);
	}
}

void CMainDlg::InsertImageComboItemsA(void)
{
	controls.imgcmbA->PuthImageList(CBLCtlsLibA::ilItems, HandleToLong(controls.imageList.m_hImageList));
	int cImages = controls.imageList.GetImageCount();

	CComPtr<CBLCtlsLibA::IImageComboBoxItems> pItems = controls.imgcmbA->GetComboItems();
	for(int i = 1; i <= 20; i++) {
		WTL::CString s;
		s.Format(TEXT("Item %i"), i);
		pItems->Add(_bstr_t(s), -1, (i - 1) % cImages, i % cImages, 0, 0, i);
	}
}

void CMainDlg::InsertListItemsA(void)
{
	CComPtr<CBLCtlsLibA::IListBoxItems> pItems = controls.lstA->GetListItems();
	for(int i = 1; i <= 20; i++) {
		WTL::CString s;
		s.Format(TEXT("Item %i"), i);
		pItems->Add(_bstr_t(s), -1, i);
	}
}

void CMainDlg::InsertComboItemsU(void)
{
	CComPtr<CBLCtlsLibU::IComboBoxItems> pItems = controls.cmbU->GetComboItems();
	for(int i = 1; i <= 20; i++) {
		WTL::CString s;
		s.Format(TEXT("Item %i"), i);
		pItems->Add(_bstr_t(s), -1, i);
	}
}

void CMainDlg::InsertImageComboItemsU(void)
{
	controls.imgcmbU->PuthImageList(CBLCtlsLibU::ilItems, HandleToLong(controls.imageList.m_hImageList));
	int cImages = controls.imageList.GetImageCount();

	CComPtr<CBLCtlsLibU::IImageComboBoxItems> pItems = controls.imgcmbU->GetComboItems();
	for(int i = 1; i <= 20; i++) {
		WTL::CString s;
		s.Format(TEXT("Item %i"), i);
		pItems->Add(_bstr_t(s), -1, (i - 1) % cImages, i % cImages, 0, 0, i);
	}
}

void CMainDlg::InsertListItemsU(void)
{
	CComPtr<CBLCtlsLibU::IListBoxItems> pItems = controls.lstU->GetListItems();
	for(int i = 1; i <= 20; i++) {
		WTL::CString s;
		s.Format(TEXT("Item %i"), i);
		pItems->Add(_bstr_t(s), -1, i);
	}
}


void __stdcall CMainDlg::CaretChangedLstU(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("LstU_CaretChanged: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_CaretChanged: previousCaretItem=NULL");
	}
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::AbortedDragLstU(void)
{
	AddLogEntry(TEXT("LstU_AbortedDrag"));
}

void __stdcall CMainDlg::ClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_Click: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_Click: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareItemsLstU(LPDISPATCH firstItem, LPDISPATCH secondItem, LONG locale, CBLCtlsLibU::CompareResultConstants* result)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("LstU_CompareItems: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_CompareItems: firstItem=NULL");
	}
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pSecondItem = secondItem;
	if(pSecondItem) {
		BSTR text = pSecondItem->GetText();
		str += TEXT(", secondItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", locale=%i, result=%i"), locale, *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_ContextMenu: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_ContextMenu: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_DblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_DblClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowLstU(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("LstU_DestroyedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveLstU(LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_DragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_DragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, autoHScrollVelocity, autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropLstU(LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_Drop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_Drop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataLstU(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_FreeItemData: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_FreeItemData: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemLstU(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_InsertedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_InsertedItem: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemLstU(LPDISPATCH listItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IVirtualListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_InsertingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_InsertingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_ItemBeginDrag: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_ItemBeginDrag: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginRDragLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_ItemBeginRDrag: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_ItemBeginRDrag: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoLstU(LPDISPATCH listItem, CBLCtlsLibU::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* overlayIndex)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_ItemGetDisplayInfo: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_ItemGetDisplayInfo: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, iconIndex=%i, overlayIndex=%i"), requestedInfo, *iconIndex, *overlayIndex);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetInfoTipTextLstU(LPDISPATCH listItem, LONG maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_ItemGetInfoTipText: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_ItemGetInfoTipText: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", maxInfoTipLength=%i, infoTipText=%s, abortToolTip=%i"), maxInfoTipLength, *infoTipText, *abortToolTip);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_ItemMouseEnter: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_ItemMouseEnter: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_ItemMouseLeave: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_ItemMouseLeave: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownLstU(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("LstU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressLstU(SHORT* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("LstU_KeyPress: keyAscii=%i"), *keyAscii);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpLstU(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("LstU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MDblClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MeasureItemLstU(LPDISPATCH listItem, OLE_YSIZE_PIXELS* itemHeight)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MeasureItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MeasureItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemHeight=%i"), *itemHeight);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MouseDown: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MouseDown: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MouseEnter: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MouseEnter: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MouseHover: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MouseHover: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MouseLeave: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MouseLeave: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MouseMove: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MouseMove: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MouseUp: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MouseUp: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_MouseWheel: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_MouseWheel: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragLstU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstU_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropLstU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.lstU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterLstU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoDropDown=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetLstU(LONG hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("LstU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveLstU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstU_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetLstU(void)
{
	AddLogEntry(TEXT("LstU_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveLstU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackLstU(CBLCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("LstU_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragLstU(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("LstU_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataLstU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstU_OLEReceivedNewData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataLstU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstU_OLESetData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragLstU(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstU_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryLstU(void)
{
	AddLogEntry(TEXT("LstU_OutOfMemory"));
}

void __stdcall CMainDlg::OwnerDrawItemLstU(LPDISPATCH listItem, CBLCtlsLibU::OwnerDrawActionConstants requiredAction, CBLCtlsLibU::OwnerDrawItemStateConstants itemState, LONG hDC, CBLCtlsLibU::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_OwnerDrawItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_OwnerDrawItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requiredAction=%i, itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), requiredAction, itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ProcessCharacterInputLstU(SHORT keyAscii, SHORT shift, LONG* itemToSelect)
{
	CAtlString str;
	str.Format(TEXT("LstU_ProcessCharacterInput: keyAscii=%i, shift=%i, itemToSelect=%i"), keyAscii, shift, *itemToSelect);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ProcessKeyStrokeLstU(SHORT keyCode, SHORT shift, LONG* itemToSelect)
{
	CAtlString str;
	str.Format(TEXT("LstU_ProcessKeyStroke: keyCode=%i, shift=%i, itemToSelect=%i"), keyCode, shift, *itemToSelect);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_RClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_RClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_RDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_RDblClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowLstU(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("LstU_RecreatedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
	InsertListItemsU();
}

void __stdcall CMainDlg::RemovedItemLstU(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IVirtualListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_RemovedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_RemovedItem: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemLstU(LPDISPATCH listItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_RemovingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_RemovingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowLstU(void)
{
	AddLogEntry(TEXT("LstU_ResizedControlWindow"));
}

void __stdcall CMainDlg::SelectionChangedLstU(void)
{
	AddLogEntry(TEXT("LstU_SelectionChanged"));
}

void __stdcall CMainDlg::XClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_XClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_XClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstU_XDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstU_XDblClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectionChangedCmbU(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pPrevItem = previousSelectedItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("CmbU_SelectionChanged: previousSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_SelectionChanged: previousSelectedItem=NULL");
	}
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newSelectedItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeforeDrawTextCmbU(void)
{
	AddLogEntry(TEXT("CmbU_BeforeDrawText"));
}

void __stdcall CMainDlg::ClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_Click: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareItemsCmbU(LPDISPATCH firstItem, LPDISPATCH secondItem, LONG locale, CBLCtlsLibU::CompareResultConstants* result)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("CmbU_CompareItems: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_CompareItems: firstItem=NULL");
	}
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pSecondItem = secondItem;
	if(pSecondItem) {
		BSTR text = pSecondItem->GetText();
		str += TEXT(", secondItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", locale=%i, result=%i"), locale, *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("CmbU_ContextMenu: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowCmbU(LONG hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("CmbU_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedListBoxControlWindowCmbU(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("CmbU_CreatedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_DblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowCmbU(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("CmbU_DestroyedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowCmbU(LONG hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("CmbU_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedListBoxControlWindowCmbU(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("CmbU_DestroyedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_FreeItemData: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_FreeItemData: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_InsertedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_InsertedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IVirtualComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_InsertingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_InsertingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_ItemMouseEnter: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_ItemMouseEnter: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_ItemMouseLeave: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_ItemMouseLeave: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownCmbU(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("CmbU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressCmbU(SHORT* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("CmbU_KeyPress: keyAscii=%i"), *keyAscii);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpCmbU(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("CmbU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ListCloseUpCmbU(void)
{
	AddLogEntry(TEXT("CmbU_ListCloseUp"));
}

void __stdcall CMainDlg::ListDropDownCmbU(void)
{
	AddLogEntry(TEXT("CmbU_ListDropDown"));
}

void __stdcall CMainDlg::ListMouseDownCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_ListMouseDown: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_ListMouseDown: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseMoveCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_ListMouseMove: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_ListMouseMove: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseUpCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_ListMouseUp: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_ListMouseUp: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseWheelCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_ListMouseWheel: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_ListMouseWheel: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragDropCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbU_ListOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbU_ListOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.cmbU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::ListOLEDragEnterCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbU_ListOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbU_ListOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragLeaveCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbU_ListOLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbU_ListOLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoCloseUp=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoCloseUp);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragMouseMoveCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbU_ListOLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbU_ListOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MeasureItemCmbU(LPDISPATCH comboItem, OLE_YSIZE_PIXELS* itemHeight)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_MeasureItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_MeasureItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemHeight=%i"), *itemHeight);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MouseDown: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MouseEnter: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MouseHover: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MouseLeave: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MouseMove: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MouseUp: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.cmbU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbU_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryCmbU(void)
{
	AddLogEntry(TEXT("CmbU_OutOfMemory"));
}

void __stdcall CMainDlg::OwnerDrawItemCmbU(LPDISPATCH comboItem, CBLCtlsLibU::OwnerDrawActionConstants requiredAction, CBLCtlsLibU::OwnerDrawItemStateConstants itemState, LONG hDC, CBLCtlsLibU::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_OwnerDrawItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_OwnerDrawItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requiredAction=%i, itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), requiredAction, itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_RClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_RDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowCmbU(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("CmbU_RecreatedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
	InsertComboItemsU();
}

void __stdcall CMainDlg::RemovedItemCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IVirtualComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_RemovedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_RemovedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbU_RemovingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbU_RemovingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowCmbU(void)
{
	AddLogEntry(TEXT("CmbU_ResizedControlWindow"));
}

void __stdcall CMainDlg::SelectionCanceledCmbU(void)
{
	AddLogEntry(TEXT("CmbU_SelectionCanceled"));
}

void __stdcall CMainDlg::SelectionChangingCmbU(void)
{
	AddLogEntry(TEXT("CmbU_SelectionChanging"));
}

void __stdcall CMainDlg::TextChangedCmbU(void)
{
	AddLogEntry(TEXT("CmbU_TextChanged"));
}

void __stdcall CMainDlg::TruncatedTextCmbU(void)
{
	AddLogEntry(TEXT("CmbU_TruncatedText"));
}

void __stdcall CMainDlg::WritingDirectionChangedCmbU(CBLCtlsLibU::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("CmbU_WritingDirectionChanged: newWritingDirection=%i"), newWritingDirection);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_XClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbU_XDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectedDriveChangedDrvCmbU(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pPrevItem = previousSelectedItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("DrvCmbU_SelectedDriveChanged: previousSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_SelectedDriveChanged: previousSelectedItem=NULL");
	}
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newSelectedItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginSelectionChangeDrvCmbU(void)
{
	AddLogEntry(TEXT("DrvCmbU_BeginSelectionChange"));
}

void __stdcall CMainDlg::ClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_Click: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_ContextMenu: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedComboBoxControlWindowDrvCmbU(LONG hWndComboBox)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_CreatedComboBoxControlWindow: hWndComboBox=0x%X"), hWndComboBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedListBoxControlWindowDrvCmbU(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_CreatedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_DblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedComboBoxControlWindowDrvCmbU(LONG hWndComboBox)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_DestroyedComboBoxControlWindow: hWndComboBox=0x%X"), hWndComboBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowDrvCmbU(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_DestroyedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedListBoxControlWindowDrvCmbU(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_DestroyedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataDrvCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_FreeItemData: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_FreeItemData: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemDrvCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_InsertedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_InsertedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemDrvCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IVirtualDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_InsertingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_InsertingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragDrvCmbU(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ItemBeginDrag: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ItemBeginDrag: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), selectionFieldText, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginRDragDrvCmbU(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ItemBeginRDrag: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ItemBeginRDrag: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), selectionFieldText, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoDrvCmbU(LPDISPATCH comboItem, CBLCtlsLibU::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* selectedIconIndex, LONG* overlayIndex, LONG* indent, LONG maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		str.Format(TEXT("DrvCmbU_ItemGetDisplayInfo: comboItem=%i"), pItem->GetIndex());
	} else {
		str += TEXT("DrvCmbU_ItemGetDisplayInfo: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, iconIndex=%i, selectedIconIndex=%i, overlayIndex=%i, indent=%i, maxItemTextLength=%i, itemText=%s, dontAskAgain=%i"), requestedInfo, *iconIndex, *selectedIconIndex, *overlayIndex, *indent, maxItemTextLength, *itemText, *dontAskAgain);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ItemMouseEnter: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ItemMouseEnter: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ItemMouseLeave: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ItemMouseLeave: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownDrvCmbU(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressDrvCmbU(SHORT* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_KeyPress: keyAscii=%i"), *keyAscii);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpDrvCmbU(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ListClickDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ListClick: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ListClick: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListCloseUpDrvCmbU(void)
{
	AddLogEntry(TEXT("DrvCmbU_ListCloseUp"));
}

void __stdcall CMainDlg::ListDropDownDrvCmbU(void)
{
	AddLogEntry(TEXT("DrvCmbU_ListDropDown"));
}

void __stdcall CMainDlg::ListMouseDownDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ListMouseDown: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ListMouseDown: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseMoveDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ListMouseMove: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ListMouseMove: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseUpDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ListMouseUp: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ListMouseUp: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseWheelDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_ListMouseWheel: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_ListMouseWheel: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragDropDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_ListOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_ListOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.drvcmbU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::ListOLEDragEnterDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_ListOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_ListOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragLeaveDrvCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_ListOLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_ListOLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoCloseUp=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoCloseUp);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragMouseMoveDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_ListOLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_ListOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MouseDown: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MouseEnter: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MouseHover: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MouseLeave: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MouseMove: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MouseUp: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.drvcmbU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetDrvCmbU(LONG hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveDrvCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetDrvCmbU(void)
{
	AddLogEntry(TEXT("DrvCmbU_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackDrvCmbU(CBLCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragDrvCmbU(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataDrvCmbU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_OLEReceivedNewData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataDrvCmbU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_OLESetData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragDrvCmbU(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbU_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbU_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryDrvCmbU(void)
{
	AddLogEntry(TEXT("DrvCmbU_OutOfMemory"));
}

void __stdcall CMainDlg::RClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_RClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_RDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowDrvCmbU(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_RecreatedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovedItemDrvCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IVirtualDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_RemovedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_RemovedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemDrvCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbU_RemovingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_RemovingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowDrvCmbU(void)
{
	AddLogEntry(TEXT("DrvCmbU_ResizedControlWindow"));
}

void __stdcall CMainDlg::SelectionCanceledDrvCmbU(void)
{
	AddLogEntry(TEXT("DrvCmbU_SelectionCanceled"));
}

void __stdcall CMainDlg::SelectionChangedDrvCmbU(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pPrevItem = previousSelectedItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("DrvCmbU_SelectionChanged: previousSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_SelectionChanged: previousSelectedItem=NULL");
	}
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newSelectedItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectionChangingDrvCmbU(LPDISPATCH newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, CBLCtlsLibU::SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IDriveComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT("DrvCmbU_SelectionChanging: newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbU_SelectionChanging: newSelectedItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, selectionFieldHasBeenEdited=%i, selectionChangeReason=%i, cancelChange=%i"), selectionFieldText, selectionFieldHasBeenEdited, selectionChangeReason, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_XClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbU_XDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectionChangedImgCmbU(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pPrevItem = previousSelectedItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ImgCmbU_SelectionChanged: previousSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_SelectionChanged: previousSelectedItem=NULL");
	}
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newSelectedItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeforeDrawTextImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_BeforeDrawText"));
}

void __stdcall CMainDlg::BeginSelectionChangeImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_BeginSelectionChange"));
}

void __stdcall CMainDlg::ClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_Click: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_ContextMenu: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedComboBoxControlWindowImgCmbU(LONG hWndComboBox)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_CreatedComboBoxControlWindow: hWndComboBox=0x%X"), hWndComboBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowImgCmbU(LONG hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedListBoxControlWindowImgCmbU(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_CreatedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_DblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedComboBoxControlWindowImgCmbU(LONG hWndComboBox)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_DestroyedComboBoxControlWindow: hWndComboBox=0x%X"), hWndComboBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowImgCmbU(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_DestroyedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowImgCmbU(LONG hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedListBoxControlWindowImgCmbU(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_DestroyedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataImgCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_FreeItemData: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_FreeItemData: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemImgCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_InsertedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_InsertedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemImgCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IVirtualImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_InsertingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_InsertingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragImgCmbU(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ItemBeginDrag: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ItemBeginDrag: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), selectionFieldText, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginRDragImgCmbU(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ItemBeginRDrag: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ItemBeginRDrag: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), selectionFieldText, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoImgCmbU(LPDISPATCH comboItem, CBLCtlsLibU::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* selectedIconIndex, LONG* overlayIndex, LONG* indent, LONG maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		str.Format(TEXT("ImgCmbU_ItemGetDisplayInfo: comboItem=%i"), pItem->GetIndex());
	} else {
		str += TEXT("ImgCmbU_ItemGetDisplayInfo: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, iconIndex=%i, selectedIconIndex=%i, overlayIndex=%i, indent=%i, maxItemTextLength=%i, itemText=%s, dontAskAgain=%i"), requestedInfo, *iconIndex, *selectedIconIndex, *overlayIndex, *indent, maxItemTextLength, *itemText, *dontAskAgain);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ItemMouseEnter: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ItemMouseEnter: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ItemMouseLeave: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ItemMouseLeave: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownImgCmbU(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressImgCmbU(SHORT* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_KeyPress: keyAscii=%i"), *keyAscii);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpImgCmbU(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ListClickImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ListClick: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ListClick: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListCloseUpImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_ListCloseUp"));
}

void __stdcall CMainDlg::ListDropDownImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_ListDropDown"));
}

void __stdcall CMainDlg::ListMouseDownImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ListMouseDown: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ListMouseDown: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseMoveImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ListMouseMove: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ListMouseMove: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseUpImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ListMouseUp: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ListMouseUp: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseWheelImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_ListMouseWheel: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_ListMouseWheel: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragDropImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_ListOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_ListOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.imgcmbU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::ListOLEDragEnterImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_ListOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_ListOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragLeaveImgCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_ListOLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_ListOLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoCloseUp=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoCloseUp);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragMouseMoveImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_ListOLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_ListOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MouseDown: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MouseEnter: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MouseHover: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MouseLeave: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MouseMove: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MouseUp: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.imgcmbU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetImgCmbU(LONG hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveImgCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackImgCmbU(CBLCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragImgCmbU(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataImgCmbU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_OLEReceivedNewData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataImgCmbU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_OLESetData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragImgCmbU(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbU_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbU_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_OutOfMemory"));
}

void __stdcall CMainDlg::RClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_RClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_RDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowImgCmbU(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_RecreatedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
	InsertImageComboItemsU();
}

void __stdcall CMainDlg::RemovedItemImgCmbU(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IVirtualImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_RemovedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_RemovedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemImgCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbU_RemovingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_RemovingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_ResizedControlWindow"));
}

void __stdcall CMainDlg::SelectionCanceledImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_SelectionCanceled"));
}

void __stdcall CMainDlg::SelectionChangingImgCmbU(LPDISPATCH newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, CBLCtlsLibU::SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibU::IImageComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT("ImgCmbU_SelectionChanging: newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbU_SelectionChanging: newSelectedItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, selectionFieldHasBeenEdited=%i, selectionChangeReason=%i, cancelChange=%i"), selectionFieldText, selectionFieldHasBeenEdited, selectionChangeReason, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::TextChangedImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_TextChanged"));
}

void __stdcall CMainDlg::TruncatedTextImgCmbU(void)
{
	AddLogEntry(TEXT("ImgCmbU_TruncatedText"));
}

void __stdcall CMainDlg::WritingDirectionChangedImgCmbU(CBLCtlsLibU::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_WritingDirectionChanged: newWritingDirection=%i"), newWritingDirection);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_XClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbU_XDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CaretChangedLstA(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("LstA_CaretChanged: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_CaretChanged: previousCaretItem=NULL");
	}
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::AbortedDragLstA(void)
{
	AddLogEntry(TEXT("LstA_AbortedDrag"));
}

void __stdcall CMainDlg::ClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_Click: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_Click: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareItemsLstA(LPDISPATCH firstItem, LPDISPATCH secondItem, LONG locale, CBLCtlsLibA::CompareResultConstants* result)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("LstA_CompareItems: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_CompareItems: firstItem=NULL");
	}
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pSecondItem = secondItem;
	if(pSecondItem) {
		BSTR text = pSecondItem->GetText();
		str += TEXT(", secondItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", locale=%i, result=%i"), locale, *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_ContextMenu: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_ContextMenu: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_DblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_DblClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowLstA(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("LstA_DestroyedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveLstA(LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_DragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_DragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, autoHScrollVelocity, autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropLstA(LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_Drop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_Drop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataLstA(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_FreeItemData: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_FreeItemData: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemLstA(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_InsertedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_InsertedItem: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemLstA(LPDISPATCH listItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IVirtualListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_InsertingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_InsertingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_ItemBeginDrag: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_ItemBeginDrag: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginRDragLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_ItemBeginRDrag: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_ItemBeginRDrag: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoLstA(LPDISPATCH listItem, CBLCtlsLibA::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* overlayIndex)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_ItemGetDisplayInfo: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_ItemGetDisplayInfo: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, iconIndex=%i, overlayIndex=%i"), requestedInfo, *iconIndex, *overlayIndex);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetInfoTipTextLstA(LPDISPATCH listItem, LONG maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_ItemGetInfoTipText: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_ItemGetInfoTipText: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", maxInfoTipLength=%i, infoTipText=%s, abortToolTip=%i"), maxInfoTipLength, *infoTipText, *abortToolTip);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_ItemMouseEnter: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_ItemMouseEnter: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_ItemMouseLeave: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_ItemMouseLeave: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownLstA(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("LstA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressLstA(SHORT* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("LstA_KeyPress: keyAscii=%i"), *keyAscii);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpLstA(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("LstA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MDblClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MeasureItemLstA(LPDISPATCH listItem, OLE_YSIZE_PIXELS* itemHeight)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MeasureItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MeasureItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemHeight=%i"), *itemHeight);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MouseDown: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MouseDown: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MouseEnter: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MouseEnter: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MouseHover: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MouseHover: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MouseLeave: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MouseLeave: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MouseMove: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MouseMove: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MouseUp: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MouseUp: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_MouseWheel: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_MouseWheel: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragLstA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstA_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstA_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropLstA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.lstA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterLstA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoDropDown=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetLstA(LONG hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("LstA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveLstA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstALstA_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetLstA(void)
{
	AddLogEntry(TEXT("LstA_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveLstA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackLstA(CBLCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("LstA_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragLstA(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("LstA_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataLstA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstA_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstA_OLEReceivedNewData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataLstA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstA_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstA_OLESetData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragLstA(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("LstA_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("LstA_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryLstA(void)
{
	AddLogEntry(TEXT("LstA_OutOfMemory"));
}

void __stdcall CMainDlg::OwnerDrawItemLstA(LPDISPATCH listItem, CBLCtlsLibA::OwnerDrawActionConstants requiredAction, CBLCtlsLibA::OwnerDrawItemStateConstants itemState, LONG hDC, CBLCtlsLibA::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_OwnerDrawItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_OwnerDrawItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requiredAction=%i, itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), requiredAction, itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ProcessCharacterInputLstA(SHORT keyAscii, SHORT shift, LONG* itemToSelect)
{
	CAtlString str;
	str.Format(TEXT("LstA_ProcessCharacterInput: keyAscii=%i, shift=%i, itemToSelect=%i"), keyAscii, shift, *itemToSelect);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ProcessKeyStrokeLstA(SHORT keyCode, SHORT shift, LONG* itemToSelect)
{
	CAtlString str;
	str.Format(TEXT("LstA_ProcessKeyStroke: keyCode=%i, shift=%i, itemToSelect=%i"), keyCode, shift, *itemToSelect);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_RClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_RClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_RDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_RDblClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowLstA(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("LstA_RecreatedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
	InsertListItemsU();
}

void __stdcall CMainDlg::RemovedItemLstA(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IVirtualListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_RemovedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_RemovedItem: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemLstA(LPDISPATCH listItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_RemovingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_RemovingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowLstA(void)
{
	AddLogEntry(TEXT("LstA_ResizedControlWindow"));
}

void __stdcall CMainDlg::SelectionChangedLstA(void)
{
	AddLogEntry(TEXT("LstA_SelectionChanged"));
}

void __stdcall CMainDlg::XClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_XClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_XClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IListBoxItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("LstA_XDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("LstA_XDblClick: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectionChangedCmbA(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pPrevItem = previousSelectedItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("CmbA_SelectionChanged: previousSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_SelectionChanged: previousSelectedItem=NULL");
	}
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newSelectedItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeforeDrawTextCmbA(void)
{
	AddLogEntry(TEXT("CmbA_BeforeDrawText"));
}

void __stdcall CMainDlg::ClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_Click: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareItemsCmbA(LPDISPATCH firstItem, LPDISPATCH secondItem, LONG locale, CBLCtlsLibA::CompareResultConstants* result)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("CmbA_CompareItems: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_CompareItems: firstItem=NULL");
	}
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pSecondItem = secondItem;
	if(pSecondItem) {
		BSTR text = pSecondItem->GetText();
		str += TEXT(", secondItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", locale=%i, result=%i"), locale, *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("CmbA_ContextMenu: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowCmbA(LONG hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("CmbA_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedListBoxControlWindowCmbA(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("CmbA_CreatedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_DblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowCmbA(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("CmbA_DestroyedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowCmbA(LONG hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("CmbA_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedListBoxControlWindowCmbA(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("CmbA_DestroyedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_FreeItemData: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_FreeItemData: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_InsertedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_InsertedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IVirtualComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_InsertingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_InsertingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_ItemMouseEnter: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_ItemMouseEnter: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_ItemMouseLeave: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_ItemMouseLeave: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownCmbA(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("CmbA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressCmbA(SHORT* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("CmbA_KeyPress: keyAscii=%i"), *keyAscii);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpCmbA(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("CmbA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ListCloseUpCmbA(void)
{
	AddLogEntry(TEXT("CmbA_ListCloseUp"));
}

void __stdcall CMainDlg::ListDropDownCmbA(void)
{
	AddLogEntry(TEXT("CmbA_ListDropDown"));
}

void __stdcall CMainDlg::ListMouseDownCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_ListMouseDown: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_ListMouseDown: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseMoveCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_ListMouseMove: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_ListMouseMove: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseUpCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_ListMouseUp: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_ListMouseUp: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseWheelCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_ListMouseWheel: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_ListMouseWheel: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragDropCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbA_ListOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbA_ListOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.cmbA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::ListOLEDragEnterCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbA_ListOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbA_ListOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragLeaveCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbA_ListOLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbA_ListOLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoCloseUp=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoCloseUp);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragMouseMoveCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbA_ListOLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbA_ListOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MeasureItemCmbA(LPDISPATCH comboItem, OLE_YSIZE_PIXELS* itemHeight)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_MeasureItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_MeasureItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemHeight=%i"), *itemHeight);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MouseDown: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MouseEnter: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MouseHover: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MouseLeave: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MouseMove: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MouseUp: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.cmbA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbA_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmbA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmbA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryCmbA(void)
{
	AddLogEntry(TEXT("CmbA_OutOfMemory"));
}

void __stdcall CMainDlg::OwnerDrawItemCmbA(LPDISPATCH comboItem, CBLCtlsLibA::OwnerDrawActionConstants requiredAction, CBLCtlsLibA::OwnerDrawItemStateConstants itemState, LONG hDC, CBLCtlsLibA::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_OwnerDrawItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_OwnerDrawItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requiredAction=%i, itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), requiredAction, itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_RClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_RDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowCmbA(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("CmbA_RecreatedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
	InsertComboItemsA();
}

void __stdcall CMainDlg::RemovedItemCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IVirtualComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_RemovedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_RemovedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("CmbA_RemovingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("CmbA_RemovingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowCmbA(void)
{
	AddLogEntry(TEXT("CmbA_ResizedControlWindow"));
}

void __stdcall CMainDlg::SelectionCanceledCmbA(void)
{
	AddLogEntry(TEXT("CmbA_SelectionCanceled"));
}

void __stdcall CMainDlg::SelectionChangingCmbA(void)
{
	AddLogEntry(TEXT("CmbA_SelectionChanging"));
}

void __stdcall CMainDlg::TextChangedCmbA(void)
{
	AddLogEntry(TEXT("CmbA_TextChanged"));
}

void __stdcall CMainDlg::TruncatedTextCmbA(void)
{
	AddLogEntry(TEXT("CmbA_TruncatedText"));
}

void __stdcall CMainDlg::WritingDirectionChangedCmbA(CBLCtlsLibA::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("CmbA_WritingDirectionChanged: newWritingDirection=%i"), newWritingDirection);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_XClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("CmbA_XDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectedDriveChangedDrvCmbA(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pPrevItem = previousSelectedItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("DrvCmbA_SelectedDriveChanged: previousSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_SelectedDriveChanged: previousSelectedItem=NULL");
	}
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newSelectedItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginSelectionChangeDrvCmbA(void)
{
	AddLogEntry(TEXT("DrvCmbA_BeginSelectionChange"));
}

void __stdcall CMainDlg::ClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_Click: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_ContextMenu: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedComboBoxControlWindowDrvCmbA(LONG hWndComboBox)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_CreatedComboBoxControlWindow: hWndComboBox=0x%X"), hWndComboBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedListBoxControlWindowDrvCmbA(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_CreatedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_DblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedComboBoxControlWindowDrvCmbA(LONG hWndComboBox)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_DestroyedComboBoxControlWindow: hWndComboBox=0x%X"), hWndComboBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowDrvCmbA(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_DestroyedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedListBoxControlWindowDrvCmbA(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_DestroyedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataDrvCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_FreeItemData: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_FreeItemData: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemDrvCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_InsertedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_InsertedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemDrvCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IVirtualDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_InsertingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_InsertingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragDrvCmbA(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ItemBeginDrag: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ItemBeginDrag: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), selectionFieldText, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginRDragDrvCmbA(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ItemBeginRDrag: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ItemBeginRDrag: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), selectionFieldText, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoDrvCmbA(LPDISPATCH comboItem, CBLCtlsLibA::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* selectedIconIndex, LONG* overlayIndex, LONG* indent, LONG maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		str.Format(TEXT("DrvCmbA_ItemGetDisplayInfo: comboItem=%i"), pItem->GetIndex());
	} else {
		str += TEXT("DrvCmbA_ItemGetDisplayInfo: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, iconIndex=%i, selectedIconIndex=%i, overlayIndex=%i, indent=%i, maxItemTextLength=%i, itemText=%s, dontAskAgain=%i"), requestedInfo, *iconIndex, *selectedIconIndex, *overlayIndex, *indent, maxItemTextLength, *itemText, *dontAskAgain);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ItemMouseEnter: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ItemMouseEnter: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ItemMouseLeave: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ItemMouseLeave: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownDrvCmbA(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressDrvCmbA(SHORT* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_KeyPress: keyAscii=%i"), *keyAscii);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpDrvCmbA(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ListClickDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ListClick: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ListClick: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListCloseUpDrvCmbA(void)
{
	AddLogEntry(TEXT("DrvCmbA_ListCloseUp"));
}

void __stdcall CMainDlg::ListDropDownDrvCmbA(void)
{
	AddLogEntry(TEXT("DrvCmbA_ListDropDown"));
}

void __stdcall CMainDlg::ListMouseDownDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ListMouseDown: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ListMouseDown: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseMoveDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ListMouseMove: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ListMouseMove: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseUpDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ListMouseUp: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ListMouseUp: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseWheelDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_ListMouseWheel: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_ListMouseWheel: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragDropDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_ListOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_ListOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.drvcmbA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::ListOLEDragEnterDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_ListOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_ListOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragLeaveDrvCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_ListOLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_ListOLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoCloseUp=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoCloseUp);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragMouseMoveDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_ListOLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_ListOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MouseDown: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MouseEnter: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MouseHover: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MouseLeave: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MouseMove: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MouseUp: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.drvcmbA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetDrvCmbA(LONG hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveDrvCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetDrvCmbA(void)
{
	AddLogEntry(TEXT("DrvCmbA_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackDrvCmbA(CBLCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragDrvCmbA(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataDrvCmbA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_OLEReceivedNewData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataDrvCmbA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_OLESetData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragDrvCmbA(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("DrvCmbA_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("DrvCmbA_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryDrvCmbA(void)
{
	AddLogEntry(TEXT("DrvCmbA_OutOfMemory"));
}

void __stdcall CMainDlg::RClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_RClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_RDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowDrvCmbA(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_RecreatedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovedItemDrvCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IVirtualDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_RemovedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_RemovedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemDrvCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("DrvCmbA_RemovingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_RemovingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowDrvCmbA(void)
{
	AddLogEntry(TEXT("DrvCmbA_ResizedControlWindow"));
}

void __stdcall CMainDlg::SelectionCanceledDrvCmbA(void)
{
	AddLogEntry(TEXT("DrvCmbA_SelectionCanceled"));
}

void __stdcall CMainDlg::SelectionChangedDrvCmbA(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pPrevItem = previousSelectedItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("DrvCmbA_SelectionChanged: previousSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_SelectionChanged: previousSelectedItem=NULL");
	}
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newSelectedItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectionChangingDrvCmbA(LPDISPATCH newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, CBLCtlsLibA::SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IDriveComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT("DrvCmbA_SelectionChanging: newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("DrvCmbA_SelectionChanging: newSelectedItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, selectionFieldHasBeenEdited=%i, selectionChangeReason=%i, cancelChange=%i"), selectionFieldText, selectionFieldHasBeenEdited, selectionChangeReason, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_XClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("DrvCmbA_XDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectionChangedImgCmbA(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pPrevItem = previousSelectedItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ImgCmbA_SelectionChanged: previousSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_SelectionChanged: previousSelectedItem=NULL");
	}
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newSelectedItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeforeDrawTextImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_BeforeDrawText"));
}

void __stdcall CMainDlg::BeginSelectionChangeImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_BeginSelectionChange"));
}

void __stdcall CMainDlg::ClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_Click: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_ContextMenu: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedComboBoxControlWindowImgCmbA(LONG hWndComboBox)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_CreatedComboBoxControlWindow: hWndComboBox=0x%X"), hWndComboBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowImgCmbA(LONG hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);
	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedListBoxControlWindowImgCmbA(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_CreatedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_DblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedComboBoxControlWindowImgCmbA(LONG hWndComboBox)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_DestroyedComboBoxControlWindow: hWndComboBox=0x%X"), hWndComboBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowImgCmbA(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_DestroyedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowImgCmbA(LONG hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);
	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedListBoxControlWindowImgCmbA(LONG hWndListBox)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_DestroyedListBoxControlWindow: hWndListBox=0x%X"), hWndListBox);
	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataImgCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_FreeItemData: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_FreeItemData: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemImgCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_InsertedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_InsertedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemImgCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IVirtualImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_InsertingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_InsertingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragImgCmbA(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ItemBeginDrag: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ItemBeginDrag: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), selectionFieldText, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginRDragImgCmbA(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ItemBeginRDrag: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ItemBeginRDrag: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), selectionFieldText, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoImgCmbA(LPDISPATCH comboItem, CBLCtlsLibA::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* selectedIconIndex, LONG* overlayIndex, LONG* indent, LONG maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		str.Format(TEXT("ImgCmbA_ItemGetDisplayInfo: comboItem=%i"), pItem->GetIndex());
	} else {
		str += TEXT("ImgCmbA_ItemGetDisplayInfo: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, iconIndex=%i, selectedIconIndex=%i, overlayIndex=%i, indent=%i, maxItemTextLength=%i, itemText=%s, dontAskAgain=%i"), requestedInfo, *iconIndex, *selectedIconIndex, *overlayIndex, *indent, maxItemTextLength, *itemText, *dontAskAgain);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ItemMouseEnter: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ItemMouseEnter: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ItemMouseLeave: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ItemMouseLeave: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownImgCmbA(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressImgCmbA(SHORT* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_KeyPress: keyAscii=%i"), *keyAscii);
	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpImgCmbA(SHORT* keyCode, SHORT shift)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);
	AddLogEntry(str);
}

void __stdcall CMainDlg::ListClickImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ListClick: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ListClick: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListCloseUpImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_ListCloseUp"));
}

void __stdcall CMainDlg::ListDropDownImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_ListDropDown"));
}

void __stdcall CMainDlg::ListMouseDownImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ListMouseDown: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ListMouseDown: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseMoveImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ListMouseMove: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ListMouseMove: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseUpImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ListMouseUp: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ListMouseUp: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListMouseWheelImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_ListMouseWheel: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_ListMouseWheel: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragDropImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_ListOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_ListOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X"), button, shift, x, y, yToItemTop, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.imgcmbA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::ListOLEDragEnterImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_ListOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_ListOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragLeaveImgCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_ListOLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_ListOLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoCloseUp=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoCloseUp);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ListOLEDragMouseMoveImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_ListOLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_ListOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, yToItemTop=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, yToItemTop, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MouseDown: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MouseEnter: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MouseHover: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MouseLeave: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MouseMove: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MouseUp: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i, hitTestDetails=0x%X"), button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.imgcmbA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetImgCmbA(LONG hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveImgCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoDropDown=%i"), button, shift, x, y, hitTestDetails, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackImgCmbA(CBLCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragImgCmbA(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);
	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataImgCmbA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_OLEReceivedNewData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataImgCmbA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_OLESetData: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragImgCmbA(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ImgCmbA_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ImgCmbA_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_OutOfMemory"));
}

void __stdcall CMainDlg::RClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_RClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_RDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowImgCmbA(LONG hWnd)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_RecreatedControlWindow: hWnd=0x%X"), hWnd);
	AddLogEntry(str);
	InsertImageComboItemsA();
}

void __stdcall CMainDlg::RemovedItemImgCmbA(LPDISPATCH comboItem)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IVirtualImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_RemovedItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_RemovedItem: comboItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemImgCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pItem = comboItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ImgCmbA_RemovingItem: comboItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_RemovingItem: comboItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_ResizedControlWindow"));
}

void __stdcall CMainDlg::SelectionCanceledImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_SelectionCanceled"));
}

void __stdcall CMainDlg::SelectionChangingImgCmbA(LPDISPATCH newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, CBLCtlsLibA::SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<CBLCtlsLibA::IImageComboBoxItem> pNewItem = newSelectedItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT("ImgCmbA_SelectionChanging: newSelectedItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ImgCmbA_SelectionChanging: newSelectedItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", selectionFieldText=%s, selectionFieldHasBeenEdited=%i, selectionChangeReason=%i, cancelChange=%i"), selectionFieldText, selectionFieldHasBeenEdited, selectionChangeReason, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::TextChangedImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_TextChanged"));
}

void __stdcall CMainDlg::TruncatedTextImgCmbA(void)
{
	AddLogEntry(TEXT("ImgCmbA_TruncatedText"));
}

void __stdcall CMainDlg::WritingDirectionChangedImgCmbA(CBLCtlsLibA::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_WritingDirectionChanged: newWritingDirection=%i"), newWritingDirection);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_XClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ImgCmbA_XDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), button, shift, x, y, hitTestDetails);
	AddLogEntry(str);
}
