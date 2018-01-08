// DriveTypeProperties.cpp: The controls' "Drive Types" property page

#include "stdafx.h"
#include "DriveTypeProperties.h"


DriveTypeProperties::DriveTypeProperties()
{
	m_dwTitleID = IDS_TITLEDRIVETYPEPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGDRIVETYPEPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP DriveTypeProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP DriveTypeProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}

	IPropertyPageImpl<DriveTypeProperties>::SetObjects(objects, ppControls);

	if(properties.weAreInitialized) {
		LoadSettings();
	}
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT DriveTypeProperties::OnInitDialog(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	// attach to the controls
	controls.displayedDriveTypesList.SubclassWindow(GetDlgItem(IDC_DISPLAYEDDRIVETYPESBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.displayedDriveTypesList.GetImageList(LVSIL_STATE));
	controls.displayedDriveTypesList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.displayedDriveTypesList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.displayedDriveTypesList.AddColumn(TEXT(""), 0);
	controls.driveTypesWithVolumeNamesList.SubclassWindow(GetDlgItem(IDC_DRIVETYPESWITHVOLNAMESBOX));
	hStateImageList = SetupStateImageList(controls.driveTypesWithVolumeNamesList.GetImageList(LVSIL_STATE));
	controls.driveTypesWithVolumeNamesList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.driveTypesWithVolumeNamesList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.driveTypesWithVolumeNamesList.AddColumn(TEXT(""), 0);

	// setup the toolbar
	WTL::CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();

	properties.weAreInitialized = TRUE;
	return TRUE;
}


LRESULT DriveTypeProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uNewState & LVIS_STATEIMAGEMASK) != (pDetails->uOldState & LVIS_STATEIMAGEMASK)) {
			if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 3) {
				if(pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
					LVITEM item = {0};
					item.state = INDEXTOSTATEIMAGEMASK(1);
					item.stateMask = LVIS_STATEIMAGEMASK;
					::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem, reinterpret_cast<LPARAM>(&item));
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT DriveTypeProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT DriveTypeProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT DriveTypeProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IDriveComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IDriveComboBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT DriveTypeProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IDriveComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("DriveComboBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IDriveComboBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}


void DriveTypeProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IDriveComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DriveTypeConstants displayedDriveTypes = static_cast<DriveTypeConstants>(0);
			reinterpret_cast<IDriveComboBox*>(pControl)->get_DisplayedDriveTypes(&displayedDriveTypes);
			LONG l = static_cast<LONG>(displayedDriveTypes);
			SetBit(controls.displayedDriveTypesList.GetItemState(0, LVIS_STATEIMAGEMASK), l, dtUnknown);
			SetBit(controls.displayedDriveTypesList.GetItemState(1, LVIS_STATEIMAGEMASK), l, dtOther);
			SetBit(controls.displayedDriveTypesList.GetItemState(2, LVIS_STATEIMAGEMASK), l, dtFixed);
			SetBit(controls.displayedDriveTypesList.GetItemState(3, LVIS_STATEIMAGEMASK), l, dtOptical);
			SetBit(controls.displayedDriveTypesList.GetItemState(4, LVIS_STATEIMAGEMASK), l, dtRemovable);
			SetBit(controls.displayedDriveTypesList.GetItemState(5, LVIS_STATEIMAGEMASK), l, dtRemote);
			SetBit(controls.displayedDriveTypesList.GetItemState(6, LVIS_STATEIMAGEMASK), l, dtRAMDisk);
			reinterpret_cast<IDriveComboBox*>(pControl)->put_DisplayedDriveTypes(static_cast<DriveTypeConstants>(l));

			DriveTypeConstants driveTypesWithVolumeName = static_cast<DriveTypeConstants>(0);
			reinterpret_cast<IDriveComboBox*>(pControl)->get_DriveTypesWithVolumeName(&driveTypesWithVolumeName);
			l = static_cast<LONG>(driveTypesWithVolumeName);
			SetBit(controls.driveTypesWithVolumeNamesList.GetItemState(0, LVIS_STATEIMAGEMASK), l, dtUnknown);
			SetBit(controls.driveTypesWithVolumeNamesList.GetItemState(1, LVIS_STATEIMAGEMASK), l, dtOther);
			SetBit(controls.driveTypesWithVolumeNamesList.GetItemState(2, LVIS_STATEIMAGEMASK), l, dtFixed);
			SetBit(controls.driveTypesWithVolumeNamesList.GetItemState(3, LVIS_STATEIMAGEMASK), l, dtOptical);
			SetBit(controls.driveTypesWithVolumeNamesList.GetItemState(4, LVIS_STATEIMAGEMASK), l, dtRemovable);
			SetBit(controls.driveTypesWithVolumeNamesList.GetItemState(5, LVIS_STATEIMAGEMASK), l, dtRemote);
			SetBit(controls.driveTypesWithVolumeNamesList.GetItemState(6, LVIS_STATEIMAGEMASK), l, dtRAMDisk);
			reinterpret_cast<IDriveComboBox*>(pControl)->put_DriveTypesWithVolumeName(static_cast<DriveTypeConstants>(l));
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void DriveTypeProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	DriveTypeConstants* pDisplayedDriveTypes = static_cast<DriveTypeConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DriveTypeConstants)));
	if(pDisplayedDriveTypes) {
		ZeroMemory(pDisplayedDriveTypes, m_nObjects * sizeof(DriveTypeConstants));
		DriveTypeConstants* pDriveTypesWithVolumeName = static_cast<DriveTypeConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DriveTypeConstants)));
		if(pDriveTypesWithVolumeName) {
			ZeroMemory(pDriveTypesWithVolumeName, m_nObjects * sizeof(DriveTypeConstants));
			for(UINT object = 0; object < m_nObjects; ++object) {
				LPUNKNOWN pControl = NULL;
				if(m_ppUnk[object]->QueryInterface(IID_IDriveComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
					reinterpret_cast<IDriveComboBox*>(pControl)->get_DisplayedDriveTypes(&pDisplayedDriveTypes[object]);
					reinterpret_cast<IDriveComboBox*>(pControl)->get_DriveTypesWithVolumeName(&pDriveTypesWithVolumeName[object]);
					pControl->Release();
				}
			}

			// fill the listboxes
			LONG* pl = reinterpret_cast<LONG*>(pDisplayedDriveTypes);
			properties.hWndCheckMarksAreSetFor = controls.displayedDriveTypesList;
			controls.displayedDriveTypesList.DeleteAllItems();
			controls.displayedDriveTypesList.AddItem(0, 0, TEXT("Unknown"));
			controls.displayedDriveTypesList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, dtUnknown), LVIS_STATEIMAGEMASK);
			controls.displayedDriveTypesList.AddItem(1, 0, TEXT("Drives with invalid root path"));
			controls.displayedDriveTypesList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, dtOther), LVIS_STATEIMAGEMASK);
			controls.displayedDriveTypesList.AddItem(2, 0, TEXT("Fixed drives (e. g. hard disks)"));
			controls.displayedDriveTypesList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, dtFixed), LVIS_STATEIMAGEMASK);
			controls.displayedDriveTypesList.AddItem(3, 0, TEXT("Optical drives"));
			controls.displayedDriveTypesList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, dtOptical), LVIS_STATEIMAGEMASK);
			controls.displayedDriveTypesList.AddItem(4, 0, TEXT("Removable drives"));
			controls.displayedDriveTypesList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, dtRemovable), LVIS_STATEIMAGEMASK);
			controls.displayedDriveTypesList.AddItem(5, 0, TEXT("Remote drives"));
			controls.displayedDriveTypesList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, dtRemote), LVIS_STATEIMAGEMASK);
			controls.displayedDriveTypesList.AddItem(6, 0, TEXT("RAM disks"));
			controls.displayedDriveTypesList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, dtRAMDisk), LVIS_STATEIMAGEMASK);
			controls.displayedDriveTypesList.SetColumnWidth(0, LVSCW_AUTOSIZE);

			pl = reinterpret_cast<LONG*>(pDriveTypesWithVolumeName);
			properties.hWndCheckMarksAreSetFor = controls.driveTypesWithVolumeNamesList;
			controls.driveTypesWithVolumeNamesList.DeleteAllItems();
			controls.driveTypesWithVolumeNamesList.AddItem(0, 0, TEXT("Unknown"));
			controls.driveTypesWithVolumeNamesList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, dtUnknown), LVIS_STATEIMAGEMASK);
			controls.driveTypesWithVolumeNamesList.AddItem(1, 0, TEXT("Drives with invalid root path"));
			controls.driveTypesWithVolumeNamesList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, dtOther), LVIS_STATEIMAGEMASK);
			controls.driveTypesWithVolumeNamesList.AddItem(2, 0, TEXT("Fixed drives (e. g. hard disks)"));
			controls.driveTypesWithVolumeNamesList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, dtFixed), LVIS_STATEIMAGEMASK);
			controls.driveTypesWithVolumeNamesList.AddItem(3, 0, TEXT("Optical drives"));
			controls.driveTypesWithVolumeNamesList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, dtOptical), LVIS_STATEIMAGEMASK);
			controls.driveTypesWithVolumeNamesList.AddItem(4, 0, TEXT("Removable drives"));
			controls.driveTypesWithVolumeNamesList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, dtRemovable), LVIS_STATEIMAGEMASK);
			controls.driveTypesWithVolumeNamesList.AddItem(5, 0, TEXT("Remote drives"));
			controls.driveTypesWithVolumeNamesList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, dtRemote), LVIS_STATEIMAGEMASK);
			controls.driveTypesWithVolumeNamesList.AddItem(6, 0, TEXT("RAM disks"));
			controls.driveTypesWithVolumeNamesList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, dtRAMDisk), LVIS_STATEIMAGEMASK);
			controls.driveTypesWithVolumeNamesList.SetColumnWidth(0, LVSCW_AUTOSIZE);

			properties.hWndCheckMarksAreSetFor = NULL;

			HeapFree(GetProcessHeap(), 0, pDriveTypesWithVolumeName);
		}
		HeapFree(GetProcessHeap(), 0, pDisplayedDriveTypes);
	}

	SetDirty(FALSE);
}

int DriveTypeProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if(pArray[object] & bitsToCheckFor) {
			if(stateImageIndex == 1) {
				stateImageIndex = (object == 0 ? 2 : 3);
			}
		} else {
			if(stateImageIndex == 2) {
				stateImageIndex = (object == 0 ? 1 : 3);
			}
		}
	}

	return INDEXTOSTATEIMAGEMASK(stateImageIndex);
}

void DriveTypeProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
{
	stateImageMask >>= 12;
	switch(stateImageMask) {
		case 1:
			value &= ~bitToSet;
			break;
		case 2:
			value |= bitToSet;
			break;
	}
}