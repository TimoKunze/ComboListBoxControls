// CommonProperties.cpp: The controls' "Common" property page

#include "stdafx.h"
#include "CommonProperties.h"


CommonProperties::CommonProperties()
{
	m_dwTitleID = IDS_TITLECOMMONPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGCOMMONPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP CommonProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<CommonProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.disabledEventsList.SubclassWindow(GetDlgItem(IDC_DISABLEDEVENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.disabledEventsList.GetImageList(LVSIL_STATE));
	controls.disabledEventsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.disabledEventsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.disabledEventsList.AddColumn(TEXT(""), 0);
	controls.disabledEventsList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected events"));

	// setup the toolbar
	CRect toolbarRect;
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
	return S_OK;
}

STDMETHODIMP CommonProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<CommonProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT CommonProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(pNotificationDetails->hwndFrom != controls.disabledEventsList) {
		return 0;
	}

	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	if(pNotificationDetails->hwndFrom == controls.disabledEventsList) {
		if(pDetails->iItem == properties.disabledEventsItemIndices.deMouseEvents) {
			if(properties.numberOfListBoxes > 0) {
				if(properties.TotalNumberOfComboBoxes() > 0) {
					ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, ListBox.ItemMouseEnter, ListBox.ItemMouseLeave, MouseMove, MouseWheel"))));
				} else {
					ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, ItemMouseEnter, ItemMouseLeave, MouseMove, MouseWheel"))));
				}
			} else {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, MouseMove, MouseWheel"))));
			}
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deListBoxMouseEvents) {
			if(properties.numberOfListBoxes > 0) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ListMouseDown, ListMouseUp, *ComboBox.ItemMouseEnter, *ComboBox.ItemMouseLeave, ListMouseMove, ListMouseWheel"))));
			} else {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ListMouseDown, ListMouseUp, ItemMouseEnter, ItemMouseLeave, ListMouseMove, ListMouseWheel"))));
			}
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deClickEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Click, DblClick, MClick, MDblClick, RClick, RDblClick, XClick, XDblClick"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deListBoxClickEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ListClick"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deKeyboardEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("KeyDown, KeyUp, KeyPress"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deItemInsertionEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertingItem, InsertedItem"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deItemDeletionEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingItem, RemovedItem"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deFreeItemData) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FreeItemData"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deBeforeDrawText) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("BeforeDrawText"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deTextChangedEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("TextChanged"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deSelectionChangingEvents) {
			if(properties.numberOfDriveComboBoxes + properties.numberOfImageComboBoxes > 0) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("BeginSelectionChange, SelectionCanceled, SelectionChanging"))));
			} else {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("SelectionCanceled, SelectionChanging"))));
			}
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deProcessKeyboardInput) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ProcessCharacterInput, ProcessKeyStroke"))));
		}
	}
	ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	delete[] pBuffer;
	return 0;
}

LRESULT CommonProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT CommonProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IComboBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IDriveComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
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

	} else if(m_ppUnk[0]->QueryInterface(IID_IImageComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IImageComboBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IListBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IListBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT CommonProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("ComboBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IComboBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IDriveComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
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

	} else if(m_ppUnk[0]->QueryInterface(IID_IImageComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("ImageComboBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IImageComboBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IListBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("ListBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IListBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}


void CommonProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IComboBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deListBoxMouseEvents, LVIS_STATEIMAGEMASK), l, deListBoxMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deItemInsertionEvents, LVIS_STATEIMAGEMASK), l, deItemInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deItemDeletionEvents, LVIS_STATEIMAGEMASK), l, deItemDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deFreeItemData, LVIS_STATEIMAGEMASK), l, deFreeItemData);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deBeforeDrawText, LVIS_STATEIMAGEMASK), l, deBeforeDrawText);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deTextChangedEvents, LVIS_STATEIMAGEMASK), l, deTextChangedEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deSelectionChangingEvents, LVIS_STATEIMAGEMASK), l, deSelectionChangingEvents);
			reinterpret_cast<IComboBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IDriveComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IDriveComboBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deListBoxMouseEvents, LVIS_STATEIMAGEMASK), l, deListBoxMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deListBoxClickEvents, LVIS_STATEIMAGEMASK), l, deListBoxClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deItemInsertionEvents, LVIS_STATEIMAGEMASK), l, deItemInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deItemDeletionEvents, LVIS_STATEIMAGEMASK), l, deItemDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deFreeItemData, LVIS_STATEIMAGEMASK), l, deFreeItemData);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deSelectionChangingEvents, LVIS_STATEIMAGEMASK), l, deSelectionChangingEvents);
			reinterpret_cast<IDriveComboBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IImageComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IImageComboBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deListBoxMouseEvents, LVIS_STATEIMAGEMASK), l, deListBoxMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deListBoxClickEvents, LVIS_STATEIMAGEMASK), l, deListBoxClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deItemInsertionEvents, LVIS_STATEIMAGEMASK), l, deItemInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deItemDeletionEvents, LVIS_STATEIMAGEMASK), l, deItemDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deFreeItemData, LVIS_STATEIMAGEMASK), l, deFreeItemData);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deBeforeDrawText, LVIS_STATEIMAGEMASK), l, deBeforeDrawText);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deTextChangedEvents, LVIS_STATEIMAGEMASK), l, deTextChangedEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deSelectionChangingEvents, LVIS_STATEIMAGEMASK), l, deSelectionChangingEvents);
			reinterpret_cast<IImageComboBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IListBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IListBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deItemInsertionEvents, LVIS_STATEIMAGEMASK), l, deItemInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deItemDeletionEvents, LVIS_STATEIMAGEMASK), l, deItemDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deFreeItemData, LVIS_STATEIMAGEMASK), l, deFreeItemData);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deProcessKeyboardInput, LVIS_STATEIMAGEMASK), l, deProcessKeyboardInput);
			reinterpret_cast<IListBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void CommonProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	properties.numberOfComboBoxes = 0;
	properties.numberOfDriveComboBoxes = 0;
	properties.numberOfImageComboBoxes = 0;
	properties.numberOfListBoxes = 0;
	DisabledEventsConstants* pDisabledEvents = static_cast<DisabledEventsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DisabledEventsConstants)));
	if(pDisabledEvents) {
		ZeroMemory(pDisabledEvents, m_nObjects * sizeof(DisabledEventsConstants));
		for(UINT object = 0; object < m_nObjects; ++object) {
			LPUNKNOWN pControl = NULL;
			if(m_ppUnk[object]->QueryInterface(IID_IComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++properties.numberOfComboBoxes;
				reinterpret_cast<IComboBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
				pControl->Release();

			} else if(m_ppUnk[object]->QueryInterface(IID_IDriveComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++properties.numberOfDriveComboBoxes;
				reinterpret_cast<IDriveComboBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
				pControl->Release();

			} else if(m_ppUnk[object]->QueryInterface(IID_IImageComboBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++properties.numberOfImageComboBoxes;
				reinterpret_cast<IImageComboBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
				pControl->Release();

			} else if(m_ppUnk[object]->QueryInterface(IID_IListBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++properties.numberOfListBoxes;
				reinterpret_cast<IListBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
				pControl->Release();
			}
		}

		// fill the listbox
		properties.disabledEventsItemIndices.Reset();
		LONG* pl = reinterpret_cast<LONG*>(pDisabledEvents);
		properties.hWndCheckMarksAreSetFor = controls.disabledEventsList;
		controls.disabledEventsList.DeleteAllItems();
		properties.disabledEventsItemIndices.deMouseEvents = controls.disabledEventsList.AddItem(0, 0, TEXT("Mouse events"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deMouseEvents, CalculateStateImageMask(m_nObjects, pl, deMouseEvents), LVIS_STATEIMAGEMASK);
		if(properties.TotalNumberOfComboBoxes() > 0) {
			properties.disabledEventsItemIndices.deListBoxMouseEvents = controls.disabledEventsList.AddItem(1, 0, TEXT("Mouse events (drop-down list box)"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deListBoxMouseEvents, CalculateStateImageMask(m_nObjects, pl, deListBoxMouseEvents), LVIS_STATEIMAGEMASK);
		}
		properties.disabledEventsItemIndices.deClickEvents = controls.disabledEventsList.AddItem(2, 0, TEXT("Click events"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deClickEvents, CalculateStateImageMask(m_nObjects, pl, deClickEvents), LVIS_STATEIMAGEMASK);
		if(properties.numberOfDriveComboBoxes + properties.numberOfImageComboBoxes > 0) {
			properties.disabledEventsItemIndices.deListBoxClickEvents = controls.disabledEventsList.AddItem(3, 0, TEXT("Click events (drop-down list box)"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deListBoxClickEvents, CalculateStateImageMask(m_nObjects, pl, deListBoxClickEvents), LVIS_STATEIMAGEMASK);
		}
		properties.disabledEventsItemIndices.deKeyboardEvents = controls.disabledEventsList.AddItem(4, 0, TEXT("Keyboard events"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, CalculateStateImageMask(m_nObjects, pl, deKeyboardEvents), LVIS_STATEIMAGEMASK);
		properties.disabledEventsItemIndices.deItemInsertionEvents = controls.disabledEventsList.AddItem(5, 0, TEXT("Item insertions"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deItemInsertionEvents, CalculateStateImageMask(m_nObjects, pl, deItemInsertionEvents), LVIS_STATEIMAGEMASK);
		properties.disabledEventsItemIndices.deItemDeletionEvents = controls.disabledEventsList.AddItem(6, 0, TEXT("Item deletions"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deItemDeletionEvents, CalculateStateImageMask(m_nObjects, pl, deItemDeletionEvents), LVIS_STATEIMAGEMASK);
		properties.disabledEventsItemIndices.deFreeItemData = controls.disabledEventsList.AddItem(7, 0, TEXT("FreeItemData event"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deFreeItemData, CalculateStateImageMask(m_nObjects, pl, deFreeItemData), LVIS_STATEIMAGEMASK);
		if(properties.numberOfComboBoxes + properties.numberOfImageComboBoxes > 0) {
			properties.disabledEventsItemIndices.deBeforeDrawText = controls.disabledEventsList.AddItem(8, 0, TEXT("BeforeDrawText event"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deBeforeDrawText, CalculateStateImageMask(m_nObjects, pl, deBeforeDrawText), LVIS_STATEIMAGEMASK);
			properties.disabledEventsItemIndices.deTextChangedEvents = controls.disabledEventsList.AddItem(9, 0, TEXT("TextChanged event"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deTextChangedEvents, CalculateStateImageMask(m_nObjects, pl, deTextChangedEvents), LVIS_STATEIMAGEMASK);
		}
		if(properties.TotalNumberOfComboBoxes() > 0) {
			properties.disabledEventsItemIndices.deSelectionChangingEvents = controls.disabledEventsList.AddItem(10, 0, TEXT("Selection changes"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deSelectionChangingEvents, CalculateStateImageMask(m_nObjects, pl, deSelectionChangingEvents), LVIS_STATEIMAGEMASK);
		}
		if(properties.numberOfListBoxes > 0) {
			properties.disabledEventsItemIndices.deProcessKeyboardInput = controls.disabledEventsList.AddItem(11, 0, TEXT("Keyboard input processing"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deProcessKeyboardInput, CalculateStateImageMask(m_nObjects, pl, deProcessKeyboardInput), LVIS_STATEIMAGEMASK);
		}
		controls.disabledEventsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

		properties.hWndCheckMarksAreSetFor = NULL;

		HeapFree(GetProcessHeap(), 0, pDisabledEvents);
	}

	SetDirty(FALSE);
}

int CommonProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
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

void CommonProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
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