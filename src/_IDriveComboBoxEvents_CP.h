//////////////////////////////////////////////////////////////////////
/// \class Proxy_IDriveComboBoxEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IDriveComboBoxEvents interface</em>
///
/// \if UNICODE
///   \sa DriveComboBox, CBLCtlsLibU::_IDriveComboBoxEvents
/// \else
///   \sa DriveComboBox, CBLCtlsLibA::_IDriveComboBoxEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IDriveComboBoxEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IDriveComboBoxEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c BeginSelectionChange event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::BeginSelectionChange,
	///       DriveComboBox::Raise_BeginSelectionChange, Fire_SelectionChanging
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::BeginSelectionChange,
	///       DriveComboBox::Raise_BeginSelectionChange, Fire_SelectionChanging
	/// \endif
	HRESULT Fire_BeginSelectionChange(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DCBE_BEGINSELECTIONCHANGE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Click event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::Click, DriveComboBox::Raise_Click, Fire_DblClick,
	///       Fire_MClick, Fire_RClick, Fire_XClick, Fire_ListClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::Click, DriveComboBox::Raise_Click, Fire_DblClick,
	///       Fire_MClick, Fire_RClick, Fire_XClick, Fire_ListClick
	/// \endif
	HRESULT Fire_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the menu's proposed position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the control from
	///                showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ContextMenu, DriveComboBox::Raise_ContextMenu, Fire_RClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ContextMenu, DriveComboBox::Raise_ContextMenu, Fire_RClick
	/// \endif
	HRESULT Fire_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pShowDefaultMenu;									p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DCBE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedComboBoxControlWindow event</em>
	///
	/// \param[in] hWndComboBox The combo box control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::CreatedComboBoxControlWindow,
	///       DriveComboBox::Raise_CreatedComboBoxControlWindow, Fire_DestroyedComboBoxControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::CreatedComboBoxControlWindow,
	///       DriveComboBox::Raise_CreatedComboBoxControlWindow, Fire_DestroyedComboBoxControlWindow
	/// \endif
	HRESULT Fire_CreatedComboBoxControlWindow(LONG hWndComboBox)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndComboBox;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_CREATEDCOMBOBOXCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedListBoxControlWindow event</em>
	///
	/// \param[in] hWndListBox The list box control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::CreatedListBoxControlWindow,
	///       DriveComboBox::Raise_CreatedListBoxControlWindow, Fire_DestroyedListBoxControlWindow,
	///       Fire_ListDropDown
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::CreatedListBoxControlWindow,
	///       DriveComboBox::Raise_CreatedListBoxControlWindow, Fire_DestroyedListBoxControlWindow,
	///       Fire_ListDropDown
	/// \endif
	HRESULT Fire_CreatedListBoxControlWindow(LONG hWndListBox)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndListBox;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_CREATEDLISTBOXCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::DblClick, DriveComboBox::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::DblClick, DriveComboBox::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedComboBoxControlWindow event</em>
	///
	/// \param[in] hWndComboBox The combo box control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::DestroyedComboBoxControlWindow,
	///       DriveComboBox::Raise_DestroyedComboBoxControlWindow, Fire_CreatedComboBoxControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::DestroyedComboBoxControlWindow,
	///       DriveComboBox::Raise_DestroyedComboBoxControlWindow, Fire_CreatedComboBoxControlWindow
	/// \endif
	HRESULT Fire_DestroyedComboBoxControlWindow(LONG hWndComboBox)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndComboBox;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_DESTROYEDCOMBOBOXCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::DestroyedControlWindow,
	///       DriveComboBox::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::DestroyedControlWindow,
	///       DriveComboBox::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \endif
	HRESULT Fire_DestroyedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedListBoxControlWindow event</em>
	///
	/// \param[in] hWndListBox The list box control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::DestroyedListBoxControlWindow,
	///       DriveComboBox::Raise_DestroyedListBoxControlWindow, Fire_CreatedListBoxControlWindow,
	///       Fire_ListCloseUp
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::DestroyedListBoxControlWindow,
	///       DriveComboBox::Raise_DestroyedListBoxControlWindow, Fire_CreatedListBoxControlWindow,
	///       Fire_ListCloseUp
	/// \endif
	HRESULT Fire_DestroyedListBoxControlWindow(LONG hWndListBox)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndListBox;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_DESTROYEDLISTBOXCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeItemData event</em>
	///
	/// \param[in] pComboItem The item whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::FreeItemData, DriveComboBox::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::FreeItemData, DriveComboBox::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \endif
	HRESULT Fire_FreeItemData(IDriveComboBoxItem* pComboItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pComboItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_FREEITEMDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedItem event</em>
	///
	/// \param[in] pComboItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::InsertedItem, DriveComboBox::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_RemovedItem
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::InsertedItem, DriveComboBox::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_RemovedItem
	/// \endif
	HRESULT Fire_InsertedItem(IDriveComboBoxItem* pComboItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pComboItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_INSERTEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingItem event</em>
	///
	/// \param[in] pComboItem The item being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::InsertingItem, DriveComboBox::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_RemovingItem
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::InsertingItem, DriveComboBox::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertingItem(IVirtualDriveComboBoxItem* pComboItem, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pComboItem;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DCBE_INSERTINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemBeginDrag event</em>
	///
	/// \param[in] pComboItem The item that the user wants to drag. May be \c NULL, indicating that the
	///            content of the selection field shall be dragged.
	/// \param[in] selectionFieldText The text currently being displayed in the selection field.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ItemBeginDrag, DriveComboBox::Raise_ItemBeginDrag,
	///       Fire_ItemBeginRDrag
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ItemBeginDrag, DriveComboBox::Raise_ItemBeginDrag,
	///       Fire_ItemBeginRDrag
	/// \endif
	HRESULT Fire_ItemBeginDrag(IDriveComboBoxItem* pComboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pComboItem;
				p[5] = selectionFieldText;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_DCBE_ITEMBEGINDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemBeginRDrag event</em>
	///
	/// \param[in] pComboItem The item that the user wants to drag. May be \c NULL, indicating that the
	///            content of the selection field shall be dragged.
	/// \param[in] selectionFieldText The text currently being displayed in the selection field.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbRightButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ItemBeginRDrag, DriveComboBox::Raise_ItemBeginRDrag,
	///       Fire_ItemBeginDrag
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ItemBeginRDrag, DriveComboBox::Raise_ItemBeginRDrag,
	///       Fire_ItemBeginDrag
	/// \endif
	HRESULT Fire_ItemBeginRDrag(IDriveComboBoxItem* pComboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pComboItem;
				p[5] = selectionFieldText;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_DCBE_ITEMBEGINRDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetDisplayInfo event</em>
	///
	/// \param[in] pComboItem The item whose properties are requested.
	/// \param[in] requestedInfo Specifies which properties' values are required. Any combination of
	///            the values defined by the \c RequestedInfoConstants enumeration is valid.
	/// \param[out] pIconIndex The zero-based index of the item's icon. If the \c requestedInfo parameter
	///             doesn't include \c riIconIndex, the caller should ignore this value.
	/// \param[out] pSelectedIconIndex The zero-based index of the requested selected icon. If the
	///             \c requestedInfo parameter doesn't include \c riSelectedIconIndex, the caller should
	///             ignore this value.
	/// \param[out] pOverlayIndex The zero-based index of the item's overlay icon. If the \c requestedInfo
	///             parameter doesn't include \c riOverlayIndex, the caller should ignore this value.
	/// \param[out] pIndent The requested indentation. If the \c requestedInfo parameter doesn't include
	///             \c riIndent, the caller should ignore this value.
	/// \param[in] maxItemTextLength The maximum number of characters the item's text may consist of. If the
	///            \c requestedInfo parameter doesn't include \c riItemText, the client should ignore this
	///            value.
	/// \param[out] pItemText The item's text. If the \c requestedInfo parameter doesn't include
	///             \c riItemText, the caller should ignore this value.
	/// \param[in,out] pDontAskAgain If \c VARIANT_TRUE, the caller should always use the same settings
	///                and never fire this event again for these properties of this item; otherwise it
	///                shouldn't make the values persistent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ItemGetDisplayInfo, DriveComboBox::Raise_ItemGetDisplayInfo
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ItemGetDisplayInfo, DriveComboBox::Raise_ItemGetDisplayInfo
	/// \endif
	HRESULT Fire_ItemGetDisplayInfo(IDriveComboBoxItem* pComboItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pSelectedIconIndex, LONG* pOverlayIndex, LONG* pIndent, LONG maxItemTextLength, BSTR* pItemText, VARIANT_BOOL* pDontAskAgain)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pComboItem;
				p[7] = static_cast<LONG>(requestedInfo);		p[7].vt = VT_I4;
				p[6].plVal = pIconIndex;										p[6].vt = VT_I4 | VT_BYREF;
				p[5].plVal = pSelectedIconIndex;						p[5].vt = VT_I4 | VT_BYREF;
				p[4].plVal = pOverlayIndex;									p[4].vt = VT_I4 | VT_BYREF;
				p[3].plVal = pIndent;												p[3].vt = VT_I4 | VT_BYREF;
				p[2] = maxItemTextLength;
				p[1].pbstrVal = pItemText;									p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pDontAskAgain;							p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_DCBE_ITEMGETDISPLAYINFO, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemMouseEnter event</em>
	///
	/// \param[in] pComboItem The item that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Most of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ItemMouseEnter, DriveComboBox::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_ListMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ItemMouseEnter, DriveComboBox::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_ListMouseMove
	/// \endif
	HRESULT Fire_ItemMouseEnter(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pComboItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DCBE_ITEMMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemMouseLeave event</em>
	///
	/// \param[in] pComboItem The item that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Most of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ItemMouseLeave, DriveComboBox::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_ListMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ItemMouseLeave, DriveComboBox::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_ListMouseMove
	/// \endif
	HRESULT Fire_ItemMouseLeave(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pComboItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DCBE_ITEMMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::KeyDown, DriveComboBox::Raise_KeyDown, Fire_KeyUp,
	///       Fire_KeyPress
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::KeyDown, DriveComboBox::Raise_KeyDown, Fire_KeyUp,
	///       Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DCBE_KEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::KeyPress, DriveComboBox::Raise_KeyPress, Fire_KeyDown,
	///       Fire_KeyUp
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::KeyPress, DriveComboBox::Raise_KeyPress, Fire_KeyDown,
	///       Fire_KeyUp
	/// \endif
	HRESULT Fire_KeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_KEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::KeyUp, DriveComboBox::Raise_KeyUp, Fire_KeyDown,
	///       Fire_KeyPress
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::KeyUp, DriveComboBox::Raise_KeyUp, Fire_KeyDown,
	///       Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DCBE_KEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListClick event</em>
	///
	/// \param[in] pComboItem The clicked item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down list box
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down list box
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that was clicked. Any of
	///            the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ListClick, DriveComboBox::Raise_ListClick, Fire_Click
	///      
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ListClick, DriveComboBox::Raise_ListClick, Fire_Click
	/// \endif
	HRESULT Fire_ListClick(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pComboItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListCloseUp event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ListCloseUp, DriveComboBox::Raise_ListCloseUp
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ListCloseUp, DriveComboBox::Raise_ListCloseUp
	/// \endif
	HRESULT Fire_ListCloseUp(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTCLOSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListDropDown event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ListDropDown, DriveComboBox::Raise_ListDropDown
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ListDropDown, DriveComboBox::Raise_ListDropDown
	/// \endif
	HRESULT Fire_ListDropDown(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTDROPDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListMouseDown event</em>
	///
	/// \param[in] pComboItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ListMouseDown, DriveComboBox::Raise_ListMouseDown,
	///       Fire_ListMouseUp, Fire_ListClick, Fire_MouseDown
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ListMouseDown, DriveComboBox::Raise_ListMouseDown,
	///       Fire_ListMouseUp, Fire_ListClick, Fire_MouseDown
	/// \endif
	HRESULT Fire_ListMouseDown(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pComboItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTMOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListMouseMove event</em>
	///
	/// \param[in] pComboItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ListMouseMove, DriveComboBox::Raise_ListMouseMove,
	///       Fire_ListMouseDown, Fire_ListMouseUp, Fire_ListMouseWheel, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ListMouseMove, DriveComboBox::Raise_ListMouseMove,
	///       Fire_ListMouseDown, Fire_ListMouseUp, Fire_ListMouseWheel, Fire_MouseMove
	/// \endif
	HRESULT Fire_ListMouseMove(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pComboItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListMouseUp event</em>
	///
	/// \param[in] pComboItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ListMouseUp, DriveComboBox::Raise_ListMouseUp,
	///       Fire_ListMouseDown, Fire_ListClick, Fire_MouseUp
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ListMouseUp, DriveComboBox::Raise_ListMouseUp,
	///       Fire_ListMouseDown, Fire_ListClick, Fire_MouseUp
	/// \endif
	HRESULT Fire_ListMouseUp(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pComboItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTMOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListMouseWheel event</em>
	///
	/// \param[in] pComboItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ListMouseWheel, DriveComboBox::Raise_ListMouseWheel,
	///       Fire_ListMouseMove, Fire_MouseWheel
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ListMouseWheel, DriveComboBox::Raise_ListMouseWheel,
	///       Fire_ListMouseMove, Fire_MouseWheel
	/// \endif
	HRESULT Fire_ListMouseWheel(IDriveComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pComboItem;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(scrollAxis);				p[2].vt = VT_I4;
				p[1] = wheelDelta;																p[1].vt = VT_I2;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTMOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListOLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in,out] ppDropTarget The item that is the target of the drag'n'drop operation. The client
	///                may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragDrop, DriveComboBox::Raise_ListOLEDragDrop,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragMouseMove, Fire_ListOLEDragLeave, Fire_ListMouseUp,
	///       Fire_OLEDragDrop
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragDrop, DriveComboBox::Raise_ListOLEDragDrop,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragMouseMove, Fire_ListOLEDragLeave, Fire_ListMouseUp,
	///       Fire_OLEDragDrop
	/// \endif
	HRESULT Fire_ListOLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IDriveComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7].plVal = reinterpret_cast<PLONG>(pEffect);									p[7].vt = VT_I4 | VT_BYREF;
				p[6].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[6].vt = VT_DISPATCH | VT_BYREF;
				p[5] = button;																									p[5].vt = VT_I2;
				p[4] = shift;																										p[4].vt = VT_I2;
				p[3] = x;																												p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																												p[2].vt = VT_YPOS_PIXELS;
				p[1] = yToItemTop;																							p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(hitTestDetails);									p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTOLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListOLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the drop-down list box control to the left; if set to a value greater
	///                than 0, it should scroll the drop-down list box control to the right. The higher/lower
	///                the value is, the faster the caller should scroll the drop-down list box control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the drop-down list box control upwardly; if set to a value greater than
	///                0, it should scroll the drop-down list box control downwards. The higher/lower the
	///                value is, the faster the caller should scroll the drop-down list box control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragEnter, DriveComboBox::Raise_ListOLEDragEnter,
	///       Fire_ListOLEDragMouseMove, Fire_ListOLEDragLeave, Fire_ListOLEDragDrop, Fire_OLEDragEnter
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragEnter, DriveComboBox::Raise_ListOLEDragEnter,
	///       Fire_ListOLEDragMouseMove, Fire_ListOLEDragLeave, Fire_ListOLEDragDrop, Fire_OLEDragEnter
	/// \endif
	HRESULT Fire_ListOLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IDriveComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[11];
				p[10] = pData;
				p[9].plVal = reinterpret_cast<PLONG>(pEffect);									p[9].vt = VT_I4 | VT_BYREF;
				p[8].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[8].vt = VT_DISPATCH | VT_BYREF;
				p[7] = button;																									p[7].vt = VT_I2;
				p[6] = shift;																										p[6].vt = VT_I2;
				p[5] = x;																												p[5].vt = VT_XPOS_PIXELS;
				p[4] = y;																												p[4].vt = VT_YPOS_PIXELS;
				p[3] = yToItemTop;																							p[3].vt = VT_I4;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 11, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTOLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListOLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] pDropTarget The item that is the current target of the drag'n'drop operation.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	/// \param[in,out] pAutoCloseUp If set to \c VARIANT_TRUE, the caller should close the drop-down list box
	///                control; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragLeave, DriveComboBox::Raise_ListOLEDragLeave,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragMouseMove, Fire_ListOLEDragDrop, Fire_OLEDragLeave
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragLeave, DriveComboBox::Raise_ListOLEDragLeave,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragMouseMove, Fire_ListOLEDragDrop, Fire_OLEDragLeave
	/// \endif
	HRESULT Fire_ListOLEDragLeave(IOLEDataObject* pData, IDriveComboBoxItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoCloseUp)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7] = pDropTarget;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2] = yToItemTop;																p[2].vt = VT_I4;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pAutoCloseUp;											p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTOLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListOLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the drop-down list box control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the drop-down list box control to the left; if set to a value greater
	///                than 0, it should scroll the drop-down list box control to the right. The higher/lower
	///                the value is, the faster the caller should scroll the drop-down list box control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the drop-down list box control upwardly; if set to a value greater than
	///                0, it should scroll the drop-down list box control downwards. The higher/lower the
	///                value is, the faster the caller should scroll the drop-down list box control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragMouseMove,
	///       DriveComboBox::Raise_ListOLEDragMouseMove, Fire_ListOLEDragEnter, Fire_ListOLEDragLeave,
	///       Fire_ListOLEDragDrop, Fire_ListMouseMove, Fire_OLEDragMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragMouseMove,
	///       DriveComboBox::Raise_ListOLEDragMouseMove, Fire_ListOLEDragEnter, Fire_ListOLEDragLeave,
	///       Fire_ListOLEDragDrop, Fire_ListMouseMove, Fire_OLEDragMouseMove
	/// \endif
	HRESULT Fire_ListOLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IDriveComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[11];
				p[10] = pData;
				p[9].plVal = reinterpret_cast<PLONG>(pEffect);									p[9].vt = VT_I4 | VT_BYREF;
				p[8].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[8].vt = VT_DISPATCH | VT_BYREF;
				p[7] = button;																									p[7].vt = VT_I2;
				p[6] = shift;																										p[6].vt = VT_I2;
				p[5] = x;																												p[5].vt = VT_XPOS_PIXELS;
				p[4] = y;																												p[4].vt = VT_YPOS_PIXELS;
				p[3] = yToItemTop;																							p[3].vt = VT_I4;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 11, 0};
				hr = pConnection->Invoke(DISPID_DCBE_LISTOLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MClick, DriveComboBox::Raise_MClick, Fire_MDblClick,
	///       Fire_Click, Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MClick, DriveComboBox::Raise_MClick, Fire_MDblClick,
	///       Fire_Click, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MDblClick, DriveComboBox::Raise_MDblClick, Fire_MClick,
	///       Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MDblClick, DriveComboBox::Raise_MDblClick, Fire_MClick,
	///       Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MouseDown, DriveComboBox::Raise_MouseDown, Fire_MouseUp,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MouseDown, DriveComboBox::Raise_MouseDown, Fire_MouseUp,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MouseEnter, DriveComboBox::Raise_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MouseEnter, DriveComboBox::Raise_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MouseHover, DriveComboBox::Raise_MouseHover,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MouseHover, DriveComboBox::Raise_MouseHover,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MouseLeave, DriveComboBox::Raise_MouseLeave,
	///       Fire_MouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MouseLeave, DriveComboBox::Raise_MouseLeave,
	///       Fire_MouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MouseMove, DriveComboBox::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MouseMove, DriveComboBox::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel
	/// \endif
	HRESULT Fire_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MouseUp, DriveComboBox::Raise_MouseUp, Fire_MouseDown,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MouseUp, DriveComboBox::Raise_MouseUp, Fire_MouseDown,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseWheel event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::MouseWheel, DriveComboBox::Raise_MouseWheel,
	///       Fire_MouseMove, Fire_ListMouseWheel
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::MouseWheel, DriveComboBox::Raise_MouseWheel,
	///       Fire_MouseMove, Fire_ListMouseWheel
	/// \endif
	HRESULT Fire_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(scrollAxis);				p[2].vt = VT_I4;
				p[1] = wheelDelta;																p[1].vt = VT_I2;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_DCBE_MOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLECompleteDrag, DriveComboBox::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLECompleteDrag, DriveComboBox::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pData;
				p[0].lVal = static_cast<LONG>(performedEffect);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLECOMPLETEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in,out] ppDropTarget The item that is the target of the drag'n'drop operation. The client
	///                may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragDrop, DriveComboBox::Raise_OLEDragDrop,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp, Fire_ListOLEDragDrop
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragDrop, DriveComboBox::Raise_OLEDragDrop,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp, Fire_ListOLEDragDrop
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IDriveComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6].plVal = reinterpret_cast<PLONG>(pEffect);									p[6].vt = VT_I4 | VT_BYREF;
				p[5].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[5].vt = VT_DISPATCH | VT_BYREF;
				p[4] = button;																									p[4].vt = VT_I2;
				p[3] = shift;																										p[3].vt = VT_I2;
				p[2] = x;																												p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																												p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);									p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoDropDown If set to \c VARIANT_TRUE, the caller should open the drop-down list box
	///                control; otherwise it should cancel automatic drop-down.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragEnter, DriveComboBox::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter,
	///       Fire_ListOLEDragEnter
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragEnter, DriveComboBox::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter,
	///       Fire_ListOLEDragEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IDriveComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoDropDown)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7].plVal = reinterpret_cast<PLONG>(pEffect);									p[7].vt = VT_I4 | VT_BYREF;
				p[6].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[6].vt = VT_DISPATCH | VT_BYREF;
				p[5] = button;																									p[5].vt = VT_I2;
				p[4] = shift;																										p[4].vt = VT_I2;
				p[3] = x;																												p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																												p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);									p[1].vt = VT_I4;
				p[0].pboolVal = pAutoDropDown;																	p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The handle of the potential drag'n'drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragEnterPotentialTarget,
	///       DriveComboBox::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragEnterPotentialTarget,
	///       DriveComboBox::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \endif
	HRESULT Fire_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndPotentialTarget;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLEDRAGENTERPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] pDropTarget The item that is the current target of the drag'n'drop operation.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragLeave, DriveComboBox::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave,
	///       Fire_ListOLEDragLeave
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragLeave, DriveComboBox::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave,
	///       Fire_ListOLEDragLeave
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, IDriveComboBoxItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pData;
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragLeavePotentialTarget,
	///       DriveComboBox::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragLeavePotentialTarget,
	///       DriveComboBox::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \endif
	HRESULT Fire_OLEDragLeavePotentialTarget(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLEDRAGLEAVEPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoDropDown If set to \c VARIANT_TRUE, the caller should open the drop-down list box
	///                control; otherwise it should cancel automatic drop-down.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEDragMouseMove, DriveComboBox::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove,
	///       Fire_ListOLEDragMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEDragMouseMove, DriveComboBox::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove,
	///       Fire_ListOLEDragMouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IDriveComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoDropDown)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7].plVal = reinterpret_cast<PLONG>(pEffect);									p[7].vt = VT_I4 | VT_BYREF;
				p[6].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[6].vt = VT_DISPATCH | VT_BYREF;
				p[5] = button;																									p[5].vt = VT_I2;
				p[4] = shift;																										p[4].vt = VT_I2;
				p[3] = x;																												p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																												p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);									p[1].vt = VT_I4;
				p[0].pboolVal = pAutoDropDown;																	p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c OLEDropEffectConstants enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEGiveFeedback, DriveComboBox::Raise_OLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEGiveFeedback, DriveComboBox::Raise_OLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \endif
	HRESULT Fire_OLEGiveFeedback(OLEDropEffectConstants effect, VARIANT_BOOL* pUseDefaultCursors)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = static_cast<LONG>(effect);			p[1].vt = VT_I4;
				p[0].pboolVal = pUseDefaultCursors;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLEGIVEFEEDBACK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c VARIANT_TRUE, the user has pressed the \c ESC key since the last
	///            time this event was fired; otherwise not.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue, cancel or complete the
	///                drag'n'drop operation. Any of the values defined by the
	///                \c OLEActionToContinueWithConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEQueryContinueDrag,
	///       DriveComboBox::Raise_OLEQueryContinueDrag, Fire_OLEGiveFeedback
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEQueryContinueDrag,
	///       DriveComboBox::Raise_OLEQueryContinueDrag, Fire_OLEGiveFeedback
	/// \endif
	HRESULT Fire_OLEQueryContinueDrag(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, OLEActionToContinueWithConstants* pActionToContinueWith)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pressedEscape;
				p[2] = button;																									p[2].vt = VT_I2;
				p[1] = shift;																										p[1].vt = VT_I2;
				p[0].plVal = reinterpret_cast<PLONG>(pActionToContinueWith);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLEQUERYCONTINUEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEReceivedNewData,
	///       DriveComboBox::Raise_OLEReceivedNewData, Fire_OLESetData
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEReceivedNewData,
	///       DriveComboBox::Raise_OLEReceivedNewData, Fire_OLESetData
	/// \endif
	HRESULT Fire_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLERECEIVEDNEWDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLESetData, DriveComboBox::Raise_OLESetData,
	///       Fire_OLEStartDrag
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLESetData, DriveComboBox::Raise_OLESetData,
	///       Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLESETDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OLEStartDrag, DriveComboBox::Raise_OLEStartDrag,
	///       Fire_OLESetData, Fire_OLECompleteDrag
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OLEStartDrag, DriveComboBox::Raise_OLEStartDrag,
	///       Fire_OLESetData, Fire_OLECompleteDrag
	/// \endif
	HRESULT Fire_OLEStartDrag(IOLEDataObject* pData)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pData;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OLESTARTDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OutOfMemory event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::OutOfMemory, DriveComboBox::Raise_OutOfMemory
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::OutOfMemory, DriveComboBox::Raise_OutOfMemory
	/// \endif
	HRESULT Fire_OutOfMemory(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DCBE_OUTOFMEMORY, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::RClick, DriveComboBox::Raise_RClick, Fire_ContextMenu,
	///       Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::RClick, DriveComboBox::Raise_RClick, Fire_ContextMenu,
	///       Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \endif
	HRESULT Fire_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::RDblClick, DriveComboBox::Raise_RDblClick, Fire_RClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::RDblClick, DriveComboBox::Raise_RDblClick, Fire_RClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::RecreatedControlWindow,
	///       DriveComboBox::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::RecreatedControlWindow,
	///       DriveComboBox::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \endif
	HRESULT Fire_RecreatedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedItem event</em>
	///
	/// \param[in] pComboItem The removed item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::RemovedItem, DriveComboBox::Raise_RemovedItem,
	///       Fire_RemovingItem, Fire_InsertedItem
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::RemovedItem, DriveComboBox::Raise_RemovedItem,
	///       Fire_RemovingItem, Fire_InsertedItem
	/// \endif
	HRESULT Fire_RemovedItem(IVirtualDriveComboBoxItem* pComboItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pComboItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DCBE_REMOVEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingItem event</em>
	///
	/// \param[in] pComboItem The item being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::RemovingItem, DriveComboBox::Raise_RemovingItem,
	///       Fire_RemovedItem, Fire_InsertingItem
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::RemovingItem, DriveComboBox::Raise_RemovingItem,
	///       Fire_RemovedItem, Fire_InsertingItem
	/// \endif
	HRESULT Fire_RemovingItem(IDriveComboBoxItem* pComboItem, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pComboItem;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DCBE_REMOVINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::ResizedControlWindow,
	///       DriveComboBox::Raise_ResizedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::ResizedControlWindow,
	///       DriveComboBox::Raise_ResizedControlWindow
	/// \endif
	HRESULT Fire_ResizedControlWindow(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DCBE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectedDriveChanged event</em>
	///
	/// \param[in] pPreviousSelectedItem The previous selected item.
	/// \param[in] pNewSelectedItem The new selected item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::SelectedDriveChanged,
	///       DriveComboBox::Raise_SelectedDriveChanged, Fire_SelectedDriveChanged, Fire_SelectionChanged
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::SelectedDriveChanged,
	///       DriveComboBox::Raise_SelectedDriveChanged, Fire_SelectedDriveChanged, Fire_SelectionChanged
	/// \endif
	HRESULT Fire_SelectedDriveChanged(IDriveComboBoxItem* pPreviousSelectedItem, IDriveComboBoxItem* pNewSelectedItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pPreviousSelectedItem;
				p[0] = pNewSelectedItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DCBE_SELECTEDDRIVECHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectionCanceled event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::SelectionCanceled, DriveComboBox::Raise_SelectionCanceled,
	///       Fire_SelectionChanging, Fire_SelectionChanged
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::SelectionCanceled, DriveComboBox::Raise_SelectionCanceled,
	///       Fire_SelectionChanging, Fire_SelectionChanged
	/// \endif
	HRESULT Fire_SelectionCanceled(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DCBE_SELECTIONCANCELED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectionChanged event</em>
	///
	/// \param[in] pPreviousSelectedItem The previous selected item.
	/// \param[in] pNewSelectedItem The new selected item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::SelectionChanged, DriveComboBox::Raise_SelectionChanged,
	///       Fire_SelectedDriveChanged, Fire_SelectionCanceled, Fire_SelectionChanging
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::SelectionChanged, DriveComboBox::Raise_SelectionChanged,
	///       Fire_SelectedDriveChanged, Fire_SelectionCanceled, Fire_SelectionChanging
	/// \endif
	HRESULT Fire_SelectionChanged(IDriveComboBoxItem* pPreviousSelectedItem, IDriveComboBoxItem* pNewSelectedItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pPreviousSelectedItem;
				p[0] = pNewSelectedItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DCBE_SELECTIONCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectionChanging event</em>
	///
	/// \param[in] pNewSelectedItem The item that will become the selected item. May be \c NULL.
	/// \param[in] selectionFieldText The text currently being displayed in the selection field.
	/// \param[in] selectionFieldHasBeenEdited Specifies whether the text displayed in the control's edit
	///            box has been edited. If \c VARIANT_TRUE, the text has been edited; otherwise not.
	/// \param[in] selectionChangeReason Specifies the action that led to this event being raised. Any
	///            of the values defined by the \c SelectionChangeReasonConstants enumeration is valid.
	/// \param[in,out] cancelChange If set to \c VARIANT_TRUE, the caller should abort the selection change,
	///                i. e. the currently selected item should remain selected. If set to \c VARIANT_FALSE,
	///                the selection should be changed to the item specified by \c pNewSelectedItem.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::SelectionChanging, DriveComboBox::Raise_SelectionChanging,
	///       Fire_BeginSelectionChange, Fire_SelectionChanged
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::SelectionChanging, DriveComboBox::Raise_SelectionChanging,
	///       Fire_BeginSelectionChange, Fire_SelectionChanged
	/// \endif
	HRESULT Fire_SelectionChanging(IDriveComboBoxItem* pNewSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pNewSelectedItem;
				p[3] = selectionFieldText;
				p[2] = selectionFieldHasBeenEdited;
				p[1].lVal = static_cast<LONG>(selectionChangeReason);		p[1].vt = VT_I4;
				p[0].pboolVal = pCancelChange;													p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_SELECTIONCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be a
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::XClick, DriveComboBox::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::XClick, DriveComboBox::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick
	/// \endif
	HRESULT Fire_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IDriveComboBoxEvents::XDblClick, DriveComboBox::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IDriveComboBoxEvents::XDblClick, DriveComboBox::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \endif
	HRESULT Fire_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DCBE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IDriveComboBoxEvents