//////////////////////////////////////////////////////////////////////
/// \class Proxy_IComboBoxEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IComboBoxEvents interface</em>
///
/// \if UNICODE
///   \sa ComboBox, CBLCtlsLibU::_IComboBoxEvents
/// \else
///   \sa ComboBox, CBLCtlsLibA::_IComboBoxEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IComboBoxEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IComboBoxEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c BeforeDrawText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::BeforeDrawText, ComboBox::Raise_BeforeDrawText
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::BeforeDrawText, ComboBox::Raise_BeforeDrawText
	/// \endif
	HRESULT Fire_BeforeDrawText(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_CBE_BEFOREDRAWTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::Click, ComboBox::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::Click, ComboBox::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
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
				hr = pConnection->Invoke(DISPID_CBE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CompareItems event</em>
	///
	/// \param[in] pFirstItem The first item to compare.
	/// \param[in] pSecondItem The second item to compare.
	/// \param[in] locale The identifier of the locale to use for comparison.
	/// \param[out] pResult Receives one of the values defined by the \c CompareResultConstants enumeration.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::CompareItems, ComboBox::Raise_CompareItems
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::CompareItems, ComboBox::Raise_CompareItems
	/// \endif
	HRESULT Fire_CompareItems(IComboBoxItem* pFirstItem, IComboBoxItem* pSecondItem, LONG locale, CompareResultConstants* pResult)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pFirstItem;
				p[2] = pSecondItem;
				p[1] = locale;
				p[0].plVal = reinterpret_cast<PLONG>(pResult);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_CBE_COMPAREITEMS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ContextMenu, ComboBox::Raise_ContextMenu, Fire_RClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ContextMenu, ComboBox::Raise_ContextMenu, Fire_RClick
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
				hr = pConnection->Invoke(DISPID_CBE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::CreatedEditControlWindow,
	///       ComboBox::Raise_CreatedEditControlWindow, Fire_DestroyedEditControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::CreatedEditControlWindow,
	///       ComboBox::Raise_CreatedEditControlWindow, Fire_DestroyedEditControlWindow
	/// \endif
	HRESULT Fire_CreatedEditControlWindow(LONG hWndEdit)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndEdit;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_CBE_CREATEDEDITCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::CreatedListBoxControlWindow,
	///       ComboBox::Raise_CreatedListBoxControlWindow, Fire_DestroyedListBoxControlWindow,
	///       Fire_ListDropDown
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::CreatedListBoxControlWindow,
	///       ComboBox::Raise_CreatedListBoxControlWindow, Fire_DestroyedListBoxControlWindow,
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
				hr = pConnection->Invoke(DISPID_CBE_CREATEDLISTBOXCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::DblClick, ComboBox::Raise_DblClick, Fire_Click, Fire_MDblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::DblClick, ComboBox::Raise_DblClick, Fire_Click, Fire_MDblClick,
	///       Fire_RDblClick, Fire_XDblClick
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
				hr = pConnection->Invoke(DISPID_CBE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::DestroyedControlWindow, ComboBox::Raise_DestroyedControlWindow,
	///       Fire_RecreatedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::DestroyedControlWindow, ComboBox::Raise_DestroyedControlWindow,
	///       Fire_RecreatedControlWindow
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
				hr = pConnection->Invoke(DISPID_CBE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::DestroyedEditControlWindow,
	///       ComboBox::Raise_DestroyedEditControlWindow, Fire_CreatedEditControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::DestroyedEditControlWindow,
	///       ComboBox::Raise_DestroyedEditControlWindow, Fire_CreatedEditControlWindow
	/// \endif
	HRESULT Fire_DestroyedEditControlWindow(LONG hWndEdit)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndEdit;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_CBE_DESTROYEDEDITCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::DestroyedListBoxControlWindow,
	///       ComboBox::Raise_DestroyedListBoxControlWindow, Fire_CreatedListBoxControlWindow,
	///       Fire_ListCloseUp
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::DestroyedListBoxControlWindow,
	///       ComboBox::Raise_DestroyedListBoxControlWindow, Fire_CreatedListBoxControlWindow,
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
				hr = pConnection->Invoke(DISPID_CBE_DESTROYEDLISTBOXCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::FreeItemData, ComboBox::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::FreeItemData, ComboBox::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \endif
	HRESULT Fire_FreeItemData(IComboBoxItem* pComboItem)
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
				hr = pConnection->Invoke(DISPID_CBE_FREEITEMDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::InsertedItem, ComboBox::Raise_InsertedItem, Fire_InsertingItem,
	///       Fire_RemovedItem
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::InsertedItem, ComboBox::Raise_InsertedItem, Fire_InsertingItem,
	///       Fire_RemovedItem
	/// \endif
	HRESULT Fire_InsertedItem(IComboBoxItem* pComboItem)
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
				hr = pConnection->Invoke(DISPID_CBE_INSERTEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::InsertingItem, ComboBox::Raise_InsertingItem, Fire_InsertedItem,
	///       Fire_RemovingItem
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::InsertingItem, ComboBox::Raise_InsertingItem, Fire_InsertedItem,
	///       Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertingItem(IVirtualComboBoxItem* pComboItem, VARIANT_BOOL* pCancelInsertion)
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
				hr = pConnection->Invoke(DISPID_CBE_INSERTINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ItemMouseEnter, ComboBox::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_ListMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ItemMouseEnter, ComboBox::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_ListMouseMove
	/// \endif
	HRESULT Fire_ItemMouseEnter(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_ITEMMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ItemMouseLeave, ComboBox::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_ListMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ItemMouseLeave, ComboBox::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_ListMouseMove
	/// \endif
	HRESULT Fire_ItemMouseLeave(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_ITEMMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::KeyDown, ComboBox::Raise_KeyDown, Fire_KeyUp, Fire_KeyPress
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::KeyDown, ComboBox::Raise_KeyDown, Fire_KeyUp, Fire_KeyPress
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
				hr = pConnection->Invoke(DISPID_CBE_KEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::KeyPress, ComboBox::Raise_KeyPress, Fire_KeyDown, Fire_KeyUp
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::KeyPress, ComboBox::Raise_KeyPress, Fire_KeyDown, Fire_KeyUp
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
				hr = pConnection->Invoke(DISPID_CBE_KEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::KeyUp, ComboBox::Raise_KeyUp, Fire_KeyDown, Fire_KeyPress
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::KeyUp, ComboBox::Raise_KeyUp, Fire_KeyDown, Fire_KeyPress
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
				hr = pConnection->Invoke(DISPID_CBE_KEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListCloseUp event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ListCloseUp, ComboBox::Raise_ListCloseUp
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ListCloseUp, ComboBox::Raise_ListCloseUp
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
				hr = pConnection->Invoke(DISPID_CBE_LISTCLOSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ListDropDown event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ListDropDown, ComboBox::Raise_ListDropDown
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ListDropDown, ComboBox::Raise_ListDropDown
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
				hr = pConnection->Invoke(DISPID_CBE_LISTDROPDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ListMouseDown, ComboBox::Raise_ListMouseDown, Fire_ListMouseUp,
	///       Fire_MouseDown
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ListMouseDown, ComboBox::Raise_ListMouseDown, Fire_ListMouseUp,
	///       Fire_MouseDown
	/// \endif
	HRESULT Fire_ListMouseDown(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_LISTMOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ListMouseMove, ComboBox::Raise_ListMouseMove,
	///       Fire_ListMouseDown, Fire_ListMouseUp, Fire_ListMouseWheel, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ListMouseMove, ComboBox::Raise_ListMouseMove,
	///       Fire_ListMouseDown, Fire_ListMouseUp, Fire_ListMouseWheel, Fire_MouseMove
	/// \endif
	HRESULT Fire_ListMouseMove(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_LISTMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ListMouseUp, ComboBox::Raise_ListMouseUp, Fire_ListMouseDown,
	///       Fire_MouseUp
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ListMouseUp, ComboBox::Raise_ListMouseUp, Fire_ListMouseDown,
	///       Fire_MouseUp
	/// \endif
	HRESULT Fire_ListMouseUp(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_LISTMOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ListMouseWheel, ComboBox::Raise_ListMouseWheel,
	///       Fire_ListMouseMove, Fire_MouseWheel
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ListMouseWheel, ComboBox::Raise_ListMouseWheel,
	///       Fire_ListMouseMove, Fire_MouseWheel
	/// \endif
	HRESULT Fire_ListMouseWheel(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_LISTMOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OLEDragDrop, ComboBox::Raise_ListOLEDragDrop,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragMouseMove, Fire_ListOLEDragLeave, Fire_ListMouseUp,
	///       Fire_OLEDragDrop
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OLEDragDrop, ComboBox::Raise_ListOLEDragDrop,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragMouseMove, Fire_ListOLEDragLeave, Fire_ListMouseUp,
	///       Fire_OLEDragDrop
	/// \endif
	HRESULT Fire_ListOLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_LISTOLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OLEDragEnter, ComboBox::Raise_ListOLEDragEnter,
	///       Fire_ListOLEDragMouseMove, Fire_ListOLEDragLeave, Fire_ListOLEDragDrop, Fire_OLEDragEnter
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OLEDragEnter, ComboBox::Raise_ListOLEDragEnter,
	///       Fire_ListOLEDragMouseMove, Fire_ListOLEDragLeave, Fire_ListOLEDragDrop, Fire_OLEDragEnter
	/// \endif
	HRESULT Fire_ListOLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
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
				hr = pConnection->Invoke(DISPID_CBE_LISTOLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OLEDragLeave, ComboBox::Raise_ListOLEDragLeave,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragMouseMove, Fire_ListOLEDragDrop, Fire_OLEDragLeave
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OLEDragLeave, ComboBox::Raise_ListOLEDragLeave,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragMouseMove, Fire_ListOLEDragDrop, Fire_OLEDragLeave
	/// \endif
	HRESULT Fire_ListOLEDragLeave(IOLEDataObject* pData, IComboBoxItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoCloseUp)
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
				hr = pConnection->Invoke(DISPID_CBE_LISTOLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OLEDragMouseMove, ComboBox::Raise_ListOLEDragMouseMove,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragLeave, Fire_ListOLEDragDrop, Fire_ListMouseMove,
	///       Fire_OLEDragMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OLEDragMouseMove, ComboBox::Raise_ListOLEDragMouseMove,
	///       Fire_ListOLEDragEnter, Fire_ListOLEDragLeave, Fire_ListOLEDragDrop, Fire_ListMouseMove,
	///       Fire_OLEDragMouseMove
	/// \endif
	HRESULT Fire_ListOLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
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
				hr = pConnection->Invoke(DISPID_CBE_LISTOLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MClick, ComboBox::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MClick, ComboBox::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
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
				hr = pConnection->Invoke(DISPID_CBE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MDblClick, ComboBox::Raise_MDblClick, Fire_MClick,
	///       Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MDblClick, ComboBox::Raise_MDblClick, Fire_MClick,
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
				hr = pConnection->Invoke(DISPID_CBE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MeasureItem event</em>
	///
	/// \param[in] pComboItem The item for which the size is required. If the event is raised for the
	///            selection field, this parameter will be \c NULL.
	/// \param[out] pItemHeight Must be set to the item's height in pixels by the client app.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MeasureItem, ComboBox::Raise_MeasureItem, Fire_OwnerDrawItem
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MeasureItem, ComboBox::Raise_MeasureItem, Fire_OwnerDrawItem
	/// \endif
	HRESULT Fire_MeasureItem(IComboBoxItem* pComboItem, OLE_YSIZE_PIXELS* pItemHeight)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pComboItem;
				p[0].plVal = pItemHeight;		p[0].vt = VT_YSIZE_PIXELS | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_CBE_MEASUREITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MouseDown, ComboBox::Raise_MouseDown, Fire_MouseUp, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_ListMouseDown, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MouseDown, ComboBox::Raise_MouseDown, Fire_MouseUp, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_ListMouseDown, Fire_XClick
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
				hr = pConnection->Invoke(DISPID_CBE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MouseEnter, ComboBox::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MouseEnter, ComboBox::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_MouseHover, Fire_MouseMove
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
				hr = pConnection->Invoke(DISPID_CBE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MouseHover, ComboBox::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MouseHover, ComboBox::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
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
				hr = pConnection->Invoke(DISPID_CBE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MouseLeave, ComboBox::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MouseLeave, ComboBox::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_MouseHover, Fire_MouseMove
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
				hr = pConnection->Invoke(DISPID_CBE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MouseMove, ComboBox::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel, Fire_ListMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MouseMove, ComboBox::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel, Fire_ListMouseMove
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
				hr = pConnection->Invoke(DISPID_CBE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MouseUp, ComboBox::Raise_MouseUp, Fire_MouseDown, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick, Fire_ListMouseUp
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MouseUp, ComboBox::Raise_MouseUp, Fire_MouseDown, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick, Fire_ListMouseUp
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
				hr = pConnection->Invoke(DISPID_CBE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::MouseWheel, ComboBox::Raise_MouseWheel, Fire_MouseMove,
	///       Fire_ListMouseWheel
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::MouseWheel, ComboBox::Raise_MouseWheel, Fire_MouseMove,
	///       Fire_ListMouseWheel
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
				hr = pConnection->Invoke(DISPID_CBE_MOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OLEDragDrop, ComboBox::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp, Fire_ListOLEDragDrop
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OLEDragDrop, ComboBox::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp, Fire_ListOLEDragDrop
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OLEDragEnter, ComboBox::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter,
	///       Fire_ListOLEDragEnter
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OLEDragEnter, ComboBox::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter,
	///       Fire_ListOLEDragEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoDropDown)
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
				hr = pConnection->Invoke(DISPID_CBE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OLEDragLeave, ComboBox::Raise_OLEDragLeave, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave, Fire_ListOLEDragLeave
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OLEDragLeave, ComboBox::Raise_OLEDragLeave, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave, Fire_ListOLEDragLeave
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, IComboBoxItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_CBE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OLEDragMouseMove, ComboBox::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove,
	///       Fire_ListOLEDragMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OLEDragMouseMove, ComboBox::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove,
	///       Fire_ListOLEDragMouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IComboBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoDropDown)
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
				hr = pConnection->Invoke(DISPID_CBE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OutOfMemory event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OutOfMemory, ComboBox::Raise_OutOfMemory
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OutOfMemory, ComboBox::Raise_OutOfMemory
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
				hr = pConnection->Invoke(DISPID_CBE_OUTOFMEMORY, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OwnerDrawItem event</em>
	///
	/// \param[in] pComboItem The item to draw. If the event is raised for the selection field, this
	///            parameter will be \c NULL.
	/// \param[in] requiredAction Specifies the required drawing action. Any combination of the values
	///            defined by the \c OwnerDrawActionConstants enumeration is valid.
	/// \param[in] itemState The item's current state (focused, selected etc.). Most of the values
	///            defined by the \c OwnerDrawItemStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::OwnerDrawItem, ComboBox::Raise_OwnerDrawItem,
	///       Fire_MeasureItem
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::OwnerDrawItem, ComboBox::Raise_OwnerDrawItem,
	///       Fire_MeasureItem
	/// \endif
	HRESULT Fire_OwnerDrawItem(IComboBoxItem* pComboItem, OwnerDrawActionConstants requiredAction, OwnerDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pComboItem;
				p[3].lVal = static_cast<LONG>(requiredAction);		p[3].vt = VT_I4;
				p[2].lVal = static_cast<LONG>(itemState);					p[2].vt = VT_I4;
				p[1] = hDC;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					CLSIDFromString(OLESTR("{8940470A-8420-402b-8949-7D715B9C11CD}"), &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_CBLCtlsLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					CLSIDFromString(OLESTR("{8B9E68EE-76A6-43d6-AF06-0535679ECE84}"), &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_CBLCtlsLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[0]);
				p[0].vt = VT_RECORD | VT_BYREF;
				p[0].pRecInfo = pRecordInfo;
				p[0].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Top = pDrawingRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_CBE_OWNERDRAWITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[0].pvRecord);
				}
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::RClick, ComboBox::Raise_RClick, Fire_ContextMenu,
	///       Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::RClick, ComboBox::Raise_RClick, Fire_ContextMenu,
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
				hr = pConnection->Invoke(DISPID_CBE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::RDblClick, ComboBox::Raise_RDblClick, Fire_RClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::RDblClick, ComboBox::Raise_RDblClick, Fire_RClick,
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
				hr = pConnection->Invoke(DISPID_CBE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::RecreatedControlWindow, ComboBox::Raise_RecreatedControlWindow,
	///       Fire_DestroyedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::RecreatedControlWindow, ComboBox::Raise_RecreatedControlWindow,
	///       Fire_DestroyedControlWindow
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
				hr = pConnection->Invoke(DISPID_CBE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::RemovedItem, ComboBox::Raise_RemovedItem, Fire_RemovingItem,
	///       Fire_InsertedItem
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::RemovedItem, ComboBox::Raise_RemovedItem, Fire_RemovingItem,
	///       Fire_InsertedItem
	/// \endif
	HRESULT Fire_RemovedItem(IVirtualComboBoxItem* pComboItem)
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
				hr = pConnection->Invoke(DISPID_CBE_REMOVEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::RemovingItem, ComboBox::Raise_RemovingItem, Fire_RemovedItem,
	///       Fire_InsertingItem
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::RemovingItem, ComboBox::Raise_RemovingItem, Fire_RemovedItem,
	///       Fire_InsertingItem
	/// \endif
	HRESULT Fire_RemovingItem(IComboBoxItem* pComboItem, VARIANT_BOOL* pCancelDeletion)
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
				hr = pConnection->Invoke(DISPID_CBE_REMOVINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::ResizedControlWindow, ComboBox::Raise_ResizedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::ResizedControlWindow, ComboBox::Raise_ResizedControlWindow
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
				hr = pConnection->Invoke(DISPID_CBE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectionCanceled event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::SelectionCanceled, ComboBox::Raise_SelectionCanceled,
	///       Fire_SelectionChanging, Fire_SelectionChanged
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::SelectionCanceled, ComboBox::Raise_SelectionCanceled,
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
				hr = pConnection->Invoke(DISPID_CBE_SELECTIONCANCELED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::SelectionChanged, ComboBox::Raise_SelectionChanged,
	///       Fire_SelectionCanceled, Fire_SelectionChanging
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::SelectionChanged, ComboBox::Raise_SelectionChanged,
	///       Fire_SelectionCanceled, Fire_SelectionChanging
	/// \endif
	HRESULT Fire_SelectionChanged(IComboBoxItem* pPreviousSelectedItem, IComboBoxItem* pNewSelectedItem)
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
				hr = pConnection->Invoke(DISPID_CBE_SELECTIONCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectionChanging event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::SelectionChanging, ComboBox::Raise_SelectionChanging,
	///       Fire_SelectionChanged, Fire_SelectionCanceled
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::SelectionChanging, ComboBox::Raise_SelectionChanging,
	///       Fire_SelectionChanged, Fire_SelectionCanceled
	/// \endif
	HRESULT Fire_SelectionChanging(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_CBE_SELECTIONCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TextChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::TextChanged, ComboBox::Raise_TextChanged
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::TextChanged, ComboBox::Raise_TextChanged
	/// \endif
	HRESULT Fire_TextChanged(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_CBE_TEXTCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TruncatedText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::TruncatedText, ComboBox::Raise_TruncatedText
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::TruncatedText, ComboBox::Raise_TruncatedText
	/// \endif
	HRESULT Fire_TruncatedText(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_CBE_TRUNCATEDTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c WritingDirectionChanged event</em>
	///
	/// \param[in] newWritingDirection The control's new writing direction. Any of the values defined by the
	///            \c WritingDirectionConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IComboBoxEvents::WritingDirectionChanged, ComboBox::Raise_WritingDirectionChanged
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::WritingDirectionChanged, ComboBox::Raise_WritingDirectionChanged
	/// \endif
	HRESULT Fire_WritingDirectionChanged(WritingDirectionConstants newWritingDirection)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].lVal = static_cast<LONG>(newWritingDirection);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_CBE_WRITINGDIRECTIONCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::XClick, ComboBox::Raise_XClick, Fire_XDblClick, Fire_Click,
	///       Fire_MClick, Fire_RClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::XClick, ComboBox::Raise_XClick, Fire_XDblClick, Fire_Click,
	///       Fire_MClick, Fire_RClick
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
				hr = pConnection->Invoke(DISPID_CBE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IComboBoxEvents::XDblClick, ComboBox::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IComboBoxEvents::XDblClick, ComboBox::Raise_XDblClick, Fire_XClick,
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
				hr = pConnection->Invoke(DISPID_CBE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IComboBoxEvents