//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListBoxEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListBoxEvents interface</em>
///
/// \if UNICODE
///   \sa ListBox, CBLCtlsLibU::_IListBoxEvents
/// \else
///   \sa ListBox, CBLCtlsLibA::_IListBoxEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListBoxEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListBoxEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c AbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::AbortedDrag, ListBox::Raise_AbortedDrag, Fire_Drop
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::AbortedDrag, ListBox::Raise_AbortedDrag, Fire_Drop
	/// \endif
	HRESULT Fire_AbortedDrag(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_LBE_ABORTEDDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CaretChanged event</em>
	///
	/// \param[in] pPreviousCaretItem The previous caret item.
	/// \param[in] pNewCaretItem The new caret item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::CaretChanged, ListBox::Raise_CaretChanged,
	///       Fire_SelectionChanged
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::CaretChanged, ListBox::Raise_CaretChanged,
	///       Fire_SelectionChanged
	/// \endif
	HRESULT Fire_CaretChanged(IListBoxItem* pPreviousCaretItem, IListBoxItem* pNewCaretItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pPreviousCaretItem;
				p[0] = pNewCaretItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_LBE_CARETCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Click event</em>
	///
	/// \param[in] pListItem The clicked item. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::Click, ListBox::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::Click, ListBox::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_Click(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::CompareItems, ListBox::Raise_CompareItems
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::CompareItems, ListBox::Raise_CompareItems
	/// \endif
	HRESULT Fire_CompareItems(IListBoxItem* pFirstItem, IListBoxItem* pSecondItem, LONG locale, CompareResultConstants* pResult)
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
				hr = pConnection->Invoke(DISPID_LBE_COMPAREITEMS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ContextMenu event</em>
	///
	/// \param[in] pListItem The item that the context menu refers to. May be \c NULL.
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::ContextMenu, ListBox::Raise_ContextMenu, Fire_RClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ContextMenu, ListBox::Raise_ContextMenu, Fire_RClick
	/// \endif
	HRESULT Fire_ContextMenu(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DblClick event</em>
	///
	/// \param[in] pListItem The double-clicked item. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::DblClick, ListBox::Raise_DblClick, Fire_Click, Fire_MDblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::DblClick, ListBox::Raise_DblClick, Fire_Click, Fire_MDblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_DblClick(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::DestroyedControlWindow, ListBox::Raise_DestroyedControlWindow,
	///       Fire_RecreatedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::DestroyedControlWindow, ListBox::Raise_DestroyedControlWindow,
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
				hr = pConnection->Invoke(DISPID_LBE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DragMouseMove event</em>
	///
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
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::DragMouseMove, ListBox::Raise_DragMouseMove, Fire_MouseMove,
	///       Fire_OLEDragMouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::DragMouseMove, ListBox::Raise_DragMouseMove, Fire_MouseMove,
	///       Fire_OLEDragMouseMove
	/// \endif
	HRESULT Fire_DragMouseMove(IListBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
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
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_LBE_DRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Drop event</em>
	///
	/// \param[in] pDropTarget The item object that is the nearest one from the mouse cursor's position.
	///            If the mouse cursor's position lies outside the control's client area, this parameter
	///            should be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c pDropTarget item's upper border.
	/// \param[in] hitTestDetails The part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::Drop, ListBox::Raise_Drop, Fire_AbortedDrag
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::Drop, ListBox::Raise_Drop, Fire_AbortedDrag
	/// \endif
	HRESULT Fire_Drop(IListBoxItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pDropTarget;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1] = yToItemTop;																p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_LBE_DROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeItemData event</em>
	///
	/// \param[in] pListItem The item whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::FreeItemData, ListBox::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::FreeItemData, ListBox::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \endif
	HRESULT Fire_FreeItemData(IListBoxItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_LBE_FREEITEMDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedItem event</em>
	///
	/// \param[in] pListItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::InsertedItem, ListBox::Raise_InsertedItem, Fire_InsertingItem,
	///       Fire_RemovedItem
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::InsertedItem, ListBox::Raise_InsertedItem, Fire_InsertingItem,
	///       Fire_RemovedItem
	/// \endif
	HRESULT Fire_InsertedItem(IListBoxItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_LBE_INSERTEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingItem event</em>
	///
	/// \param[in] pListItem The item being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::InsertingItem, ListBox::Raise_InsertingItem, Fire_InsertedItem,
	///       Fire_RemovingItem
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::InsertingItem, ListBox::Raise_InsertingItem, Fire_InsertedItem,
	///       Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertingItem(IVirtualListBoxItem* pListItem, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListItem;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_LBE_INSERTINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemBeginDrag event</em>
	///
	/// \param[in] pListItem The item that the user wants to drag.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::ItemBeginDrag, ListBox::Raise_ItemBeginDrag, Fire_ItemBeginRDrag
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ItemBeginDrag, ListBox::Raise_ItemBeginDrag, Fire_ItemBeginRDrag
	/// \endif
	HRESULT Fire_ItemBeginDrag(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_ITEMBEGINDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemBeginRDrag event</em>
	///
	/// \param[in] pListItem The item that the user wants to drag.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::ItemBeginRDrag, ListBox::Raise_ItemBeginRDrag, Fire_ItemBeginDrag
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ItemBeginRDrag, ListBox::Raise_ItemBeginRDrag, Fire_ItemBeginDrag
	/// \endif
	HRESULT Fire_ItemBeginRDrag(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_ITEMBEGINRDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetDisplayInfo event</em>
	///
	/// \param[in] pListItem The item whose properties are requested.
	/// \param[in] requestedInfo Specifies which properties' values are required. Some combinations of
	///            the values defined by the \c RequestedInfoConstants enumeration are valid.
	/// \param[out] pIconIndex The zero-based index of the item's icon. If the \c requestedInfo
	///             parameter doesn't include \c riIconIndex, the caller should ignore this value.
	/// \param[out] pOverlayIndex The zero-based index of the item's overlay icon. If the \c requestedInfo
	///             parameter doesn't include \c riOverlayIndex, the caller should ignore this value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::ItemGetDisplayInfo, ListBox::Raise_ItemGetDisplayInfo
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ItemGetDisplayInfo, ListBox::Raise_ItemGetDisplayInfo
	/// \endif
	HRESULT Fire_ItemGetDisplayInfo(IListBoxItem* pListItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pOverlayIndex)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pListItem;
				p[2] = static_cast<LONG>(requestedInfo);		p[2].vt = VT_I4;
				p[1].plVal = pIconIndex;										p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pOverlayIndex;									p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_LBE_ITEMGETDISPLAYINFO, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetInfoTipText event</em>
	///
	/// \param[in] pListItem The item whose info tip text is requested.
	/// \param[in] maxInfoTipLength The maximum number of characters the info tip text may consist of.
	/// \param[out] pInfoTipText The item's info tip text.
	/// \param[in,out] pAbortToolTip If \c VARIANT_TRUE, the caller should NOT show the tooltip;
	///                otherwise it should.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::ItemGetInfoTipText, ListBox::Raise_ItemGetInfoTipText
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ItemGetInfoTipText, ListBox::Raise_ItemGetInfoTipText
	/// \endif
	HRESULT Fire_ItemGetInfoTipText(IListBoxItem* pListItem, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pListItem;
				p[2] = maxInfoTipLength;
				p[1].pbstrVal = pInfoTipText;			p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pAbortToolTip;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_LBE_ITEMGETINFOTIPTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemMouseEnter event</em>
	///
	/// \param[in] pListItem The item that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::ItemMouseEnter, ListBox::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ItemMouseEnter, ListBox::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_ItemMouseEnter(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_ITEMMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemMouseLeave event</em>
	///
	/// \param[in] pListItem The item that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::ItemMouseLeave, ListBox::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ItemMouseLeave, ListBox::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_MouseMove
	/// \endif
	HRESULT Fire_ItemMouseLeave(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_ITEMMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::KeyDown, ListBox::Raise_KeyDown, Fire_KeyUp, Fire_KeyPress,
	///       Fire_ProcessKeyStroke
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::KeyDown, ListBox::Raise_KeyDown, Fire_KeyUp, Fire_KeyPress,
	///       Fire_ProcessKeyStroke
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
				hr = pConnection->Invoke(DISPID_LBE_KEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::KeyPress, ListBox::Raise_KeyPress, Fire_KeyDown, Fire_KeyUp,
	///       Fire_ProcessKeyStroke
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::KeyPress, ListBox::Raise_KeyPress, Fire_KeyDown, Fire_KeyUp,
	///       Fire_ProcessKeyStroke
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
				hr = pConnection->Invoke(DISPID_LBE_KEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::KeyUp, ListBox::Raise_KeyUp, Fire_KeyDown, Fire_KeyPress,
	///       Fire_ProcessKeyStroke
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::KeyUp, ListBox::Raise_KeyUp, Fire_KeyDown, Fire_KeyPress,
	///       Fire_ProcessKeyStroke
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
				hr = pConnection->Invoke(DISPID_LBE_KEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MClick event</em>
	///
	/// \param[in] pListItem The clicked item. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MClick, ListBox::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MClick, ListBox::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MClick(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MDblClick event</em>
	///
	/// \param[in] pListItem The double-clicked item. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MDblClick, ListBox::Raise_MDblClick, Fire_MClick, Fire_DblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MDblClick, ListBox::Raise_MDblClick, Fire_MClick, Fire_DblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_MDblClick(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MeasureItem event</em>
	///
	/// \param[in] pListItem The item for which the size is required. If the list box is empty, this
	///            parameter will be \c NULL.
	/// \param[out] pItemHeight Must be set to the item's height in pixels by the client app.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::MeasureItem, ListBox::Raise_MeasureItem, Fire_OwnerDrawItem
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MeasureItem, ListBox::Raise_MeasureItem, Fire_OwnerDrawItem
	/// \endif
	HRESULT Fire_MeasureItem(IListBoxItem* pListItem, OLE_YSIZE_PIXELS* pItemHeight)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListItem;
				p[0].plVal = pItemHeight;		p[0].vt = VT_YSIZE_PIXELS | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_LBE_MEASUREITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseDown event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MouseDown, ListBox::Raise_MouseDown, Fire_MouseUp, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MouseDown, ListBox::Raise_MouseDown, Fire_MouseUp, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseDown(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseEnter event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MouseEnter, ListBox::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_ItemMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MouseEnter, ListBox::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_ItemMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseEnter(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseHover event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MouseHover, ListBox::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MouseHover, ListBox::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseHover(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseLeave event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MouseLeave, ListBox::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_ItemMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MouseLeave, ListBox::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_ItemMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseLeave(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseMove event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MouseMove, ListBox::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MouseMove, ListBox::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel
	/// \endif
	HRESULT Fire_MouseMove(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseUp event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MouseUp, ListBox::Raise_MouseUp, Fire_MouseDown, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MouseUp, ListBox::Raise_MouseUp, Fire_MouseDown, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseUp(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseWheel event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::MouseWheel, ListBox::Raise_MouseWheel, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::MouseWheel, ListBox::Raise_MouseWheel, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseWheel(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pListItem;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(scrollAxis);				p[2].vt = VT_I4;
				p[1] = wheelDelta;																p[1].vt = VT_I2;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_LBE_MOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLECompleteDrag, ListBox::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLECompleteDrag, ListBox::Raise_OLECompleteDrag,
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
				hr = pConnection->Invoke(DISPID_LBE_OLECOMPLETEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEDragDrop, ListBox::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEDragDrop, ListBox::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_LBE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEDragEnter, ListBox::Raise_OLEDragEnter, Fire_OLEDragMouseMove,
	///       Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEDragEnter, ListBox::Raise_OLEDragEnter, Fire_OLEDragMouseMove,
	///       Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
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
				hr = pConnection->Invoke(DISPID_LBE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEDragEnterPotentialTarget,
	///       ListBox::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEDragEnterPotentialTarget,
	///       ListBox::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
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
				hr = pConnection->Invoke(DISPID_LBE_OLEDRAGENTERPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEDragLeave, ListBox::Raise_OLEDragLeave, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEDragLeave, ListBox::Raise_OLEDragLeave, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, IListBoxItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6] = pDropTarget;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1] = yToItemTop;																p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_LBE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEDragLeavePotentialTarget,
	///       ListBox::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEDragLeavePotentialTarget,
	///       ListBox::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
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
				hr = pConnection->Invoke(DISPID_LBE_OLEDRAGLEAVEPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	/// \param[in] yToItemTop The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            \c ppDropTarget item's upper border.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEDragMouseMove, ListBox::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEDragMouseMove, ListBox::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListBoxItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
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
				hr = pConnection->Invoke(DISPID_LBE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEGiveFeedback, ListBox::Raise_OLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEGiveFeedback, ListBox::Raise_OLEGiveFeedback,
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
				hr = pConnection->Invoke(DISPID_LBE_OLEGIVEFEEDBACK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEQueryContinueDrag, ListBox::Raise_OLEQueryContinueDrag,
	///       Fire_OLEGiveFeedback
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEQueryContinueDrag, ListBox::Raise_OLEQueryContinueDrag,
	///       Fire_OLEGiveFeedback
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
				hr = pConnection->Invoke(DISPID_LBE_OLEQUERYCONTINUEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEReceivedNewData, ListBox::Raise_OLEReceivedNewData,
	///       Fire_OLESetData
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEReceivedNewData, ListBox::Raise_OLEReceivedNewData,
	///       Fire_OLESetData
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
				hr = pConnection->Invoke(DISPID_LBE_OLERECEIVEDNEWDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLESetData, ListBox::Raise_OLESetData, Fire_OLEStartDrag
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLESetData, ListBox::Raise_OLESetData, Fire_OLEStartDrag
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
				hr = pConnection->Invoke(DISPID_LBE_OLESETDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::OLEStartDrag, ListBox::Raise_OLEStartDrag, Fire_OLESetData,
	///       Fire_OLECompleteDrag
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OLEStartDrag, ListBox::Raise_OLEStartDrag, Fire_OLESetData,
	///       Fire_OLECompleteDrag
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
				hr = pConnection->Invoke(DISPID_LBE_OLESTARTDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OutOfMemory event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::OutOfMemory, ListBox::Raise_OutOfMemory
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OutOfMemory, ListBox::Raise_OutOfMemory
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
				hr = pConnection->Invoke(DISPID_LBE_OUTOFMEMORY, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OwnerDrawItem event</em>
	///
	/// \param[in] pListItem The item to draw. If the list box is empty, this parameter will be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::OwnerDrawItem, ListBox::Raise_OwnerDrawItem,
	///       Fire_MeasureItem
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::OwnerDrawItem, ListBox::Raise_OwnerDrawItem,
	///       Fire_MeasureItem
	/// \endif
	HRESULT Fire_OwnerDrawItem(IListBoxItem* pListItem, OwnerDrawActionConstants requiredAction, OwnerDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pListItem;
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
				hr = pConnection->Invoke(DISPID_LBE_OWNERDRAWITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[0].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ProcessCharacterInput event</em>
	///
	/// \param[in] keyAscii The pressed key's ASCII code.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[out] pItemToSelect The zero-based index of the item to select. If set to -1, the default
	///             processing takes place. If set to -2, the control does not select any item, assuming
	///             that the client application already performed selection. Otherwise the item with
	///             the specified index gets selected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::ProcessCharacterInput, ListBox::Raise_ProcessCharacterInput,
	///       Fire_ProcessKeyStroke, Fire_KeyDown, Fire_KeyPress, Fire_KeyUp
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ProcessCharacterInput, ListBox::Raise_ProcessCharacterInput,
	///       Fire_ProcessKeyStroke, Fire_KeyDown, Fire_KeyPress, Fire_KeyUp
	/// \endif
	HRESULT Fire_ProcessCharacterInput(SHORT keyAscii, SHORT shift, LONG* pItemToSelect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = keyAscii;							p[2].vt = VT_I2;
				p[1] = shift;									p[1].vt = VT_I2;
				p[0].plVal = pItemToSelect;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_LBE_PROCESSCHARACTERINPUT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ProcessKeyStroke event</em>
	///
	/// \param[in] keyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///            enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[out] pItemToSelect The zero-based index of the item to select. If set to -1, the default
	///             processing takes place. If set to -2, the control does not select any item, assuming
	///             that the client application already performed selection. Otherwise the item with
	///             the specified index gets selected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::ProcessKeyStroke, ListBox::Raise_ProcessKeyStroke,
	///       Fire_ProcessCharacterInput, Fire_KeyDown, Fire_KeyPress, Fire_KeyUp
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ProcessKeyStroke, ListBox::Raise_ProcessKeyStroke,
	///       Fire_ProcessCharacterInput, Fire_KeyDown, Fire_KeyPress, Fire_KeyUp
	/// \endif
	HRESULT Fire_ProcessKeyStroke(SHORT keyCode, SHORT shift, LONG* pItemToSelect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = keyCode;								p[2].vt = VT_I2;
				p[1] = shift;									p[1].vt = VT_I2;
				p[0].plVal = pItemToSelect;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_LBE_PROCESSKEYSTROKE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RClick event</em>
	///
	/// \param[in] pListItem The clicked item. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::RClick, ListBox::Raise_RClick, Fire_ContextMenu, Fire_RDblClick,
	///       Fire_Click, Fire_MClick, Fire_XClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::RClick, ListBox::Raise_RClick, Fire_ContextMenu, Fire_RDblClick,
	///       Fire_Click, Fire_MClick, Fire_XClick
	/// \endif
	HRESULT Fire_RClick(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RDblClick event</em>
	///
	/// \param[in] pListItem The double-clicked item. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::RDblClick, ListBox::Raise_RDblClick, Fire_RClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::RDblClick, ListBox::Raise_RDblClick, Fire_RClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_RDblClick(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::RecreatedControlWindow, ListBox::Raise_RecreatedControlWindow,
	///       Fire_DestroyedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::RecreatedControlWindow, ListBox::Raise_RecreatedControlWindow,
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
				hr = pConnection->Invoke(DISPID_LBE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedItem event</em>
	///
	/// \param[in] pListItem The removed item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::RemovedItem, ListBox::Raise_RemovedItem, Fire_RemovingItem,
	///       Fire_InsertedItem
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::RemovedItem, ListBox::Raise_RemovedItem, Fire_RemovingItem,
	///       Fire_InsertedItem
	/// \endif
	HRESULT Fire_RemovedItem(IVirtualListBoxItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_LBE_REMOVEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingItem event</em>
	///
	/// \param[in] pListItem The item being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::RemovingItem, ListBox::Raise_RemovingItem, Fire_RemovedItem,
	///       Fire_InsertingItem
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::RemovingItem, ListBox::Raise_RemovingItem, Fire_RemovedItem,
	///       Fire_InsertingItem
	/// \endif
	HRESULT Fire_RemovingItem(IListBoxItem* pListItem, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListItem;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_LBE_REMOVINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::ResizedControlWindow, ListBox::Raise_ResizedControlWindow
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::ResizedControlWindow, ListBox::Raise_ResizedControlWindow
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
				hr = pConnection->Invoke(DISPID_LBE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectionChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::_IListBoxEvents::SelectionChanged, ListBox::Raise_SelectionChanged,
	///       Fire_CaretChanged
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::SelectionChanged, ListBox::Raise_SelectionChanged,
	///       Fire_CaretChanged
	/// \endif
	HRESULT Fire_SelectionChanged(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_LBE_SELECTIONCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XClick event</em>
	///
	/// \param[in] pListItem The clicked item. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::XClick, ListBox::Raise_XClick, Fire_XDblClick, Fire_Click,
	///       Fire_MClick, Fire_RClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::XClick, ListBox::Raise_XClick, Fire_XDblClick, Fire_Click,
	///       Fire_MClick, Fire_RClick
	/// \endif
	HRESULT Fire_XClick(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XDblClick event</em>
	///
	/// \param[in] pListItem The double-clicked item. May be \c NULL.
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
	///   \sa CBLCtlsLibU::_IListBoxEvents::XDblClick, ListBox::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa CBLCtlsLibA::_IListBoxEvents::XDblClick, ListBox::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \endif
	HRESULT Fire_XDblClick(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_LBE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IListBoxEvents