//////////////////////////////////////////////////////////////////////
/// \class ListBox
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c ListBox</em>
///
/// This class superclasses \c ListBox and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IListBox
/// \else
///   \sa CBLCtlsLibA::IListBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IListBoxEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "ListBoxItem.h"
#include "VirtualListBoxItem.h"
#include "ListBoxItems.h"
#include "ListBoxItemContainer.h"
#include "CommonProperties.h"
#include "TargetOLEDataObject.h"
#include "SourceOLEDataObject.h"


class ATL_NO_VTABLE ListBox : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IListBox, &IID_IListBox, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IListBox, &IID_IListBox, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<ListBox>,
    public IOleControlImpl<ListBox>,
    public IOleObjectImpl<ListBox>,
    public IOleInPlaceActiveObjectImpl<ListBox>,
    public IViewObjectExImpl<ListBox>,
    public IOleInPlaceObjectWindowlessImpl<ListBox>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListBox>,
    public Proxy_IListBoxEvents<ListBox>,
    public IPersistStorageImpl<ListBox>,
    public IPersistPropertyBagImpl<ListBox>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<ListBox>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_ListBox, &__uuidof(_IListBoxEvents), &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_ListBox, &__uuidof(_IListBoxEvents), &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<ListBox>,
    public CComCoClass<ListBox, &CLSID_ListBox>,
    public CComControl<ListBox>,
    public IPerPropertyBrowsingImpl<ListBox>,
    public IDropTarget,
    public IDropSource,
    public IDropSourceNotify,
    public ICategorizeProperties,
    public ICreditsProvider
{
	friend class ListBoxItem;
	friend class ListBoxItems;
	friend class ListBoxItemContainer;
	friend class SourceOLEDataObject;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	ListBox();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTBOX)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("ListBoxU"), WC_LISTBOXW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("ListBoxA"), WC_LISTBOXA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(ListBox)
			COM_INTERFACE_ENTRY(IListBox)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IQuickActivate)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IDropTarget)
			COM_INTERFACE_ENTRY(IDropSource)
			COM_INTERFACE_ENTRY(IDropSourceNotify)
		END_COM_MAP()

		BEGIN_PROP_MAP(ListBox)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AllowDragDrop", DISPID_LB_ALLOWDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowItemSelection", DISPID_LB_ALLOWITEMSELECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AlwaysShowVerticalScrollBar", DISPID_LB_ALWAYSSHOWVERTICALSCROLLBAR, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_LB_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BackColor", DISPID_LB_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_LB_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ColumnWidth", DISPID_LB_COLUMNWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_LB_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_LB_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragScrollTimeBase", DISPID_LB_DRAGSCROLLTIMEBASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Enabled", DISPID_LB_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_LB_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("ForeColor", DISPID_LB_FORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("HasStrings", DISPID_LB_HASSTRINGS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("HoverTime", DISPID_LB_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IMEMode", DISPID_LB_IMEMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("InsertMarkColor", DISPID_LB_INSERTMARKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("InsertMarkStyle", DISPID_LB_INSERTMARKSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IntegralHeight", DISPID_LB_INTEGRALHEIGHT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ItemHeight", DISPID_LB_ITEMHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Locale", DISPID_LB_LOCALE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_LB_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_LB_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MultiColumn", DISPID_LB_MULTICOLUMN, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MultiSelect", DISPID_LB_MULTISELECT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("OLEDragImageStyle", DISPID_LB_OLEDRAGIMAGESTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("OwnerDrawItems", DISPID_LB_OWNERDRAWITEMS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_LB_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ProcessTabs", DISPID_LB_PROCESSTABS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_LB_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_LB_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ScrollableWidth", DISPID_LB_SCROLLABLEWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Sorted", DISPID_LB_SORTED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_LB_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("TabWidth", DISPID_LB_TABWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ToolTips", DISPID_LB_TOOLTIPS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_LB_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("VirtualMode", DISPID_LB_VIRTUALMODE, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListBox)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IListBoxEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(ListBox)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_HSCROLL, OnScroll)
			MESSAGE_HANDLER(WM_INPUTLANGCHANGE, OnInputLangChange)
			MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_NCMOUSEMOVE, OnNCMouseMove)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(WM_SIZE, OnSize)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_VSCROLL, OnScroll)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(OCM_CHARTOITEM, OnReflectedCharToItem)
			MESSAGE_HANDLER(OCM_COMPAREITEM, OnReflectedCompareItem)
			MESSAGE_HANDLER(OCM_CTLCOLORLISTBOX, OnReflectedCtlColorListBox)
			MESSAGE_HANDLER(OCM_DRAWITEM, OnReflectedDrawItem)
			MESSAGE_HANDLER(OCM_MEASUREITEM, OnReflectedMeasureItem)
			MESSAGE_HANDLER(OCM_VKEYTOITEM, OnReflectedVKeyToItem)

			MESSAGE_HANDLER(GetDragImageMessage(), OnGetDragImage)

			MESSAGE_HANDLER(LB_ADDSTRING, OnAddString)
			MESSAGE_HANDLER(LB_DELETESTRING, OnDeleteString)
			MESSAGE_HANDLER(LB_INSERTSTRING, OnInsertString)
			MESSAGE_HANDLER(LB_RESETCONTENT, OnResetContent)
			MESSAGE_HANDLER(LB_SELECTSTRING, OnSelectionChange)
			MESSAGE_HANDLER(LB_SELITEMRANGE, OnSelectionChange)
			MESSAGE_HANDLER(LB_SELITEMRANGEEX, OnSelectionChange)
			MESSAGE_HANDLER(LB_SETCARETINDEX, OnSetCaretIndex)
			MESSAGE_HANDLER(LB_SETCOLUMNWIDTH, OnSetColumnWidth)
			MESSAGE_HANDLER(LB_SETCURSEL, OnSelectionChange)
			MESSAGE_HANDLER(LB_SETSEL, OnSelectionChange)
			MESSAGE_HANDLER(LB_SETTABSTOPS, OnSetTabStops)

			REFLECTED_COMMAND_CODE_HANDLER(LBN_ERRSPACE, OnReflectedErrSpace)
			//REFLECTED_COMMAND_CODE_HANDLER(LBN_SELCANCEL, OnReflectedSelCancel)
			REFLECTED_COMMAND_CODE_HANDLER(LBN_SELCHANGE, OnReflectedSelChange)
			REFLECTED_COMMAND_CODE_HANDLER(LBN_SETFOCUS, OnReflectedSetFocus)

			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOA, OnToolTipGetDispInfoNotificationA)
			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOW, OnToolTipGetDispInfoNotificationW)
			NOTIFY_CODE_HANDLER(TTN_SHOW, OnToolTipShowNotification)

			CHAIN_MSG_MAP(CComControl<ListBox>)
		END_MSG_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IListBox
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AllowDragDrop property</em>
	///
	/// Retrieves whether drag'n'drop mode can be entered. If set to \c VARIANT_TRUE, drag'n'drop mode
	/// can be entered by pressing the left or right mouse button over an item and then moving the
	/// mouse with the button still pressed. If set to \c VARIANT_FALSE, drag'n'drop mode is not
	/// available.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowDragDrop, get_RegisterForOLEDragDrop, get_DragScrollTimeBase, SetInsertMarkPosition,
	///     Raise_ItemBeginDrag, Raise_ItemBeginRDrag
	virtual HRESULT STDMETHODCALLTYPE get_AllowDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowDragDrop property</em>
	///
	/// Sets whether drag'n'drop mode can be entered. If set to \c VARIANT_TRUE, drag'n'drop mode
	/// can be entered by pressing the left or right mouse button over an item and then moving the
	/// mouse with the button still pressed. If set to \c VARIANT_FALSE, drag'n'drop mode is not
	/// available.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowDragDrop, put_RegisterForOLEDragDrop, put_DragScrollTimeBase, SetInsertMarkPosition,
	///     Raise_ItemBeginDrag, Raise_ItemBeginRDrag
	virtual HRESULT STDMETHODCALLTYPE put_AllowDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowItemSelection property</em>
	///
	/// Retrieves whether list box items can be selected by the user. If set to \c VARIANT_TRUE, the user
	/// can select items; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowItemSelection, get_MultiSelect
	virtual HRESULT STDMETHODCALLTYPE get_AllowItemSelection(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowItemSelection property</em>
	///
	/// Sets whether list box items can be selected by the user. If set to \c VARIANT_TRUE, the user
	/// can select items; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_AllowItemSelection, put_MultiSelect
	virtual HRESULT STDMETHODCALLTYPE put_AllowItemSelection(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AlwaysShowVerticalScrollBar property</em>
	///
	/// Retrieves whether the vertical scroll bar is disabled instead of hidden if the control does not
	/// contain enough items to scroll. If set to \c VARIANT_TRUE, the scroll bar is disabled; otherwise it
	/// is hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AlwaysShowVerticalScrollBar, get_ListItems, get_ScrollableWidth
	virtual HRESULT STDMETHODCALLTYPE get_AlwaysShowVerticalScrollBar(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AlwaysShowVerticalScrollBar property</em>
	///
	/// Sets whether the vertical scroll bar is disabled instead of hidden if the control does not
	/// contain enough items to scroll. If set to \c VARIANT_TRUE, the scroll bar is disabled; otherwise it
	/// is hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_AlwaysShowVerticalScrollBar, get_ListItems, put_ScrollableWidth
	virtual HRESULT STDMETHODCALLTYPE put_AlwaysShowVerticalScrollBar(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AnchorItem property</em>
	///
	/// Retrieves the control's anchor item. The anchor item is the item with which range-selection
	/// begins.
	///
	/// \param[out] ppAnchorItem Receives the anchor item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_AnchorItem, ListBoxItem::get_Anchor, get_MultiSelect, get_CaretItem
	virtual HRESULT STDMETHODCALLTYPE get_AnchorItem(IListBoxItem** ppAnchorItem);
	/// \brief <em>Sets the \c AnchorItem property</em>
	///
	/// Sets the control's anchor item. The anchor item is the item with which range-selection begins.
	///
	/// \param[in] pNewAnchorItem The new anchor item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AnchorItem, put_MultiSelect, putref_CaretItem
	virtual HRESULT STDMETHODCALLTYPE putref_AnchorItem(IListBoxItem* pNewAnchorItem);
	/// \brief <em>Retrieves the current setting of the \c Appearance property</em>
	///
	/// Retrieves the kind of border that is drawn around the control. Any of the values defined by
	/// the \c AppearanceConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Appearance, get_BorderStyle, CBLCtlsLibU::AppearanceConstants
	/// \else
	///   \sa put_Appearance, get_BorderStyle, CBLCtlsLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Appearance(AppearanceConstants* pValue);
	/// \brief <em>Sets the \c Appearance property</em>
	///
	/// Sets the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Appearance, put_BorderStyle, CBLCtlsLibU::AppearanceConstants
	/// \else
	///   \sa get_Appearance, put_BorderStyle, CBLCtlsLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Appearance(AppearanceConstants newValue);
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BackColor, get_ForeColor, get_InsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BackColor, put_ForeColor, put_InsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of inner border that is drawn around the control. Any of the values defined
	/// by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BorderStyle, get_Appearance, CBLCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, get_Appearance, CBLCtlsLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(BorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of inner border that is drawn around the control. Any of the values defined by
	/// the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BorderStyle, put_Appearance, CBLCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, put_Appearance, CBLCtlsLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(BorderStyleConstants newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c CaretItem property</em>
	///
	/// Retrieves the control's caret item. The caret item is the item that has the focus.
	///
	/// \param[in] partialVisibilityIsOkay Ignored.
	/// \param[out] ppCaretItem Receives the caret item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_CaretItem, ListBoxItem::get_Caret, get_AnchorItem
	virtual HRESULT STDMETHODCALLTYPE get_CaretItem(VARIANT_BOOL partialVisibilityIsOkay = VARIANT_FALSE, IListBoxItem** ppCaretItem = NULL);
	/// \brief <em>Sets the \c CaretItem property</em>
	///
	/// Sets the control's caret item. The caret item is the item that has the focus.
	///
	/// \param[in] partialVisibilityIsOkay If set to \c VARIANT_TRUE, the control is scrolled until the caret
	///            item is at least partially visible. If set to \c VARIANT_FALSE, the control is scrolled
	///            until the entire caret item is visible.
	/// \param[in] pNewCaretItem The new caret item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks In single-selection mode the caret item can be set only if the \c SelectedItem property is
	///          set to \c NULL.
	///
	/// \sa get_CaretItem, put_MultiSelect, putref_AnchorItem
	virtual HRESULT STDMETHODCALLTYPE putref_CaretItem(VARIANT_BOOL partialVisibilityIsOkay = VARIANT_FALSE, IListBoxItem* pNewCaretItem = NULL);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c ColumnWidth property</em>
	///
	/// Retrieves the width of a column in a multi-column list box. If set to -1, the system's default
	/// setting is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ColumnWidth, get_MultiColumn
	virtual HRESULT STDMETHODCALLTYPE get_ColumnWidth(LONG* pValue);
	/// \brief <em>Sets the \c ColumnWidth property</em>
	///
	/// Sets the width of a column in a multi-column list box. If set to -1, the system's default
	/// setting is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ColumnWidth, put_MultiColumn
	virtual HRESULT STDMETHODCALLTYPE put_ColumnWidth(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisabledEvents, CBLCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa put_DisabledEvents, CBLCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisabledEvents, CBLCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa get_DisabledEvents, CBLCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DontRedraw property</em>
	///
	/// Retrieves whether automatic redrawing of the control is enabled or disabled. If set to
	/// \c VARIANT_FALSE, the control will redraw itself automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE get_DontRedraw(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DontRedraw property</em>
	///
	/// Enables or disables automatic redrawing of the control. Disabling redraw while doing large changes
	/// on the control may increase performance. If set to \c VARIANT_FALSE, the control will redraw itself
	/// automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE put_DontRedraw(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DraggedItems property</em>
	///
	/// Retrieves a collection object wrapping the control's items that are currently dragged. These
	/// are the same items that were passed to the \c BeginDrag or \c OLEDrag method.
	///
	/// \param[out] ppItems Receives the collection object's \c IListBoxItemContainer implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa BeginDrag, OLEDrag, ListBoxItemContainer, get_ListItems
	virtual HRESULT STDMETHODCALLTYPE get_DraggedItems(IListBoxItemContainer** ppItems);
	/// \brief <em>Retrieves the current setting of the \c DragScrollTimeBase property</em>
	///
	/// Retrieves the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time, divided by 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragScrollTimeBase, get_AllowDragDrop, get_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragScrollTimeBase(LONG* pValue);
	/// \brief <em>Sets the \c DragScrollTimeBase property</em>
	///
	/// Sets the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time divided by 4 is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragScrollTimeBase, put_AllowDragDrop, put_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragScrollTimeBase(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the control for user input. If set to \c VARIANT_TRUE, the control reacts
	/// to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FirstVisibleItem property</em>
	///
	/// Retrieves the first list box item, that is entirely located within the control's client area and
	/// therefore visible to the user.
	///
	/// \param[out] ppFirstItem Receives the first visible item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_FirstVisibleItem, get_ListItems
	virtual HRESULT STDMETHODCALLTYPE get_FirstVisibleItem(IListBoxItem** ppFirstItem);
	/// \brief <em>Sets the \c FirstVisibleItem property</em>
	///
	/// Sets the first list box item, that is entirely located within the control's client area and
	/// therefore visible to the user.
	///
	/// \param[in] pNewFirstItem The new first visible item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstVisibleItem, get_ListItems
	virtual HRESULT STDMETHODCALLTYPE putref_FirstVisibleItem(IListBoxItem* pNewFirstItem);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the control's content.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont, get_ForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont, put_ForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Font, put_Font, put_UseSystemFont, put_ForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the control's text color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ForeColor, get_BackColor, get_InsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ForeColor, put_BackColor, put_InsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c HasStrings property</em>
	///
	/// Retrieves whether the items in the control are strings. If set to \c VARIANT_TRUE, the control
	/// contains strings; otherwise the items consist of the value specified by the \c ItemData property
	/// only.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c OwnerDrawItems property is set to \c odiDontOwnerDraw.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_HasStrings, get_OwnerDrawItems, ListBoxItem::get_ItemData, ListBoxItem::get_Text
	virtual HRESULT STDMETHODCALLTYPE get_HasStrings(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c HasStrings property</em>
	///
	/// Sets whether the items in the control are strings. If set to \c VARIANT_TRUE, the control
	/// contains strings; otherwise the items consist of the value specified by the \c ItemData property
	/// only.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c OwnerDrawItems property is set to \c odiDontOwnerDraw.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_HasStrings, put_OwnerDrawItems, ListBoxItem::put_ItemData, ListBoxItem::put_Text
	virtual HRESULT STDMETHODCALLTYPE put_HasStrings(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c hDragImageList property</em>
	///
	/// Retrieves the handle to the image list containing the drag image that is used during a
	/// drag'n'drop operation to visualize the dragged items.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, BeginDrag, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_hDragImageList(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hImageList property</em>
	///
	/// Retrieves the handle of the specified imagelist.
	///
	/// \param[in] imageList The imageList to retrieve. Only the \c ilHighResolution value defined
	///            by the \c ImageListConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hImageList property</em>
	///
	/// Sets the handle of the specified imagelist.
	///
	/// \param[in] imageList The imageList to set. Only the \c ilHighResolution value defined
	///            by the \c ImageListConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE get_HoverTime(LONG* pValue);
	/// \brief <em>Sets the \c HoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE put_HoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndToolTip property</em>
	///
	/// Retrieves the tooltip control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hWndToolTip, get_ToolTips
	virtual HRESULT STDMETHODCALLTYPE get_hWndToolTip(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hWndToolTip property</em>
	///
	/// Sets the window handle of the tooltip control to use.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set tooltip window does NOT get destroyed automatically.
	///
	/// \sa get_hWndToolTip, put_ToolTips
	virtual HRESULT STDMETHODCALLTYPE put_hWndToolTip(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c IMEMode property</em>
	///
	/// Retrieves the control's IME mode. IME is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IMEMode, CBLCtlsLibU::IMEModeConstants
	/// \else
	///   \sa put_IMEMode, CBLCtlsLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IMEMode(IMEModeConstants* pValue);
	/// \brief <em>Sets the \c IMEMode property</em>
	///
	/// Sets the control's IME mode. IME is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IMEMode, CBLCtlsLibU::IMEModeConstants
	/// \else
	///   \sa get_IMEMode, CBLCtlsLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IMEMode(IMEModeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c InsertMarkColor property</em>
	///
	/// Retrieves the color that is the control's insertion mark is drawn in.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c InsertMarkStyle property is set to \c imsNative.
	///
	/// \sa put_InsertMarkColor, get_InsertMarkStyle, GetInsertMarkPosition, get_BackColor, get_ForeColor
	virtual HRESULT STDMETHODCALLTYPE get_InsertMarkColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c InsertMarkColor property</em>
	///
	/// Sets the color that is the control's insertion mark is drawn in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c InsertMarkStyle property is set to \c imsNative.
	///
	/// \sa get_InsertMarkColor, put_InsertMarkStyle, SetInsertMarkPosition, put_BackColor, put_ForeColor
	virtual HRESULT STDMETHODCALLTYPE put_InsertMarkColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c InsertMarkStyle property</em>
	///
	/// Retrieves the style of the control's insertion mark. Any of the values defined by the
	/// \c InsertMarkStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_InsertMarkStyle, get_InsertMarkColor, GetInsertMarkPosition,
	///       CBLCtlsLibU::InsertMarkStyleConstants
	/// \else
	///   \sa put_InsertMarkStyle, get_InsertMarkColor, GetInsertMarkPosition,
	///       CBLCtlsLibA::InsertMarkStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_InsertMarkStyle(InsertMarkStyleConstants* pValue);
	/// \brief <em>Sets the \c InsertMarkStyle property</em>
	///
	/// Sets the style of the control's insertion mark. Any of the values defined by the
	/// \c InsertMarkStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_InsertMarkStyle, put_InsertMarkColor, SetInsertMarkPosition,
	///       CBLCtlsLibU::InsertMarkStyleConstants
	/// \else
	///   \sa get_InsertMarkStyle, put_InsertMarkColor, SetInsertMarkPosition,
	///       CBLCtlsLibA::InsertMarkStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_InsertMarkStyle(InsertMarkStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c IntegralHeight property</em>
	///
	/// Sets whether the control resizes itself so that an integral number of items is displayed.
	/// If set to \c VARIANT_TRUE, an integral number of items is displayed and the control's height
	/// may be changed to achieve this; otherwise partial items may be displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_IntegralHeight, get_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE get_IntegralHeight(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c IntegralHeight property</em>
	///
	/// Retrieves whether the control resizes itself so that an integral number of items is displayed.
	/// If set to \c VARIANT_TRUE, an integral number of items is displayed and the control's height
	/// may be changed to achieve this; otherwise partial items may be displayed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_IntegralHeight, put_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE put_IntegralHeight(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemHeight property</em>
	///
	/// Retrieves the items' height in pixels. If set to -1, the default setting is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c OwnerDrawItems property is set to \c odiOwnerDrawVariableHeight, this property
	///          is ignored. Use the \c IListBoxItem::Height property instead.
	///
	/// \sa put_ItemHeight, get_OwnerDrawItems, get_IntegralHeight, ListBoxItem::get_Height
	virtual HRESULT STDMETHODCALLTYPE get_ItemHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ItemHeight property</em>
	///
	/// Sets the items' height in pixels. If set to -1, the default setting is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c OwnerDrawItems property is set to \c odiOwnerDrawVariableHeight, this property
	///          is ignored. Use the \c IListBoxItem::Height property instead.
	///
	/// \sa get_ItemHeight, put_OwnerDrawItems, put_IntegralHeight, ListBoxItem::put_Height
	virtual HRESULT STDMETHODCALLTYPE put_ItemHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ItemsPerColumn property</em>
	///
	/// Retrieves the number of items displayed in a column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MultiColumn
	virtual HRESULT STDMETHODCALLTYPE get_ItemsPerColumn(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ListItems property</em>
	///
	/// Retrieves a collection object wrapping the control's list box items.
	///
	/// \param[out] ppItems Receives the collection object's \c IListBoxItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ListBoxItems
	virtual HRESULT STDMETHODCALLTYPE get_ListItems(IListBoxItems** ppItems);
	/// \brief <em>Retrieves the current setting of the \c Locale property</em>
	///
	/// Retrieves the control's current locale. The locale influences how items are sorted if the
	/// \c Sorted property is set to \c VARIANT_TRUE. For a list of possible locale identifiers see
	/// <a href="https://msdn.microsoft.com/en-us/library/dd318693.aspx">MSDN Online</a>.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Locale, get_Sorted
	virtual HRESULT STDMETHODCALLTYPE get_Locale(LONG* pValue);
	/// \brief <em>Sets the \c Locale property</em>
	///
	/// Sets the control's current locale. The locale influences how items are sorted if the
	/// \c Sorted property is set to \c VARIANT_TRUE. For a list of possible locale identifiers see
	/// <a href="https://msdn.microsoft.com/en-us/library/dd318693.aspx">MSDN Online</a>.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Locale, put_Sorted
	virtual HRESULT STDMETHODCALLTYPE put_Locale(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, CBLCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, CBLCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, CBLCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, CBLCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, CBLCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, CBLCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c MousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MousePointer, get_MouseIcon, CBLCtlsLibU::MousePointerConstants
	/// \else
	///   \sa put_MousePointer, get_MouseIcon, CBLCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c MousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MousePointer, put_MouseIcon, CBLCtlsLibU::MousePointerConstants
	/// \else
	///   \sa get_MousePointer, put_MouseIcon, CBLCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiColumn property</em>
	///
	/// Retrieves whether the control is scrolled horizontally and the items are displayed in multiple
	/// columns. If set to \c VARIANT_TRUE, the items are displayed in multiple columns; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MultiColumn, get_ColumnWidth, get_ItemsPerColumn
	virtual HRESULT STDMETHODCALLTYPE get_MultiColumn(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MultiColumn property</em>
	///
	/// Sets whether the control is scrolled horizontally and the items are displayed in multiple
	/// columns. If set to \c VARIANT_TRUE, the items are displayed in multiple columns; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_MultiColumn, put_ColumnWidth, get_ItemsPerColumn
	virtual HRESULT STDMETHODCALLTYPE put_MultiColumn(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiSelect property</em>
	///
	/// Retrieves whether multiple items can be selected at the same time and how an item's selection state
	/// can be changed by the user. Any of the values defined by the \c MultiSelectConstants enumeration is
	/// valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MultiSelect, CBLCtlsLibU::MultiSelectConstants
	/// \else
	///   \sa put_MultiSelect, CBLCtlsLibA::MultiSelectConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MultiSelect(MultiSelectConstants* pValue);
	/// \brief <em>Sets the \c MultiSelect property</em>
	///
	/// Sets whether multiple items can be selected at the same time and how an item's selection state
	/// can be changed by the user. Any of the values defined by the \c MultiSelectConstants enumeration is
	/// valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_MultiSelect, CBLCtlsLibU::MultiSelectConstants
	/// \else
	///   \sa get_MultiSelect, CBLCtlsLibA::MultiSelectConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MultiSelect(MultiSelectConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c OLEDragImageStyle property</em>
	///
	/// Retrieves the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       CBLCtlsLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       CBLCtlsLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue);
	/// \brief <em>Sets the \c OLEDragImageStyle property</em>
	///
	/// Sets the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       CBLCtlsLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       CBLCtlsLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OLEDragImageStyle(OLEDragImageStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c OwnerDrawItems property</em>
	///
	/// Retrieves whether the client application draws the items itself. Any of the values defined by the
	/// \c OwnerDrawItemsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OwnerDrawItems, get_HasStrings, CBLCtlsLibU::OwnerDrawItemsConstants
	/// \else
	///   \sa put_OwnerDrawItems, get_HasStrings, CBLCtlsLibA::OwnerDrawItemsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OwnerDrawItems(OwnerDrawItemsConstants* pValue);
	/// \brief <em>Sets the \c OwnerDrawItems property</em>
	///
	/// Sets whether the client application draws the items itself. Any of the values defined by the
	/// \c OwnerDrawItemsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_OwnerDrawItems, put_HasStrings, CBLCtlsLibU::OwnerDrawItemsConstants
	/// \else
	///   \sa get_OwnerDrawItems, put_HasStrings, CBLCtlsLibA::OwnerDrawItemsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OwnerDrawItems(OwnerDrawItemsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE get_ProcessContextMenuKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessContextMenuKeys property</em>
	///
	/// Sets whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE put_ProcessContextMenuKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessTabs property</em>
	///
	/// Retrieves whether tab characters in the item texts are expanded or ignored. If set to
	/// \c VARIANT_TRUE, tab characters are expanded resulting in multi-column texts; otherwise tab
	/// characters are ignored.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessTabs, get_TabStops, ListBoxItem::get_Text
	virtual HRESULT STDMETHODCALLTYPE get_ProcessTabs(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessTabs property</em>
	///
	/// Sets whether tab characters in the item texts are expanded or ignored. If set to
	/// \c VARIANT_TRUE, tab characters are expanded resulting in multi-column texts; otherwise tab
	/// characters are ignored.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_ProcessTabs, put_TabStops, ListBoxItem::put_Text
	virtual HRESULT STDMETHODCALLTYPE put_ProcessTabs(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RegisterForOLEDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RegisterForOLEDragDrop property</em>
	///
	/// Sets whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RegisterForOLEDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether bidirectional features are enabled or disabled. Any combination of the values
	/// defined by the \c RightToLeftConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RightToLeft, CBLCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa put_RightToLeft, CBLCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(RightToLeftConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Enables or disables bidirectional features. Any combination of the values defined by the
	/// \c RightToLeftConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Setting or clearing the \c rtlLayout flag destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, CBLCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, CBLCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ScrollableWidth property</em>
	///
	/// Retrieves the width in pixels, by which the list box can be scrolled horizontally. If the width of
	/// the control is greater than this value, a horizontal scroll bar is displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ScrollableWidth, get_AlwaysShowVerticalScrollBar
	virtual HRESULT STDMETHODCALLTYPE get_ScrollableWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ScrollableWidth property</em>
	///
	/// Sets the width in pixels, by which the list box can be scrolled horizontally. If the width of
	/// the control is greater than this value, a horizontal scroll bar is displayed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ScrollableWidth, put_AlwaysShowVerticalScrollBar
	virtual HRESULT STDMETHODCALLTYPE put_ScrollableWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c SelectedItem property</em>
	///
	/// Retrieves the control's currently selected item.
	///
	/// \param[out] ppSelectedItem Receives the selected item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_SelectedItem, get_SelectedItemsArray, ListBoxItem::get_Selected, get_MultiSelect,
	///     get_CaretItem, SelectItemByText
	virtual HRESULT STDMETHODCALLTYPE get_SelectedItem(IListBoxItem** ppSelectedItem);
	/// \brief <em>Sets the \c SelectedItem property</em>
	///
	/// Sets the control's currently selected item.
	///
	/// \param[in] pNewSelectedItem The new selected item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_SelectedItem, put_SelectedItemsArray, ListBoxItem::put_Selected, put_MultiSelect,
	///     putref_CaretItem, SelectItemByText
	virtual HRESULT STDMETHODCALLTYPE putref_SelectedItem(IListBoxItem* pNewSelectedItem);
	/// \brief <em>Retrieves the current setting of the \c SelectedItemsArray property</em>
	///
	/// Retrieves the selected items. The sub-type of this property is an array of \c Longs. Each element
	/// of this array contains the index of a selected item.
	///
	/// \sa put_SelectedItemsArray, get_SelectedItem, ListBoxItem::get_Selected, get_MultiSelect
	virtual HRESULT STDMETHODCALLTYPE get_SelectedItemsArray(VARIANT* pValue);
	/// \brief <em>Sets the \c SelectedItemsArray property</em>
	///
	/// Sets the selected items. The sub-type of this property is an array of \c Longs. Each element
	/// of this array contains the index of a selected item.
	///
	/// \sa get_SelectedItemsArray, putref_SelectedItem, ListBoxItem::put_Selected, put_MultiSelect
	virtual HRESULT STDMETHODCALLTYPE put_SelectedItemsArray(VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowDragImage property</em>
	///
	/// Retrieves whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible;
	/// otherwise hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowDragImage, get_hDragImageList, get_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_ShowDragImage(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowDragImage property</em>
	///
	/// Sets whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible; otherwise
	/// hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, get_hDragImageList, put_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_ShowDragImage(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Sorted property</em>
	///
	/// Retrieves whether the items in the control are sorted alphabetically. If set to \c VARIANT_TRUE,
	/// the control sorts the items alphabetically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Sorted, get_ListItems, get_Locale, Raise_CompareItems
	virtual HRESULT STDMETHODCALLTYPE get_Sorted(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Sorted property</em>
	///
	/// Sets whether the items in the control are sorted alphabetically. If set to \c VARIANT_TRUE,
	/// the control sorts the items alphabetically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_Sorted, get_ListItems, put_Locale, Raise_CompareItems
	virtual HRESULT STDMETHODCALLTYPE put_Sorted(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SupportOLEDragImages property</em>
	///
	/// Retrieves whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, FinishOLEDragDrop, get_hImageList,
	///     get_OLEDragImageStyle,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE get_SupportOLEDragImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SupportOLEDragImages property</em>
	///
	/// Sets whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, FinishOLEDragDrop, put_hImageList,
	///     put_OLEDragImageStyle,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c TabStops property</em>
	///
	/// Retrieves the positions (in pixels) of the control's tab stops. The property expects a \c VARIANT
	/// containing an array of integer values, each specifying a tab stop's position.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c ProcessTabs property is set to \c VARIANT_FALSE.
	///
	/// \sa put_TabStops, get_ProcessTabs, get_TabWidth
	virtual HRESULT STDMETHODCALLTYPE get_TabStops(VARIANT* pValue);
	/// \brief <em>Sets the \c TabStops property</em>
	///
	/// Sets the positions (in pixels) of the control's tab stops. The property expects a \c VARIANT
	/// containing an array of integer values, each specifying a tab stop's position.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c ProcessTabs property is set to \c VARIANT_FALSE.
	///
	/// \sa get_TabStops, put_ProcessTabs, put_TabWidth
	virtual HRESULT STDMETHODCALLTYPE put_TabStops(VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c TabWidth property</em>
	///
	/// Retrieves the distance (in pixels) between 2 tab stops. If set to -1, the system's default value is
	/// used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c ProcessTabs property is set to \c VARIANT_FALSE.\n
	///          This property is ignored, if the \c TabStops property is not set to \c Empty.
	///
	/// \sa put_TabWidth, get_ProcessTabs, get_TabStops
	virtual HRESULT STDMETHODCALLTYPE get_TabWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c TabWidth property</em>
	///
	/// Sets the distance (in pixels) between 2 tab stops. If set to -1, the system's default value is
	/// used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c ProcessTabs property is set to \c VARIANT_FALSE.\n
	///          This property is ignored, if the \c TabStops property is not set to \c Empty.
	///
	/// \sa get_TabWidth, put_ProcessTabs, put_TabStops
	virtual HRESULT STDMETHODCALLTYPE put_TabWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c ToolTips property</em>
	///
	/// Retrieves which kinds of tooltips the control is displaying. Any combination of the values
	/// defined by the \c ToolTipsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ToolTips, get_hWndToolTip, Raise_ItemGetInfoTipText, CBLCtlsLibU::ToolTipsConstants
	/// \else
	///   \sa put_ToolTips, get_hWndToolTip, Raise_ItemGetInfoTipText, CBLCtlsLibA::ToolTipsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ToolTips(ToolTipsConstants* pValue);
	/// \brief <em>Sets the \c ToolTips property</em>
	///
	/// Sets which kinds of tooltips the control is displaying. Any combination of the values
	/// defined by the \c ToolTipsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ToolTips, put_hWndToolTip, Raise_ItemGetInfoTipText, CBLCtlsLibU::ToolTipsConstants
	/// \else
	///   \sa get_ToolTips, put_hWndToolTip, Raise_ItemGetInfoTipText, CBLCtlsLibA::ToolTipsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ToolTips(ToolTipsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSystemFont property</em>
	///
	/// Retrieves whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseSystemFont, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSystemFont property</em>
	///
	/// Sets whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseSystemFont, put_Font, putref_Font
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c VirtualItemCount property</em>
	///
	/// Retrieves the number of items in the control if the \c VirtualMode property is set to
	/// \c VARIANT_TRUE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_VirtualItemCount, get_VirtualMode
	virtual HRESULT STDMETHODCALLTYPE get_VirtualItemCount(LONG* pValue);
	/// \brief <em>Sets the \c VirtualItemCount property</em>
	///
	/// Sets the number of items in the control if the \c VirtualMode property is set to \c VARIANT_TRUE.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_VirtualItemCount, put_VirtualMode
	virtual HRESULT STDMETHODCALLTYPE put_VirtualItemCount(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c VirtualMode property</em>
	///
	/// Retrieves whether the data-management for the control is done by the client or by the control itself.
	/// If set to \c VARIANT_TRUE, data-management is done by the client; otherwise it's done by the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Using your own data-management system may increase performance, especially if the
	///          list box contains many items. However, virtual list box controls have several
	///          limitations (list may be incomplete):\n
	///          - The \c HasStrings property is not supported.
	///          - The \c Sorted property is not supported.
	///          - The \c OwnerDrawItems property must be set to \c odiOwnerDrawFixedHeight.
	///
	/// \sa put_VirtualMode, get_VirtualItemCount, get_HasStrings, get_Sorted, get_OwnerDrawItems
	virtual HRESULT STDMETHODCALLTYPE get_VirtualMode(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c VirtualMode property</em>
	///
	/// Sets whether the data-management for the control is done by the client or by the control itself.
	/// If set to \c VARIANT_TRUE, data-management is done by the client; otherwise it's done by the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Using your own data-management system may increase performance, especially if the
	///          list box contains many items. However, virtual list box controls have several
	///          limitations (list may be incomplete):\n
	///          - The \c HasStrings property is not supported.
	///          - The \c Sorted property is not supported.
	///          - The \c OwnerDrawItems property must be set to \c odiOwnerDrawFixedHeight.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_VirtualMode, put_VirtualItemCount, put_HasStrings, put_Sorted, put_OwnerDrawItems
	virtual HRESULT STDMETHODCALLTYPE put_VirtualMode(VARIANT_BOOL newValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Enters drag'n'drop mode</em>
	///
	/// \param[in] pDraggedItems The dragged items collection object's \c IListBoxItemContainer
	///            implementation.
	/// \param[in] hDragImageList The image list containing the drag image that shall be used to
	///            visualize the drag'n'drop operation. If -1, the method creates the drag image itself;
	///            if \c NULL, no drag image is used.
	/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
	///                drag image's upper-left corner. If the \c hDragImageList parameter is set to -1 or
	///                \c NULL, this parameter is ignored. This parameter will be changed to the value that
	///                finally was used by the method.
	/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
	///                drag image's upper-left corner. If the \c hDragImageList parameter is set to -1 or
	///                \c NULL, this parameter is ignored. This parameter will be changed to the value that
	///                finally was used by the method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OLEDrag, get_DraggedItems, EndDrag, get_hDragImageList, Raise_ItemBeginDrag,
	///     Raise_ItemBeginRDrag, ListBoxItem::CreateDragImage, ListBoxItemContainer::CreateDragImage
	virtual HRESULT STDMETHODCALLTYPE BeginDrag(IListBoxItemContainer* pDraggedItems, OLE_HANDLE hDragImageList = NULL, OLE_XPOS_PIXELS* pXHotSpot = NULL, OLE_YPOS_PIXELS* pYHotSpot = NULL);
	/// \brief <em>Creates a new \c ListBoxItemContainer object</em>
	///
	/// Retrieves a new \c ListBoxItemContainer object and fills it with the specified items.
	///
	/// \param[in] items The item(s) to add to the collection. May be either \c Empty, an item ID, a
	///            \c ListBoxItem object or a \c ListBoxItems collection.
	/// \param[out] ppContainer Receives the new collection object's \c IListBoxItemContainer
	///             implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ListBoxItemContainer::Clone, ListBoxItemContainer::Add
	virtual HRESULT STDMETHODCALLTYPE CreateItemContainer(VARIANT items = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IListBoxItemContainer** ppContainer = NULL);
	/// \brief <em>Deselects a range of items</em>
	///
	/// Deselects the specified range of items.
	///
	/// \param[in] firstItem The first item to deselect. If \c Empty, the control's first item is used.
	/// \param[in] lastItem The last item to deselect. If \c Empty, the control's first item is used.
	///
	/// \sa SelectItems, get_MultiSelect, ListBoxItem::put_Selected
	virtual HRESULT STDMETHODCALLTYPE DeselectItems(VARIANT firstItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), VARIANT lastItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR));
	/// \brief <em>Exits drag'n'drop mode</em>
	///
	/// \param[in] abort If \c VARIANT_TRUE, the drag'n'drop operation will be handled as aborted;
	///            otherwise it will be handled as a drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DraggedItems, BeginDrag, Raise_AbortedDrag, Raise_Drop
	virtual HRESULT STDMETHODCALLTYPE EndDrag(VARIANT_BOOL abort);
	/// \brief <em>Finds an item by its \c ItemData property</em>
	///
	/// Searches the list box control for the first item that has the \c ItemData property set to the
	/// specified value.
	///
	/// \param[in] itemData The \c ItemData value for which to search.
	/// \param[in] startAfterItem The item after which the search shall be started. If the bottom of the
	///            control is reached, the search is continued from the top of the control back to the item
	///            specified by \c startAfterItem. If set to \c Empty, the control is searched from top to
	///            bottom.
	/// \param[out] ppFoundItem Receives the found item or \c NULL if no matching item was found.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa FindItemByText, SelectItemByItemData, ListBoxItem::get_ItemData
	virtual HRESULT STDMETHODCALLTYPE FindItemByItemData(LONG itemData, VARIANT startAfterItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IListBoxItem** ppFoundItem = NULL);
	/// \brief <em>Finds an item by its \c ItemData property</em>
	///
	/// Searches the list box control for the first item that has the \c ItemData property set to the
	/// specified value.
	///
	/// \param[in] searchString The string for which to search.
	/// \param[in] exactMatch If \c VARIANT_TRUE, only exact matches are returned; otherwise any item that
	///            starts with the specified string may be returned.
	/// \param[in] startAfterItem The item after which the search shall be started. If the bottom of the
	///            control is reached, the search is continued from the top of the control back to the item
	///            specified by \c startAfterItem. If set to \c Empty, the control is searched from top to
	///            bottom.
	/// \param[out] ppFoundItem Receives the found item or \c NULL if no matching item was found.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The search is not case sensitive.
	///
	/// \sa FindItemByItemData, SelectItemByText, ListBoxItem::get_Text
	virtual HRESULT STDMETHODCALLTYPE FindItemByText(BSTR searchString, VARIANT_BOOL exactMatch = VARIANT_TRUE, VARIANT startAfterItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IListBoxItem** ppFoundItem = NULL);
	/// \brief <em>Finishes a pending drop operation</em>
	///
	/// During a drag'n'drop operation the drag image is displayed until the \c OLEDragDrop event has been
	/// handled. This order is intended by Microsoft Windows. However, if a message box is displayed from
	/// within the \c OLEDragDrop event, or the drop operation cannot be performed asynchronously and takes
	/// a long time, it may be desirable to remove the drag image earlier.\n
	/// This method will break the intended order and finish the drag'n'drop operation (including removal
	/// of the drag image) immediately.
	///
	/// \remarks This method will fail if not called from the \c OLEDragDrop event handler or if no drag
	///          images are used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_OLEDragDrop, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE FinishOLEDragDrop(void);
	/// \brief <em>Proposes a position for the control's insertion mark</em>
	///
	/// Retrieves the insertion mark position that is closest to the specified point.
	///
	/// \param[in] x The x-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the control's upper-left corner.
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified item. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppListItem Receives the \c IListBoxItem implementation of the item, at which
	///             the insertion mark should be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, GetInsertMarkPosition, CBLCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, GetInsertMarkPosition, CBLCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, IListBoxItem** ppListItem);
	/// \brief <em>Retrieves the position of the control's insertion mark</em>
	///
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified item. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppListItem Receives the \c IListBoxItem implementation of the item, at which the
	///             insertion mark is displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, GetClosestInsertMarkPosition, GetInsertMarkRectangle,
	///       CBLCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, GetClosestInsertMarkPosition, GetInsertMarkRectangle,
	///       CBLCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, IListBoxItem** ppListItem);
	/// \brief <em>Retrieves the bounding rectangle of the control's insertion mark</em>
	///
	/// \param[out] pXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///             relative to the control's upper-left corner.
	/// \param[out] pYTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///             relative to the control's upper-left corner.
	/// \param[out] pXRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///             relative to the control's upper-left corner.
	/// \param[out] pYBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///             relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method fails if the \c InsertMarkStyle property is set to \c imsNative.
	///
	/// \sa get_InsertMarkStyle, GetInsertMarkPosition, SetInsertMarkPosition
	virtual HRESULT STDMETHODCALLTYPE GetInsertMarkRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pHitTestDetails Receives a value specifying the exact part of the control the specified
	///             point lies in. Some of the values defined by the \c HitTestConstants enumeration are
	///             valid.
	/// \param[out] ppHitItem Receives the "hit" item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::HitTestConstants, HitTest
	/// \else
	///   \sa CBLCtlsLibA::HitTestConstants, HitTest
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IListBoxItem** ppHitItem);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Enters OLE drag'n'drop mode</em>
	///
	/// \param[in] pDataObject A pointer to the \c IDataObject implementation to use during OLE
	///            drag'n'drop. If not specified, the control's own implementation is used.
	/// \param[in] supportedEffects A bit field defining all drop effects the client wants to support.
	///            Any combination of the values defined by the \c OLEDropEffectConstants enumeration
	///            (except \c odeScroll) is valid.
	/// \param[in] hWndToAskForDragImage The handle of the window, that is awaiting the
	///            \c DI_GETDRAGIMAGE message to specify the drag image to use. If -1, the method
	///            creates the drag image itself. If \c SupportOLEDragImages is set to \c VARIANT_FALSE,
	///            no drag image is used.
	/// \param[in] pDraggedItems The dragged items collection object's \c IListBoxItemContainer
	///            implementation. It is used to generate the drag image and is ignored if
	///            \c hWndToAskForDragImage is not -1.
	/// \param[in] itemCountToDisplay The number to display in the item count label of Aero drag images.
	///            If set to 0 or 1, no item count label is displayed. If set to -1, the number of items
	///            contained in the \c pDraggedItems collection is displayed in the item count label. If
	///            set to any value larger than 1, this value is displayed in the item count label.
	/// \param[out] pPerformedEffects The performed drop effect. Any of the values defined by the
	///             \c OLEDropEffectConstants enumeration (except \c odeScroll) is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa BeginDrag, Raise_ItemBeginDrag, Raise_ItemBeginRDrag, Raise_OLEStartDrag,
	///       get_SupportOLEDragImages, get_OLEDragImageStyle, CBLCtlsLibU::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \else
	///   \sa BeginDrag, Raise_ItemBeginDrag, Raise_ItemBeginRDrag, Raise_OLEStartDrag,
	///       get_SupportOLEDragImages, get_OLEDragImageStyle, CBLCtlsLibA::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE OLEDrag(LONG* pDataObject = NULL, OLEDropEffectConstants supportedEffects = odeCopyOrMove, OLE_HANDLE hWndToAskForDragImage = -1, IListBoxItemContainer* pDraggedItems = NULL, LONG itemCountToDisplay = -1, OLEDropEffectConstants* pPerformedEffects = NULL);
	/// \brief <em>Prepares the control for insertion of many items</em>
	///
	/// Prepares the control for the insertion of many items by reserving memory in advance. It is not
	/// necessary to call this method, but doing so can improve performance when inserting many (> 100)
	/// items at once.
	///
	/// \param[in] numberOfItems The number of items for which to reserve memory in advance.
	/// \param[in] averageStringWidth The average length of the item texts.
	/// \param[out] pSucceeded Will be set to the number of items for which memory has been allocated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ListBoxItems::Add
	virtual HRESULT STDMETHODCALLTYPE PrepareForItemInsertions(LONG numberOfItems, LONG averageStringWidth, LONG* pAllocatedItems);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Finds and selects an item by its \c ItemData property</em>
	///
	/// Searches the list box control for the first item that has the \c ItemData property set to the
	/// specified value. If a matching item is found, it is made the selected item as specified by the
	/// \c SelectedItem property.
	///
	/// \param[in] itemData The \c ItemData value for which to search.
	/// \param[in] startAfterItem The item after which the search shall be started. If the bottom of the
	///            control is reached, the search is continued from the top of the control back to the item
	///            specified by \c startAfterItem. If set to \c Empty, the control is searched from top to
	///            bottom.
	/// \param[out] ppFoundItem Receives the found item or \c NULL if no matching item was found.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SelectItemByText, FindItemByItemData, ListBoxItem::get_ItemData, putref_SelectedItem
	virtual HRESULT STDMETHODCALLTYPE SelectItemByItemData(LONG itemData, VARIANT startAfterItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IListBoxItem** ppFoundItem = NULL);
	/// \brief <em>Finds and selects an item by its \c ItemData property</em>
	///
	/// Searches the list box control for the first item that has the \c ItemData property set to the
	/// specified value. If a matching item is found, it is made the selected item as specified by the
	/// \c SelectedItem property.
	///
	/// \param[in] searchString The string for which to search.
	/// \param[in] exactMatch If \c VARIANT_TRUE, only exact matches are returned; otherwise any item that
	///            starts with the specified string may be returned.
	/// \param[in] startAfterItem The item after which the search shall be started. If the bottom of the
	///            control is reached, the search is continued from the top of the control back to the item
	///            specified by \c startAfterItem. If set to \c Empty, the control is searched from top to
	///            bottom.
	/// \param[out] ppFoundItem Receives the found item or \c NULL if no matching item was found.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The search is not case sensitive.
	///
	/// \sa SelectItemByItemData, FindItemByText, ListBoxItem::get_Text, putref_SelectedItem
	virtual HRESULT STDMETHODCALLTYPE SelectItemByText(BSTR searchString, VARIANT_BOOL exactMatch = VARIANT_TRUE, VARIANT startAfterItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IListBoxItem** ppFoundItem = NULL);
	/// \brief <em>Selects a range of items</em>
	///
	/// Selects the specified range of items.
	///
	/// \param[in] firstItem The first item to select. If \c Empty, the control's first item is used.
	/// \param[in] lastItem The last item to select. If \c Empty, the control's first item is used.
	///
	/// \sa DeselectItems, get_MultiSelect, ListBoxItem::put_Selected
	virtual HRESULT STDMETHODCALLTYPE SelectItems(VARIANT firstItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), VARIANT lastItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR));
	/// \brief <em>Sets the position of the control's insertion mark</em>
	///
	/// \param[in] relativePosition The insertion mark's position relative to the specified item. Any of
	///            the values defined by the \c InsertMarkPositionConstants enumeration is valid.
	/// \param[in] pListItem The \c IListBoxItem implementation of the item, at which the insertion
	///            mark will be displayed. If set to \c NULL, the insertion mark will be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetInsertMarkPosition, GetClosestInsertMarkPosition, get_InsertMarkColor,
	///       get_InsertMarkStyle, get_RegisterForOLEDragDrop, CBLCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa GetInsertMarkPosition, GetClosestInsertMarkPosition, get_InsertMarkColor,
	///       get_InsertMarkStyle, get_RegisterForOLEDragDrop, CBLCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, IListBoxItem* pListItem);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Property object changes
	///
	//@{
	/// \brief <em>Will be called after a property object was changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnChanged
	HRESULT OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/);
	/// \brief <em>Will be called before a property object is changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that is about to be changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnRequestEdit
	HRESULT OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Called to create the control window</em>
	///
	/// Called to create the control window. This method overrides \c CWindowImpl::Create() and is
	/// used to customize the window styles.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] szWindowName The control's window name.
	/// \param[in] dwStyle The control's window style. Will be ignored.
	/// \param[in] dwExStyle The control's extended window style. Will be ignored.
	/// \param[in] MenuOrID The control's ID.
	/// \param[in] lpCreateParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnCreate, GetStyleBits, GetExStyleBits
	HWND Create(HWND hWndParent, ATL::_U_RECT rect = NULL, LPCTSTR szWindowName = NULL, DWORD dwStyle = 0, DWORD dwExStyle = 0, ATL::_U_MENUorID MenuOrID = 0U, LPVOID lpCreateParam = NULL);
	/// \brief <em>Called to draw the control</em>
	///
	/// Called to draw the control. This method overrides \c CComControlBase::OnDraw() and is used to prevent
	/// the "ATL 9.0" drawing in user mode and replace it in design mode.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/056hw3hs.aspx">CComControlBase::OnDraw</a>
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Called after receiving the last message (typically \c WM_NCDESTROY)</em>
	///
	/// \param[in] hWnd The window being destroyed.
	///
	/// \sa OnCreate, OnDestroy
	void OnFinalMessage(HWND /*hWnd*/);
	/// \brief <em>Informs an embedded object of its display location within its container</em>
	///
	/// \param[in] pClientSite The \c IOleClientSite implementation of the container application's
	///            client side.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms684013.aspx">IOleObject::SetClientSite</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite);
	/// \brief <em>Notifies the control when the container's document window is activated or deactivated</em>
	///
	/// ATL's implementation of \c OnDocWindowActivate calls \c IOleInPlaceObject_UIDeactivate if the control
	/// is deactivated. This causes a bug in MDI apps. If the control sits on a \c MDI child window and has
	/// the focus and the title bar of another top-level window (not the MDI parent window) of the same
	/// process is clicked, the focus is moved from the ATL based ActiveX control to the next control on the
	/// MDI child before it is moved to the other top-level window that was clicked. If the focus is set back
	/// to the MDI child, the ATL based control no longer has the focus.
	///
	/// \param[in] fActivate If \c TRUE, the document window is activated; otherwise deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/0kz79wfc.aspx">IOleInPlaceActiveObjectImpl::OnDocWindowActivate</a>
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL /*fActivate*/);

	/// \brief <em>A keyboard or mouse message needs to be translated</em>
	///
	/// The control's container calls this method if it receives a keyboard or mouse message. It gives
	/// us the chance to customize keystroke translation (i. e. to react to them in a non-default way).
	/// This method overrides \c CComControlBase::PreTranslateAccelerator.
	///
	/// \param[in] pMessage A \c MSG structure containing details about the received window message.
	/// \param[out] hReturnValue A reference parameter of type \c HRESULT which will be set to \c S_OK,
	///             if the message was translated, and to \c S_FALSE otherwise.
	///
	/// \return \c FALSE if the object's container should translate the message; otherwise \c TRUE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hxa56938.aspx">CComControlBase::PreTranslateAccelerator</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646373.aspx">TranslateAccelerator</a>
	BOOL PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue);

	//////////////////////////////////////////////////////////////////////
	/// \name Drag image creation
	///
	//@{
	/// \brief <em>Creates a legacy drag image for the specified item</em>
	///
	/// Creates a drag image for the specified item in the style of Windows versions prior to Vista. The
	/// drag image is added to an image list which is returned.
	///
	/// \param[in] itemIndex The item for which to create the drag image.
	/// \param[out] pUpperLeftPoint Receives the coordinates (in pixels) of the drag image's upper-left
	///             corner relative to the control's upper-left corner.
	/// \param[out] pBoundingRectangle Receives the drag image's bounding rectangle (in pixels) relative to
	///             the control's upper-left corner.
	///
	/// \return An image list containing the drag image.
	///
	/// \remarks The caller is responsible for destroying the image list.
	///
	/// \sa CreateLegacyOLEDragImage, ListBoxItemContainer::CreateDragImage
	HIMAGELIST CreateLegacyDragImage(int itemIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle);
	/// \brief <em>Creates a legacy OLE drag image for the specified items</em>
	///
	/// Creates an OLE drag image for the specified items in the style of Windows versions prior to Vista.
	///
	/// \param[in] pItems A \c IListBoxItemContainer object wrapping the items for which to create the drag
	///            image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateVistaOLEDragImage, CreateLegacyDragImage, ListBoxItemContainer,
	///     <a href="https://msdn.microsoft.com/En-US/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateLegacyOLEDragImage(IListBoxItemContainer* pItems, __in LPSHDRAGIMAGE pDragImage);
	/// \brief <em>Creates a Vista-like OLE drag image for the specified items</em>
	///
	/// Creates an OLE drag image for the specified items in the style of Windows Vista and newer.
	///
	/// \param[in] pItems A \c IListBoxItemContainer object wrapping the items for which to create the drag
	///            image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateLegacyOLEDragImage, CreateLegacyDragImage, ListBoxItemContainer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateVistaOLEDragImage(IListBoxItemContainer* pItems, __in LPSHDRAGIMAGE pDragImage);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropTarget
	///
	//@{
	/// \brief <em>Indicates whether a drop can be accepted, and, if so, the effect of the drop</em>
	///
	/// This method is called by the \c DoDragDrop function to determine the target's preferred drop
	/// effect the first time the user moves the mouse into the control during OLE drag'n'Drop. The
	/// target communicates the \c DoDragDrop function the drop effect it wants to be used on drop.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the dragged
	///            data.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragOver, DragLeave, Drop, Raise_OLEDragEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Notifies the target that it no longer is the target of the current OLE drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse out of the
	/// control during OLE drag'n'Drop or if the user canceled the operation. The target must release
	/// any references it holds to the data object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, Drop, Raise_OLEDragLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeave(void);
	/// \brief <em>Communicates the current drop effect to the \c DoDragDrop function</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse over the
	/// control during OLE drag'n'Drop. The target communicates the \c DoDragDrop function the drop
	/// effect it wants to be used on drop.
	///
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, the current drop effect (defined by the \c DROPEFFECT
	///                enumeration). On return, this paramter must be set to the drop effect that the
	///                target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragLeave, Drop, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
	virtual HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Incorporates the source data into the target and completes the drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user completes the drag'n'drop
	/// operation. The target must incorporate the dragged data into itself and pass the used drop
	/// effect back to the \c DoDragDrop function. The target must release any references it holds to
	/// the data object.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the data
	///            to transfer.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target finally executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, DragLeave, Raise_OLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSource
	///
	//@{
	/// \brief <em>Notifies the source of the current drop effect during OLE drag'n'drop</em>
	///
	/// This method is called frequently by the \c DoDragDrop function to notify the source of the
	/// last drop effect that the target has chosen. The source should set an appropriate mouse cursor.
	///
	/// \param[in] effect The drop effect chosen by the target. Any of the values defined by the
	///            \c DROPEFFECT enumeration is valid.
	///
	/// \return \c S_OK if the method has set a custom mouse cursor; \c DRAGDROP_S_USEDEFAULTCURSORS to
	///         use default mouse cursors; or an error code otherwise.
	///
	/// \sa QueryContinueDrag, Raise_OLEGiveFeedback,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693723.aspx">IDropSource::GiveFeedback</a>
	virtual HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD effect);
	/// \brief <em>Determines whether a drag'n'drop operation should be continued, canceled or completed</em>
	///
	/// This method is called by the \c DoDragDrop function to determine whether a drag'n'drop
	/// operation should be continued, canceled or completed.
	///
	/// \param[in] pressedEscape Indicates whether the user has pressed the \c ESC key since the
	///            previous call of this method. If \c TRUE, the key has been pressed; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	///
	/// \return \c S_OK if the drag'n'drop operation should continue; \c DRAGDROP_S_DROP if it should
	///         be completed; \c DRAGDROP_S_CANCEL if it should be canceled; or an error code otherwise.
	///
	/// \sa GiveFeedback, Raise_OLEQueryContinueDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690076.aspx">IDropSource::QueryContinueDrag</a>
	virtual HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL pressedEscape, DWORD keyState);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSourceNotify
	///
	//@{
	/// \brief <em>Notifies the source that the user drags the mouse cursor into a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor into a potential drop target window.
	///
	/// \param[in] hWndTarget The potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragLeaveTarget, Raise_OLEDragEnterPotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragEnterTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnterTarget(HWND hWndTarget);
	/// \brief <em>Notifies the source that the user drags the mouse cursor out of a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor out of a potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnterTarget, Raise_OLEDragLeavePotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragLeaveTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeaveTarget(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	/// \param[in] languageID The locale identifier identifying the language in which name should be
	///            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnKeyDown, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to configure the control window and to raise the \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_INPUTLANGCHANGE handler</em>
	///
	/// Will be called after an application's input language has been changed.
	/// We use this handler to update the IME mode of the control.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632629.aspx">WM_INPUTLANGCHANGE</a>
	LRESULT OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KILLFOCUS handler</em>
	///
	/// Will be called after the control lost the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnSetFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646282.aspx">WM_KILLFOCUS</a>
	LRESULT OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the left mouse
	/// button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnMButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnLButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnMButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEACTIVATE handler</em>
	///
	/// Will be called if the control is inactive and the user clicked in its client area.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnSetFocus, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645612.aspx">WM_MOUSEACTIVATE</a>
	LRESULT OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the control's client area for a previously
	/// specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, OnMouseWheel, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnMouseMove, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NCMOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// non-client area.
	/// We use this handler to relay the message to the tooltip control.
	///
	/// \sa ToolTipStatus::RelayToToolTip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645627.aspx">WM_NCMOUSEMOVE</a>
	LRESULT OnNCMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnXButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_*SCROLL handler</em>
	///
	/// Will be called if the control shall be scrolled.
	/// We use this handler to redraw the insertion mark.
	///
	/// \sa DrawInsertionMark,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651283.aspx">WM_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651284.aspx">WM_VSCROLL</a>
	LRESULT OnScroll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called after the control gained the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTINGCHANGE handler</em>
	///
	/// Will be called if a system setting was changed.
	/// We use this handler to update our appearance.
	///
	/// \attention This message is posted to top-level windows only, so actually we'll never receive it.
	///
	/// \sa OnThemeChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms725497.aspx">WM_SETTINGCHANGE</a>
	LRESULT OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SIZE handler</em>
	///
	/// Will be called if the control's size was changed.
	/// We use this handler to update the tooltip control.
	///
	/// \sa OnWindowPosChanged, toolTipStatus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632646.aspx">WM_SIZE</a>
	LRESULT OnSize(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_THEMECHANGED handler</em>
	///
	/// Will be called on themable systems if the theme was changed.
	/// We use this handler to update our appearance.
	///
	/// \sa OnSettingChange, Flags::usingThemes,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632650.aspx">WM_THEMECHANGED</a>
	LRESULT OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called when a timer expires that's associated with the control.
	/// We use this handler for the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control was moved.
	/// We use this handler to resize the control on COM level and to raise the \c ResizedControlWindow
	/// event.
	///
	/// \sa OnSize, Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CHARTOITEM handler</em>
	///
	/// Will be called if the control's parent window is sent a \c WM_CHARTOITEM message in response to a
	/// \c WM_KEYDOWN message.
	/// We use this handler to raise the \c ProcessCharacterInput event.
	///
	/// \sa OnReflectedVKeyToItem, get_DisabledEvents, Raise_ProcessCharacterInput,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761358.aspx">WM_CHARTOITEM</a>
	LRESULT OnReflectedCharToItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_COMPAREITEM handler</em>
	///
	/// Will be called if the control has the \c LBS_SORT style and is owner-drawn and an item is inserted.
	/// The handler of this message must specify the relative order of the two specified items.
	/// We use this handler to raise the \c CompareItems event.
	///
	/// \sa Raise_CompareItems,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775921.aspx">WM_COMPAREITEM</a>
	LRESULT OnReflectedCompareItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CTLCOLORLISTBOX handler</em>
	///
	/// Will be called if the control asks its parent window to configure the specified device context for
	/// drawing the control, i. e. to setup the colors and brushes.
	/// We use this handler for the \c BackColor and \c ForeColor properties.
	///
	/// \sa get_BackColor, get_ForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761360.aspx">WM_CTLCOLORLISTBOX</a>
	LRESULT OnReflectedCtlColorListBox(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DRAWITEM handler</em>
	///
	/// Will be called if the control's parent window is asked to draw a list box item.
	/// We use this handler to raise the \c OwnerDrawItem event.
	///
	/// \sa OnReflectedMeasureItem, get_OwnerDrawItems, Raise_OwnerDrawItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775923.aspx">WM_DRAWITEM</a>
	LRESULT OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MEASUREITEM handler</em>
	///
	/// Will be called if the control's parent window is asked for a list box item's height.
	/// We use this handler to raise the \c MeasureItem event.
	///
	/// \sa OnReflectedDrawItem, get_OwnerDrawItems, Raise_MeasureItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775925.aspx">WM_MEASUREITEM</a>
	LRESULT OnReflectedMeasureItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_VKEYTOITEM handler</em>
	///
	/// Will be called if the control's parent window is sent a \c WM_VKEYTOITEM message in response to a
	/// \c WM_KEYDOWN message.
	/// We use this handler to raise the \c ProcessKeyStroke event.
	///
	/// \sa OnReflectedCharToItem, get_DisabledEvents, Raise_ProcessKeyStroke,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761364.aspx">WM_VKEYTOITEM</a>
	LRESULT OnReflectedVKeyToItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c DI_GETDRAGIMAGE handler</em>
	///
	/// Will be called during OLE drag'n'drop if the control is queried for a drag image.
	///
	/// \sa OLEDrag, CreateLegacyOLEDragImage, CreateVistaOLEDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	LRESULT OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c LB_ADDSTRING handler</em>
	///
	/// Will be called if an item shall be inserted. We use this handler to raise the \c InsertingItem
	/// and \c InsertedItem events.
	///
	/// \sa OnInsertString, Raise_InsertingItem, Raise_InsertedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775181.aspx">LB_ADDSTRING</a>
	LRESULT OnAddString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c LB_DELETESTRING handler</em>
	///
	/// Will be called if an item shall be removed. We use this handler to raise the \c FreeItemData,
	/// \c RemovingItem and \c RemovedItem events.
	///
	/// \sa OnResetContent, Raise_FreeItemData, Raise_RemovingItem, Raise_RemovedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775183.aspx">LB_DELETESTRING</a>
	LRESULT OnDeleteString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c LB_INSERTSTRING handler</em>
	///
	/// Will be called if an item shall be inserted. We use this handler to raise the \c InsertingItem
	/// and \c InsertedItem events.
	///
	/// \sa OnAddString, Raise_InsertingItem, Raise_InsertedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761321.aspx">LB_INSERTSTRING</a>
	LRESULT OnInsertString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c LB_RESETCONTENT handler</em>
	///
	/// Will be called if all items shall be removed. We use this handler to raise the \c FreeItemData,
	/// \c RemovingItem and \c RemovedItem events.
	///
	/// \sa OnDeleteString, Raise_FreeItemData, Raise_RemovingItem, Raise_RemovedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761325.aspx">LB_RESETCONTENT</a>
	LRESULT OnResetContent(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c LB_SELECTSTRING, \c LB_SELITEMRANGE, \c LB_SELITEMRANGEEX and \c LB_SETSEL handler</em>
	///
	/// Will be called if the control is in multi-selection mode and the selection shall be changed.
	/// We use this handler to emulate the \ LBN_SELCHANGE command.
	///
	/// \sa get_MultiSelect, OnSetCaretIndex,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761327.aspx">LB_SELECTSTRING</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761329.aspx">LB_SELITEMRANGE</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761331.aspx">LB_SELITEMRANGEEX</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761352.aspx">LB_SETSEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775161.aspx">LBN_SELCHANGE</a>
	LRESULT OnSelectionChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c LB_SETCARETINDEX handler</em>
	///
	/// Will be called if the control's caret item shall be changed.
	/// We use this handler to raise the \c CaretChanged event.
	///
	/// \sa get_CaretItem, OnSelectionChange, Raise_CaretChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761336.aspx">LB_SETCARETINDEX</a>
	LRESULT OnSetCaretIndex(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c LB_SETCOLUMNWIDTH handler</em>
	///
	/// Will be called if the control is in multi-column mode and the column width shall be changed.
	/// We use this handler to synchronize our settings.
	///
	/// \sa get_ColumnWidth,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761338.aspx">LB_SETCOLUMNWIDTH</a>
	LRESULT OnSetColumnWidth(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c LB_SETTABSTOPS handler</em>
	///
	/// Will be called if the control's tab stops shall be changed.
	/// We use this handler to synchronize our settings.
	///
	/// \sa get_TabStops, get_TabWidth,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761354.aspx">LB_SETTABSTOPS</a>
	LRESULT OnSetTabStops(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c TTN_GETDISPINFOA handler</em>
	///
	/// Will be called if the control is queried by the tooltip control for the tooltip text to display.
	/// We use this handler to provide the appropriate tooltip text.
	///
	/// \sa OnToolTipGetDispInfoNotificationW,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760269.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOW handler</em>
	///
	/// Will be called if the control is queried by the tooltip control for the tooltip text to display.
	/// We use this handler to provide the appropriate tooltip text.
	///
	/// \sa OnToolTipGetDispInfoNotificationA,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760269.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_SHOW handler</em>
	///
	/// Will be called if the control is notified, that the tooltip control is about to be displayed.
	/// We use this handler to place the tooltip directly over the item label if necessary.
	///
	/// \sa OnToolTipGetDispInfoNotificationW,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760314.aspx">TTN_SHOW</a>
	LRESULT OnToolTipShowNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c LBN_ERRSPACE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control couldn't allocate enough
	/// memory to meet a specific request.
	/// We use this handler to raise the \c OutOfMemory event.
	///
	/// \sa Raise_OutOfMemory,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775155.aspx">LBN_ERRSPACE</a>
	LRESULT OnReflectedErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	// \brief <em>\c LBN_SELCANCEL handler</em>
	//
	// Will be called if the control's parent window is notified, that the selection in the control has
	// changed.
	// We use this handler to raise the \c CaretChanged and \c SelectionChanged events.
	//
	// \sa Raise_CaretChanged, Raise_SelectionChanged,
	//     <a href="https://msdn.microsoft.com/en-us/library/bb775159.aspx">LBN_SELCANCEL</a>
	//LRESULT OnReflectedSelCancel(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c LBN_SELCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the selection in the control has
	/// changed.
	/// We use this handler to raise the \c CaretChanged and \c SelectionChanged events.
	///
	/// \sa Raise_CaretChanged, Raise_SelectionChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775161.aspx">LBN_SELCHANGE</a>
	LRESULT OnReflectedSelChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c LBN_SETFOCUS handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has gained the keyboard
	/// focus.
	/// We use this handler to initialize IME.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775164.aspx">LBN_SETFOCUS</a>
	LRESULT OnReflectedSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c AbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_AbortedDrag, CBLCtlsLibU::_IListBoxEvents::AbortedDrag, Raise_Drop,
	///       EndDrag
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_AbortedDrag, CBLCtlsLibA::_IListBoxEvents::AbortedDrag, Raise_Drop,
	///       EndDrag
	/// \endif
	inline HRESULT Raise_AbortedDrag(void);
	/// \brief <em>Raises the \c CaretChanged event</em>
	///
	/// \param[in] previousCaretItem The previous caret item.
	/// \param[in] newCaretItem The new caret item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_CaretChanged, CBLCtlsLibU::_IListBoxEvents::CaretChanged,
	///       Raise_SelectionChanged, get_CaretItem
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_CaretChanged, CBLCtlsLibA::_IListBoxEvents::CaretChanged,
	///       Raise_SelectionChanged, get_CaretItem
	/// \endif
	inline HRESULT Raise_CaretChanged(int previousCaretItem, int newCaretItem);
	/// \brief <em>Raises the \c Click event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_Click, CBLCtlsLibU::_IListBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_Click, CBLCtlsLibA::_IListBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CompareItems event</em>
	///
	/// \param[in] pFirstItem The first item to compare.
	/// \param[in] pSecondItem The second item to compare.
	/// \param[in] locale The identifier of the locale to use for comparison.
	/// \param[out] pResult Receives one of the values defined by the \c CompareResultConstants enumeration.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The items' indexes may be -1. For instance, this will happen if one of the items is
	///          currently being inserted. In this case all properties of \c IListBoxItem except \c Index and
	///          \c ItemData will fail and the \c ItemData property is read-only.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_CompareItems, CBLCtlsLibU::_IListBoxEvents::CompareItems,
	///       CBLCtlsLibU::CompareResultConstants, get_OwnerDrawItems, get_Sorted, get_Locale,
	///       ListBoxItem::get_Index, ListBoxItem::get_ItemData
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_CompareItems, CBLCtlsLibA::_IListBoxEvents::CompareItems,
	///       CBLCtlsLibA::CompareResultConstants, get_OwnerDrawItems, get_Sorted, get_Locale,
	///       ListBoxItem::get_Index, ListBoxItem::get_ItemData
	/// \endif
	inline HRESULT Raise_CompareItems(IListBoxItem* pFirstItem, IListBoxItem* pSecondItem, LONG locale, CompareResultConstants* pResult);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_ContextMenu, CBLCtlsLibU::_IListBoxEvents::ContextMenu,
	///       Raise_RClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ContextMenu, CBLCtlsLibA::_IListBoxEvents::ContextMenu,
	///       Raise_RClick
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_DblClick, CBLCtlsLibU::_IListBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_DblClick, CBLCtlsLibA::_IListBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_DestroyedControlWindow,
	///       CBLCtlsLibU::_IListBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow, get_hWnd
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_DestroyedControlWindow,
	///       CBLCtlsLibA::_IListBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow, get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(HWND hWnd);
	/// \brief <em>Raises the \c DragMouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_DragMouseMove, CBLCtlsLibU::_IListBoxEvents::DragMouseMove,
	///       Raise_MouseMove, Raise_OLEDragMouseMove, BeginDrag
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_DragMouseMove, CBLCtlsLibA::_IListBoxEvents::DragMouseMove,
	///       Raise_MouseMove, Raise_OLEDragMouseMove, BeginDrag
	/// \endif
	inline HRESULT Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c Drop event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_Drop, CBLCtlsLibU::_IListBoxEvents::Drop, Raise_AbortedDrag, EndDrag
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_Drop, CBLCtlsLibA::_IListBoxEvents::Drop, Raise_AbortedDrag, EndDrag
	/// \endif
	inline HRESULT Raise_Drop(void);
	/// \brief <em>Raises the \c FreeItemData event</em>
	///
	/// \param[in] pListItem The item whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_FreeItemData, CBLCtlsLibU::_IListBoxEvents::FreeItemData,
	///       Raise_RemovingItem, Raise_RemovedItem, ListBoxItem::put_ItemData
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_FreeItemData, CBLCtlsLibA::_IListBoxEvents::FreeItemData,
	///       Raise_RemovingItem, Raise_RemovedItem, ListBoxItem::put_ItemData
	/// \endif
	inline HRESULT Raise_FreeItemData(IListBoxItem* pListItem);
	/// \brief <em>Raises the \c InsertedItem event</em>
	///
	/// \param[in] pListItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_InsertedItem, CBLCtlsLibU::_IListBoxEvents::InsertedItem,
	///       Raise_InsertingItem, Raise_RemovedItem
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_InsertedItem, CBLCtlsLibA::_IListBoxEvents::InsertedItem,
	///       Raise_InsertingItem, Raise_RemovedItem
	/// \endif
	inline HRESULT Raise_InsertedItem(IListBoxItem* pListItem);
	/// \brief <em>Raises the \c InsertingItem event</em>
	///
	/// \param[in] pListItem The item being inserted. If the list box is sorted, this object's \c Index
	///            property may be wrong.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort insertion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_InsertingItem, CBLCtlsLibU::_IListBoxEvents::InsertingItem,
	///       Raise_InsertedItem, Raise_RemovingItem
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_InsertingItem, CBLCtlsLibA::_IListBoxEvents::InsertingItem,
	///       Raise_InsertedItem, Raise_RemovingItem
	/// \endif
	inline HRESULT Raise_InsertingItem(IVirtualListBoxItem* pListItem, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ItemBeginDrag event</em>
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
	///   \sa Proxy_IListBoxEvents::Fire_ItemBeginDrag, CBLCtlsLibU::_IListBoxEvents::ItemBeginDrag,
	///       BeginDrag, OLEDrag, get_AllowDragDrop, Raise_ItemBeginRDrag, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ItemBeginDrag, CBLCtlsLibA::_IListBoxEvents::ItemBeginDrag,
	///       BeginDrag, OLEDrag, get_AllowDragDrop, Raise_ItemBeginRDrag, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemBeginDrag(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c ItemBeginRDrag event</em>
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
	///   \sa Proxy_IListBoxEvents::Fire_ItemBeginRDrag, CBLCtlsLibU::_IListBoxEvents::ItemBeginRDrag,
	///       BeginDrag, OLEDrag, get_AllowDragDrop, Raise_ItemBeginDrag, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ItemBeginRDrag, CBLCtlsLibA::_IListBoxEvents::ItemBeginRDrag,
	///       BeginDrag, OLEDrag, get_AllowDragDrop, Raise_ItemBeginDrag, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemBeginRDrag(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c ItemGetDisplayInfo event</em>
	///
	/// \param[in] pListItem The item whose properties are requested.
	/// \param[in] requestedInfo Specifies which properties' values are required. Some combinations of
	///            the values defined by the \c RequestedInfoConstants enumeration are valid.
	/// \param[out] pIconIndex The zero-based index of the item's icon. If the \c requestedInfo parameter
	///             doesn't include \c riIconIndex, the caller should ignore this value.
	/// \param[out] pOverlayIndex The zero-based index of the item's overlay icon. If the \c requestedInfo
	///             parameter doesn't include \c riOverlayIndex, the caller should ignore this value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_ItemGetDisplayInfo,
	///       CBLCtlsLibU::_IListBoxEvents::ItemGetDisplayInfo, put_hImageList,
	///       CBLCtlsLibU::RequestedInfoConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ItemGetDisplayInfo,
	///       CBLCtlsLibA::_IListBoxEvents::ItemGetDisplayInfo, put_hImageList,
	///       CBLCtlsLibA::RequestedInfoConstants
	/// \endif
	inline HRESULT Raise_ItemGetDisplayInfo(IListBoxItem* pListItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pOverlayIndex);
	/// \brief <em>Raises the \c ItemGetInfoTipText event</em>
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
	///   \sa Proxy_IListBoxEvents::Fire_ItemGetInfoTipText,
	///       CBLCtlsLibU::_IListBoxEvents::ItemGetInfoTipText, put_ToolTips
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ItemGetInfoTipText,
	///       CBLCtlsLibA::_IListBoxEvents::ItemGetInfoTipText, put_ToolTips
	/// \endif
	inline HRESULT Raise_ItemGetInfoTipText(IListBoxItem* pListItem, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip);
	/// \brief <em>Raises the \c ItemMouseEnter event</em>
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
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_ItemMouseEnter, CBLCtlsLibU::_IListBoxEvents::ItemMouseEnter,
	///       Raise_ItemMouseLeave, Raise_MouseMove, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ItemMouseEnter, CBLCtlsLibA::_IListBoxEvents::ItemMouseEnter,
	///       Raise_ItemMouseLeave, Raise_MouseMove, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemMouseEnter(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c ItemMouseLeave event</em>
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
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_ItemMouseLeave, CBLCtlsLibU::_IListBoxEvents::ItemMouseLeave,
	///       Raise_ItemMouseEnter, Raise_MouseMove, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ItemMouseLeave, CBLCtlsLibA::_IListBoxEvents::ItemMouseLeave,
	///       Raise_ItemMouseEnter, Raise_MouseMove, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemMouseLeave(IListBoxItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_KeyDown, CBLCtlsLibU::_IListBoxEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress, Raise_ProcessKeyStroke
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_KeyDown, CBLCtlsLibA::_IListBoxEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress, Raise_ProcessKeyStroke
	/// \endif
	inline HRESULT Raise_KeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_KeyPress, CBLCtlsLibU::_IListBoxEvents::KeyPress, Raise_KeyDown,
	///       Raise_KeyUp, Raise_ProcessCharacterInput
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_KeyPress, CBLCtlsLibA::_IListBoxEvents::KeyPress, Raise_KeyDown,
	///       Raise_KeyUp, Raise_ProcessCharacterInput
	/// \endif
	inline HRESULT Raise_KeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_KeyUp, CBLCtlsLibU::_IListBoxEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress, Raise_ProcessKeyStroke
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_KeyUp, CBLCtlsLibA::_IListBoxEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress, Raise_ProcessKeyStroke
	/// \endif
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c MClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MClick, CBLCtlsLibU::_IListBoxEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MClick, CBLCtlsLibA::_IListBoxEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MDblClick, CBLCtlsLibU::_IListBoxEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MDblClick, CBLCtlsLibA::_IListBoxEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MeasureItem event</em>
	///
	/// \param[in] pListItem The item for which the size is required. If the list box is empty, this
	///            parameter will be \c NULL.
	/// \param[out] pItemHeight Must be set to the item's height in pixels by the client app.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MeasureItem, CBLCtlsLibU::_IListBoxEvents::MeasureItem,
	///       Raise_OwnerDrawItem, ListBox::put_ItemHeight, ListBoxItem::put_Height,
	///       ListBox::get_OwnerDrawItems
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MeasureItem, CBLCtlsLibA::_IListBoxEvents::MeasureItem,
	///       Raise_OwnerDrawItem, ListBox::put_ItemHeight, ListBoxItem::put_Height,
	///       ListBox::get_OwnerDrawItems
	/// \endif
	inline HRESULT Raise_MeasureItem(IListBoxItem* pListItem, OLE_YSIZE_PIXELS* pItemHeight);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MouseDown, CBLCtlsLibU::_IListBoxEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MouseDown, CBLCtlsLibA::_IListBoxEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MouseEnter, CBLCtlsLibU::_IListBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_ItemMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MouseEnter, CBLCtlsLibA::_IListBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_ItemMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MouseHover, CBLCtlsLibU::_IListBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MouseHover, CBLCtlsLibA::_IListBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \endif
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MouseLeave, CBLCtlsLibU::_IListBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_ItemMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MouseLeave, CBLCtlsLibA::_IListBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_ItemMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MouseMove, CBLCtlsLibU::_IListBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MouseMove, CBLCtlsLibA::_IListBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MouseUp, CBLCtlsLibU::_IListBoxEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MouseUp, CBLCtlsLibA::_IListBoxEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseWheel event</em>
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_MouseWheel, CBLCtlsLibU::_IListBoxEvents::MouseWheel,
	///       Raise_MouseMove, CBLCtlsLibU::ExtendedMouseButtonConstants, CBLCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_MouseWheel, CBLCtlsLibA::_IListBoxEvents::MouseWheel,
	///       Raise_MouseMove, CBLCtlsLibA::ExtendedMouseButtonConstants, CBLCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLECompleteDrag, CBLCtlsLibU::_IListBoxEvents::OLECompleteDrag,
	///       Raise_OLEStartDrag, SourceOLEDataObject::GetData, OLEDrag
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLECompleteDrag, CBLCtlsLibA::_IListBoxEvents::OLECompleteDrag,
	///       Raise_OLEStartDrag, SourceOLEDataObject::GetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect);
	/// \brief <em>Raises the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client finally executed.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragDrop, CBLCtlsLibU::_IListBoxEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragDrop, CBLCtlsLibA::_IListBoxEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data. If \c NULL, the cached data object is used. We use this when
	///            we call this method from other places than \c DragEnter.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragEnter, CBLCtlsLibU::_IListBoxEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragEnter, CBLCtlsLibA::_IListBoxEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The potential drop target window's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragEnterPotentialTarget,
	///       CBLCtlsLibU::_IListBoxEvents::OLEDragEnterPotentialTarget, Raise_OLEDragLeavePotentialTarget,
	///       OLEDrag
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragEnterPotentialTarget,
	///       CBLCtlsLibA::_IListBoxEvents::OLEDragEnterPotentialTarget, Raise_OLEDragLeavePotentialTarget,
	///       OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragLeave, CBLCtlsLibU::_IListBoxEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragLeave, CBLCtlsLibA::_IListBoxEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(void);
	/// \brief <em>Raises the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragLeavePotentialTarget,
	///       CBLCtlsLibU::_IListBoxEvents::OLEDragLeavePotentialTarget, Raise_OLEDragEnterPotentialTarget,
	///       OLEDrag
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragLeavePotentialTarget,
	///       CBLCtlsLibA::_IListBoxEvents::OLEDragLeavePotentialTarget, Raise_OLEDragEnterPotentialTarget,
	///       OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragLeavePotentialTarget(void);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragMouseMove, CBLCtlsLibU::_IListBoxEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEDragMouseMove, CBLCtlsLibA::_IListBoxEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c DROPEFFECT enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEGiveFeedback, CBLCtlsLibU::_IListBoxEvents::OLEGiveFeedback,
	///       Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEGiveFeedback, CBLCtlsLibA::_IListBoxEvents::OLEGiveFeedback,
	///       Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors);
	/// \brief <em>Raises the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c TRUE, the user has pressed the \c ESC key since the last time
	///            this event was raised; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue (\c S_OK), cancel
	///                (\c DRAGDROP_S_CANCEL) or complete (\c DRAGDROP_S_DROP) the drag'n'drop operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEQueryContinueDrag,
	///       CBLCtlsLibU::_IListBoxEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEQueryContinueDrag,
	///       CBLCtlsLibA::_IListBoxEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \endif
	inline HRESULT Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith);
	/// \brief <em>Raises the \c OLEReceivedNewData event</em>
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
	///   \sa Proxy_IListBoxEvents::Fire_OLEReceivedNewData,
	///       CBLCtlsLibU::_IListBoxEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEReceivedNewData,
	///       CBLCtlsLibA::_IListBoxEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLESetData event</em>
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
	///   \sa Proxy_IListBoxEvents::Fire_OLESetData, CBLCtlsLibU::_IListBoxEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLESetData, CBLCtlsLibA::_IListBoxEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OLEStartDrag, CBLCtlsLibU::_IListBoxEvents::OLEStartDrag,
	///       Raise_OLESetData, Raise_OLECompleteDrag, SourceOLEDataObject::SetData, OLEDrag
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OLEStartDrag, CBLCtlsLibA::_IListBoxEvents::OLEStartDrag,
	///       Raise_OLESetData, Raise_OLECompleteDrag, SourceOLEDataObject::SetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEStartDrag(IOLEDataObject* pData);
	/// \brief <em>Raises the \c OutOfMemory event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_OutOfMemory, CBLCtlsLibU::_IListBoxEvents::OutOfMemory
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OutOfMemory, CBLCtlsLibA::_IListBoxEvents::OutOfMemory
	/// \endif
	inline HRESULT Raise_OutOfMemory(void);
	/// \brief <em>Raises the \c OwnerDrawItem event</em>
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
	///   \sa Proxy_IListBoxEvents::Fire_OwnerDrawItem, CBLCtlsLibU::_IListBoxEvents::OwnerDrawItem,
	///       Raise_MeasureItem, ListBox::put_ItemHeight, ListBoxItem::put_Height,
	///       ListBox::get_OwnerDrawItems, CBLCtlsLibU::RECTANGLE, CBLCtlsLibU::OwnerDrawActionConstants,
	///       CBLCtlsLibU::OwnerDrawItemStateConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_OwnerDrawItem, CBLCtlsLibA::_IListBoxEvents::OwnerDrawItem,
	///       Raise_MeasureItem, ListBox::put_ItemHeight, ListBoxItem::put_Height,
	///       ListBox::get_OwnerDrawItems, CBLCtlsLibA::RECTANGLE, CBLCtlsLibA::OwnerDrawActionConstants,
	///       CBLCtlsLibA::OwnerDrawItemStateConstants
	/// \endif
	inline HRESULT Raise_OwnerDrawItem(IListBoxItem* pListItem, OwnerDrawActionConstants requiredAction, OwnerDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle);
	/// \brief <em>Raises the \c ProcessCharacterInput event</em>
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
	/// \remarks This event is raised only for owner-drawn controls with \c HasStrings being set to
	///          \c VARIANT_FALSE.\n
	///          This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_ProcessCharacterInput,
	///       CBLCtlsLibU::_IListBoxEvents::ProcessCharacterInput, Raise_ProcessKeyStroke, Raise_KeyDown,
	///       Raise_KeyPress, Raise_KeyUp, get_HasStrings
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ProcessCharacterInput,
	///       CBLCtlsLibA::_IListBoxEvents::ProcessCharacterInput, Raise_ProcessKeyStroke, Raise_KeyDown,
	///       Raise_KeyPress, Raise_KeyUp, get_HasStrings
	/// \endif
	inline HRESULT Raise_ProcessCharacterInput(SHORT keyAscii, SHORT shift, LONG* pItemToSelect);
	/// \brief <em>Raises the \c ProcessKeyStroke event</em>
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
	/// \remarks Current versions of Windows don't seem to handle a return value of -2 correctly, if the
	///          pressed key is a character.\n
	///          This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_ProcessKeyStroke, CBLCtlsLibU::_IListBoxEvents::ProcessKeyStroke,
	///       Raise_ProcessCharacterInput, Raise_KeyDown, Raise_KeyPress, Raise_KeyUp
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ProcessKeyStroke, CBLCtlsLibA::_IListBoxEvents::ProcessKeyStroke,
	///       Raise_ProcessCharacterInput, Raise_KeyDown, Raise_KeyPress, Raise_KeyUp
	/// \endif
	inline HRESULT Raise_ProcessKeyStroke(SHORT keyCode, SHORT shift, LONG* pItemToSelect);
	/// \brief <em>Raises the \c RClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_RClick, CBLCtlsLibU::_IListBoxEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_RClick, CBLCtlsLibA::_IListBoxEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_RDblClick, CBLCtlsLibU::_IListBoxEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_RDblClick, CBLCtlsLibA::_IListBoxEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_RecreatedControlWindow,
	///       CBLCtlsLibU::_IListBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow, get_hWnd
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_RecreatedControlWindow,
	///       CBLCtlsLibA::_IListBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow, get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(HWND hWnd);
	/// \brief <em>Raises the \c RemovedItem event</em>
	///
	/// \param[in] pListItem The removed item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_RemovedItem, CBLCtlsLibU::_IListBoxEvents::RemovedItem,
	///       Raise_RemovingItem, Raise_InsertedItem
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_RemovedItem, CBLCtlsLibA::_IListBoxEvents::RemovedItem,
	///       Raise_RemovingItem, Raise_InsertedItem
	/// \endif
	inline HRESULT Raise_RemovedItem(IVirtualListBoxItem* pListItem);
	/// \brief <em>Raises the \c RemovingItem event</em>
	///
	/// \param[in] pListItem The item being removed.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_RemovingItem, CBLCtlsLibU::_IListBoxEvents::RemovingItem,
	///       Raise_RemovedItem, Raise_InsertingItem
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_RemovingItem, CBLCtlsLibA::_IListBoxEvents::RemovingItem,
	///       Raise_RemovedItem, Raise_InsertingItem
	/// \endif
	inline HRESULT Raise_RemovingItem(IListBoxItem* pListItem, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_ResizedControlWindow,
	///       CBLCtlsLibU::_IListBoxEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_ResizedControlWindow,
	///       CBLCtlsLibA::_IListBoxEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c SelectionChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MultiSelect, Proxy_IListBoxEvents::Fire_SelectionChanged,
	///       CBLCtlsLibU::_IListBoxEvents::SelectionChanged, Raise_CaretChanged,
	///       ListBoxItem::put_Selected
	/// \else
	///   \sa put_MultiSelect, Proxy_IListBoxEvents::Fire_SelectionChanged,
	///       CBLCtlsLibA::_IListBoxEvents::SelectionChanged, Raise_CaretChanged,
	///       ListBoxItem::put_Selected
	/// \endif
	inline HRESULT Raise_SelectionChanged(void);
	/// \brief <em>Raises the \c XClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_XClick, CBLCtlsLibU::_IListBoxEvents::XClick, Raise_XDblClick,
	///       Raise_Click, Raise_MClick, Raise_RClick, CBLCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_XClick, CBLCtlsLibA::_IListBoxEvents::XClick, Raise_XDblClick,
	///       Raise_Click, Raise_MClick, Raise_RClick, CBLCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IListBoxEvents::Fire_XDblClick, CBLCtlsLibU::_IListBoxEvents::XDblClick, Raise_XClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_RDblClick, CBLCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IListBoxEvents::Fire_XDblClick, CBLCtlsLibA::_IListBoxEvents::XDblClick, Raise_XClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_RDblClick, CBLCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleControl
	///
	//@{
	/// \brief <em>Implementation of \c IOleControl::GetControlInfo</em>
	///
	/// Will be called if the container requests details about the control's keyboard mnemonics and keyboard
	/// behavior.
	///
	/// \param[in, out] pControlInfo The requested details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered actions (verbs) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	//////////////////////////////////////////////////////////////////////
	/// \name MFC clones
	///
	//@{
	/// \brief <em>A rewrite of MFC's \c COleControl::RecreateControlWindow</em>
	///
	/// Destroys and re-creates the control window.
	///
	/// \remarks This rewrite probably isn't complete.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
	void RecreateControlWindow(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name IME support
	///
	//@{
	/// \brief <em>Retrieves a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is requested.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa SetCurrentIMEContextMode, CBLCtlsLibU::IMEModeConstants, get_IMEMode
	/// \else
	///   \sa SetCurrentIMEContextMode, CBLCtlsLibA::IMEModeConstants, get_IMEMode
	/// \endif
	IMEModeConstants GetCurrentIMEContextMode(HWND hWnd);
	/// \brief <em>Sets a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is set.
	/// \param[in] IMEMode A constant out of the \c IMEModeConstants enumeration specifying the IME
	///            context mode to apply.
	///
	/// \if UNICODE
	///   \sa GetCurrentIMEContextMode, CBLCtlsLibU::IMEModeConstants, put_IMEMode
	/// \else
	///   \sa GetCurrentIMEContextMode, CBLCtlsLibA::IMEModeConstants, put_IMEMode
	/// \endif
	void SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode);
	/// \brief <em>Retrieves the control's effective IME context mode</em>
	///
	/// Retrieves the IME context mode that is set for the control after resolving recursive modes like
	/// \c imeInherit.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::IMEModeConstants, get_IMEMode
	/// \else
	///   \sa CBLCtlsLibA::IMEModeConstants, get_IMEMode
	/// \endif
	IMEModeConstants GetEffectiveIMEMode(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Control window configuration
	///
	//@{
	/// \brief <em>Calculates the extended window style bits</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetStyleBits, SendConfigurationMessages, Create
	DWORD GetExStyleBits(void);
	/// \brief <em>Calculates the window style bits</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* constants specifying the required window styles.
	///
	/// \sa GetExStyleBits, SendConfigurationMessages, Create
	DWORD GetStyleBits(void);
	/// \brief <em>Configures the control by sending messages</em>
	///
	/// Sends \c WM_* and \c LB_* messages to the control window to make it match the current property
	/// settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa GetExStyleBits, GetStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates a \c MousePointerConstants type mouse cursor into a \c HCURSOR type mouse cursor</em>
	///
	/// \param[in] mousePointer The \c MousePointerConstants type mouse cursor to translate.
	///
	/// \return The translated \c HCURSOR type mouse cursor.
	///
	/// \if UNICODE
	///   \sa CBLCtlsLibU::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \else
	///   \sa CBLCtlsLibA::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \endif
	HCURSOR MousePointerConst2hCursor(MousePointerConstants mousePointer);
	/// \brief <em>Translates an item's unique ID to its index</em>
	///
	/// \param[in] ID The unique ID of the item whose index is requested.
	///
	/// \return The requested item's index if successful; -1 otherwise.
	///
	/// \sa itemIDs, ItemIndexToID, ListBoxItems::get_Item, ListBoxItems::Remove
	int IDToItemIndex(LONG ID);
	/// \brief <em>Translates an item's index to its unique ID</em>
	///
	/// \param[in] itemIndex The index of the item whose unique ID is requested.
	///
	/// \return The requested item's unique ID if successful; -1 otherwise.
	///
	/// \sa itemIDs, IDToItemIndex, ListBoxItem::get_ID
	LONG ItemIndexToID(int itemIndex);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the item that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of \c HitTestConstants flags, that holds further details about the
	///             control's part below the specified point.
	/// \param[in] autoScroll If \c TRUE and the specified point lies directly above or below the list box,
	///            the control is scrolled vertically by 1 line.
	///
	/// \return The "hit" item's zero-based index.
	///
	/// \if UNICODE
	///   \sa HitTest, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa HitTest, CBLCtlsLibA::HitTestConstants
	/// \endif
	int HitTest(LONG x, LONG y, HitTestConstants* pFlags, BOOL autoScroll = FALSE);
	/// \brief <em>Checks whether the specified item is truncated</em>
	///
	/// \param[in] itemIndex The item to check.
	/// \param[in,out] pLabel Receives the item's text if the item is truncated.
	/// \param[in] bufferSize The size in characters of the buffer specified by \c pLabel.
	///
	/// \return \c TRUE, if the item is truncated; otherwise \c FALSE.
	///
	/// \sa OnToolTipGetDispInfoNotificationW
	BOOL IsItemTruncated(int itemIndex, __in_opt LPTSTR pLabel, int bufferSize);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);
	/// \brief <em>Auto-scrolls the control</em>
	///
	/// \sa OnTimer, Raise_DragMouseMove, DragDropStatus::AutoScrolling
	void AutoScroll(void);

	/// \brief <em>Draws the insertion mark</em>
	///
	/// This method is called by \c OnPaint to draw the control's insertion mark.
	///
	/// \param[in] targetDC The device context to be used for drawing.
	///
	/// \sa OnPaint
	void DrawInsertionMark(CDCHandle targetDC);
	/// \brief <em>Retrieves whether the logical left mouse button is held down</em>
	///
	/// \return \c TRUE if the logical left mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsRightMouseButtonDown
	BOOL IsLeftMouseButtonDown(void);
	/// \brief <em>Retrieves whether the logical right mouse button is held down</em>
	///
	/// \return \c TRUE if the logical right mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsLeftMouseButtonDown
	BOOL IsRightMouseButtonDown(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Item insertion and deletion
	///
	//@{
	/// \brief <em>Buffers the value of a new item's \c ItemData property until it is required by \c OnAddString or \c OnInsertString</em>
	///
	/// \param[in] itemData The value of the new item's \c ItemData property.
	///
	/// \sa Properties::itemDataOfInsertedItems, OnAddString, OnInsertString
	void BufferItemData(LPARAM itemData);
	/// \brief <em>Increments the \c silentItemInsertions flag</em>
	///
	/// \sa LeaveSilentItemInsertionSection, Flags::silentItemInsertions, EnterSilentItemDeletionSection,
	///     ListBoxItem::put_Text, Raise_InsertingItem, Raise_InsertedItem
	void EnterSilentItemInsertionSection(void);
	/// \brief <em>Decrements the \c silentItemInsertions flag</em>
	///
	/// \sa EnterSilentItemInsertionSection, Flags::silentItemInsertions, LeaveSilentItemDeletionSection,
	///     ListBoxItem::put_Text, Raise_InsertingItem, Raise_InsertedItem
	void LeaveSilentItemInsertionSection(void);
	/// \brief <em>Increments the \c silentItemDeletions flag</em>
	///
	/// \sa LeaveSilentItemDeletionSection, Flags::silentItemDeletions, EnterSilentItemInsertionSection,
	///     ListBoxItem::put_Text, Raise_RemovingItem, Raise_RemovedItem
	void EnterSilentItemDeletionSection(void);
	/// \brief <em>Decrements the \c silentItemDeletions flag</em>
	///
	/// \sa EnterSilentItemDeletionSection, Flags::silentItemDeletions, LeaveSilentItemInsertionSection,
	///     ListBoxItem::put_Text, Raise_RemovingItem, Raise_RemovedItem
	void LeaveSilentItemDeletionSection(void);
	/// \brief <em>Changes an item's unique ID</em>
	///
	/// \param[in] itemIndex The zero-based index of the item for which to change the ID.
	/// \param[in] itemID The item's new unique ID.
	///
	/// \sa ListBoxItem::put_Text
	void SetItemID(int itemIndex, LONG itemID);
	/// \brief <em>Decrements \c lastItemID by 1</em>
	///
	/// \sa lastItemID, ListBoxItem::put_Text
	void DecrementNextItemID(void);
	//@}
	//////////////////////////////////////////////////////////////////////


	/// \brief <em>Holds constants and flags used with IME support</em>
	struct IMEFlags
	{
	protected:
		/// \brief <em>A table of IME modes to use for Chinese input language</em>
		///
		/// \sa GetIMECountryTable, japaneseIMETable, koreanIMETable
		static IMEModeConstants chineseIMETable[10];
		/// \brief <em>A table of IME modes to use for Japanese input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants japaneseIMETable[10];
		/// \brief <em>A table of IME modes to use for Korean input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants koreanIMETable[10];

	public:
		/// \brief <em>The handle of the default IME context</em>
		HIMC hDefaultIMC;

		IMEFlags()
		{
			hDefaultIMC = NULL;
		}

		/// \brief <em>Retrieves a table of IME modes to use for the current keyboard layout</em>
		///
		/// Retrieves a table of IME modes which can be used to map \c IME_CMODE_* constants to
		/// \c IMEModeConstants constants. The table depends on the current keyboard layout.
		///
		/// \param[in,out] table The IME mode table for the currently active keyboard layout.
		///
		/// \if UNICODE
		///   \sa CBLCtlsLibU::IMEModeConstants, GetCurrentIMEContextMode
		/// \else
		///   \sa CBLCtlsLibA::IMEModeConstants, GetCurrentIMEContextMode
		/// \endif
		static void GetIMECountryTable(IMEModeConstants table[10])
		{
			WORD languageID = LOWORD(GetKeyboardLayout(0));
			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				switch(languageID) {
					case MAKELANGID(LANG_JAPANESE, SUBLANG_DEFAULT):
						CopyMemory(table, japaneseIMETable, sizeof(japaneseIMETable));
						return;
						break;
					case MAKELANGID(LANG_KOREAN, SUBLANG_DEFAULT):
						CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
						return;
						break;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
				if(languageID == MAKELANGID(LANG_KOREAN, SUBLANG_SYS_DEFAULT)) {
					CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
					return;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if((languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SINGAPORE)) && (languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_MACAU))) {
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
		}
	} IMEFlags;

	/// \brief <em>Holds a control instance's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>Holds a font property's settings</em>
		typedef struct FontProperty
		{
		protected:
			/// \brief <em>Holds the control's default font</em>
			///
			/// \sa GetDefaultFont
			static FONTDESC defaultFont;

		public:
			/// \brief <em>Holds whether we're listening for events fired by the font object</em>
			///
			/// If greater than 0, we're advised to the \c IFontDisp object identified by \c pFontDisp. I. e.
			/// we will be notified if a property of the font object changes. If 0, we won't receive any events
			/// fired by the \c IFontDisp object.
			///
			/// \sa pFontDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>Flag telling \c OnSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnSetFont
			UINT dontGetFontObject : 1;
			/// \brief <em>The control's current font</em>
			///
			/// \sa ApplyFont, owningFontResource
			CFont currentFont;
			/// \brief <em>If \c TRUE, \c currentFont may destroy the font resource; otherwise not</em>
			///
			/// \sa currentFont
			UINT owningFontResource : 1;
			/// \brief <em>A pointer to the font object's implementation of \c IFontDisp</em>
			IFontDisp* pFontDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<ListBox> >* pPropertyNotifySink;

			FontProperty()
			{
				watching = 0;
				dontGetFontObject = FALSE;
				owningFontResource = TRUE;
				pFontDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~FontProperty()
			{
				Release();
			}

			FontProperty& operator =(const FontProperty& source)
			{
				Release();

				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				pFontDisp = source.pFontDisp;
				if(pFontDisp) {
					pFontDisp->AddRef();
				}
				owningFontResource = source.owningFontResource;
				if(!owningFontResource) {
					currentFont.Attach(source.currentFont.m_hFont);
				}
				dontGetFontObject = source.dontGetFontObject;

				if(source.watching > 0) {
					StartWatching();
				}

				return *this;
			}

			/// \brief <em>Retrieves a default font that may be used</em>
			///
			/// \return A \c FONTDESC structure containing the default font.
			///
			/// \sa defaultFont
			static FONTDESC GetDefaultFont(void)
			{
				return defaultFont;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(ListBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ListBox> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(owningFontResource) {
					if(!currentFont.IsNull()) {
						currentFont.DeleteObject();
					}
				} else {
					currentFont.Detach();
				}
				if(pFontDisp) {
					pFontDisp->Release();
					pFontDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pFontDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} FontProperty;

		/// \brief <em>Holds a picture property's settings</em>
		typedef struct PictureProperty
		{
			/// \brief <em>Holds whether we're listening for events fired by the picture object</em>
			///
			/// If greater than 0, we're advised to the \c IPictureDisp object identified by \c pPictureDisp.
			/// I. e. we will be notified if a property of the picture object changes. If 0, we won't receive any
			/// events fired by the \c IPictureDisp object.
			///
			/// \sa pPictureDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>A pointer to the picture object's implementation of \c IPictureDisp</em>
			IPictureDisp* pPictureDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<ListBox> >* pPropertyNotifySink;

			PictureProperty()
			{
				watching = 0;
				pPictureDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~PictureProperty()
			{
				Release();
			}

			PictureProperty& operator =(const PictureProperty& source)
			{
				Release();

				pPictureDisp = source.pPictureDisp;
				if(pPictureDisp) {
					pPictureDisp->AddRef();
				}
				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				if(source.watching > 0) {
					StartWatching();
				}
				return *this;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(ListBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ListBox> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(pPictureDisp) {
					pPictureDisp->Release();
					pPictureDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pPictureDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} PictureProperty;

		/// \brief <em>The autoscroll-zone's default width</em>
		///
		/// The default width (in pixels) of the border around the control's client area, that's
		/// sensitive for auto-scrolling during a drag'n'drop operation. If the mouse cursor's position
		/// lies within this area during a drag'n'drop operation, the control will be auto-scrolled.
		///
		/// \sa dragScrollTimeBase, Raise_OLEDragMouseMove
		static const int DRAGSCROLLZONEWIDTH = 16;

		/// \brief <em>Holds the \c AllowDragDrop property's setting</em>
		///
		/// \sa get_AllowDragDrop, put_AllowDragDrop
		UINT allowDragDrop : 1;
		/// \brief <em>Holds the \c AllowItemSelection property's setting</em>
		///
		/// \sa get_AllowItemSelection, put_AllowItemSelection
		UINT allowItemSelection : 1;
		/// \brief <em>Holds the \c AlwaysShowVerticalScrollBar property's setting</em>
		///
		/// \sa get_AlwaysShowVerticalScrollBar, put_AlwaysShowVerticalScrollBar
		UINT alwaysShowVerticalScrollBar : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c ColumnWidth property's setting</em>
		///
		/// \sa get_ColumnWidth, put_ColumnWidth
		long columnWidth;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DontRedraw property's setting</em>
		///
		/// \sa get_DontRedraw, put_DontRedraw
		UINT dontRedraw : 1;
		/// \brief <em>Holds the \c DragScrollTimeBase property's setting</em>
		///
		/// \sa get_DragScrollTimeBase, put_DragScrollTimeBase
		long dragScrollTimeBase;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the \c ForeColor property's setting</em>
		///
		/// \sa get_ForeColor, put_ForeColor
		OLE_COLOR foreColor;
		/// \brief <em>Holds the \c HasStrings property's setting</em>
		///
		/// \sa get_HasStrings, put_HasStrings
		UINT hasStrings : 1;
		/// \brief <em>Holds the \c ilHighResolution imagelist</em>
		///
		/// \sa get_hImageList, put_hImageList
		HIMAGELIST hHighResImageList;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c IMEMode property's setting</em>
		///
		/// \sa get_IMEMode, put_IMEMode
		IMEModeConstants IMEMode;
		/// \brief <em>Holds the \c InsertMarkColor property's setting</em>
		///
		/// \sa get_InsertMarkColor, put_InsertMarkColor
		OLE_COLOR insertMarkColor;
		/// \brief <em>Holds the \c InsertMarkStyle property's setting</em>
		///
		/// \sa get_InsertMarkStyle, put_InsertMarkStyle
		InsertMarkStyleConstants insertMarkStyle;
		/// \brief <em>Holds the \c IntegralHeight property's setting</em>
		///
		/// \sa get_IntegralHeight, put_IntegralHeight
		UINT integralHeight : 1;
		/// \brief <em>Holds the \c ItemHeight property's setting</em>
		///
		/// \sa get_ItemHeight, put_ItemHeight
		OLE_YSIZE_PIXELS itemHeight;
		/// \brief <em>Holds the value that the \c ItemHeight property was set to</em>
		///
		/// \sa get_ItemHeight, put_ItemHeight
		OLE_YSIZE_PIXELS setItemHeight;
		/// \brief <em>Holds the \c Locale property's setting</em>
		///
		/// \sa get_Locale, put_Locale
		long locale;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c MultiColumn property's setting</em>
		///
		/// \sa get_MultiColumn, put_MultiColumn
		UINT multiColumn : 1;
		/// \brief <em>Holds the \c MultiSelect property's setting</em>
		///
		/// \sa get_MultiSelect, put_MultiSelect
		MultiSelectConstants multiSelect;
		/// \brief <em>Holds the \c OLEDragImageStyle property's setting</em>
		///
		/// \sa get_OLEDragImageStyle, put_OLEDragImageStyle
		OLEDragImageStyleConstants oleDragImageStyle;
		/// \brief <em>Holds the \c OwnerDrawItems property's setting</em>
		///
		/// \sa get_OwnerDrawItems, put_OwnerDrawItems
		OwnerDrawItemsConstants ownerDrawItems;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c ProcessTabs property's setting</em>
		///
		/// \sa get_ProcessTabs, put_ProcessTabs
		UINT processTabs : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		UINT registerForOLEDragDrop : 1;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c ScrollableWidth property's setting</em>
		///
		/// \sa get_ScrollableWidth, put_ScrollableWidth
		OLE_XSIZE_PIXELS scrollableWidth;
		/// \brief <em>Holds the \c Sorted property's setting</em>
		///
		/// \sa get_Sorted, put_Sorted
		UINT sorted : 1;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		#ifdef USE_STL
			/// \brief <em>Holds the \c TabStops property's setting</em>
			///
			/// \sa get_TabStops, put_TabStops
			std::vector<long> tabStops;
		#else
			/// \brief <em>Holds the \c TabStops property's setting</em>
			///
			/// \sa get_TabStops, put_TabStops
			CAtlArray<long> tabStops;
		#endif
		/// \brief <em>Holds the \c TabWidth property's setting in dialog template units</em>
		///
		/// \sa get_TabWidth, put_TabWidth
		long tabWidthInDTUs;
		/// \brief <em>Holds the \c TabWidth property's setting in pixels</em>
		///
		/// \sa get_TabWidth, put_TabWidth
		long tabWidthInPixels;
		/// \brief <em>Holds the \c ToolTips property's setting</em>
		///
		/// \sa get_ToolTips, put_ToolTips
		ToolTipsConstants toolTips;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;
		/// \brief <em>Holds the \c VirtualMode property's setting</em>
		///
		/// \sa get_VirtualMode, put_VirtualMode
		UINT virtualMode : 1;

		#ifdef USE_STL
			/// \brief <em>Holds the values of the \c ItemData property of items currently being inserted</em>
			///
			/// \sa BufferItemData, OnAddString, OnInsertString
			std::stack<LPARAM> itemDataOfInsertedItems;
		#else
			/// \brief <em>Holds the values of the \c ItemData property of items currently being inserted</em>
			///
			/// \sa BufferItemData, OnAddString, OnInsertString
			CAtlList<LPARAM> itemDataOfInsertedItems;
		#endif

		Properties()
		{
			ResetToDefaults();
		}

		~Properties()
		{
			Release();
		}

		/// \brief <em>Prepares the object for destruction</em>
		void Release(void)
		{
			font.Release();
			mouseIcon.Release();
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			allowDragDrop = FALSE;
			allowItemSelection = TRUE;
			alwaysShowVerticalScrollBar = FALSE;
			appearance = a3D;
			backColor = 0x80000000 | COLOR_WINDOW;
			borderStyle = bsNone;
			columnWidth = -1;
			disabledEvents = static_cast<DisabledEventsConstants>(deMouseEvents | deClickEvents | deKeyboardEvents | deItemInsertionEvents | deItemDeletionEvents | deFreeItemData | deProcessKeyboardInput);
			dontRedraw = FALSE;
			dragScrollTimeBase = -1;
			enabled = TRUE;
			foreColor = 0x80000000 | COLOR_WINDOWTEXT;
			hasStrings = TRUE;
			hHighResImageList = NULL;
			hoverTime = -1;
			IMEMode = imeInherit;
			insertMarkColor = RGB(0, 0, 0);
			insertMarkStyle = imsImproved;
			integralHeight = FALSE;
			itemHeight = -1;
			setItemHeight = -1;
			locale = LOCALE_USER_DEFAULT;
			mousePointer = mpDefault;
			multiColumn = FALSE;
			multiSelect = msNone;
			oleDragImageStyle = odistClassic;
			ownerDrawItems = odiDontOwnerDraw;
			processContextMenuKeys = TRUE;
			processTabs = TRUE;
			registerForOLEDragDrop = FALSE;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			scrollableWidth = 0;
			sorted = FALSE;
			supportOLEDragImages = TRUE;
			#ifdef USE_STL
				tabStops.clear();
			#else
				tabStops.RemoveAll();
			#endif
			tabWidthInDTUs = -1;
			tabWidthInPixels = -1;
			toolTips = static_cast<ToolTipsConstants>(0);
			useSystemFont = TRUE;
			virtualMode = FALSE;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, \c RecreateControlWindow won't recreate the control window</em>
		///
		/// \sa RecreateControlWindow, Load
		UINT dontRecreate : 1;
		/// \brief <em>If not 0, we won't raise \c Insert*Item events</em>
		///
		/// \sa silentItemDeletions, EnterSilentItemInsertionSection, LeaveSilentItemInsertionSection,
		///     OnAddString, OnInsertString, Raise_InsertingItem, Raise_InsertedItem
		int silentItemInsertions;
		/// \brief <em>If not 0, we won't raise \c Remov*Item events</em>
		///
		/// \sa silentItemInsertions, EnterSilentItemDeletionSection, LeaveSilentItemDeletionSection,
		///     OnDeleteString, OnResetContent, Raise_RemovingItem, Raise_RemovedItem
		int silentItemDeletions;
		/// \brief <em>If \c TRUE, the control has been activated by mouse and needs to be UI-activated by \c OnSetFocus</em>
		///
		/// ATL always UI-activates the control in \c OnMouseActivate. If the control is activated by mouse,
		/// \c WM_SETFOCUS is sent after \c WM_MOUSEACTIVATE, but Visual Basic 6 won't raise the \c Validate
		/// event if the control already is UI-activated when it receives the focus. Therefore we need to delay
		/// UI-activation.
		///
		/// \sa OnMouseActivate, OnSetFocus, OnKillFocus
		UINT uiActivationPending : 1;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;

		Flags()
		{
			dontRecreate = FALSE;
			silentItemInsertions = 0;
			silentItemDeletions = 0;
			uiActivationPending = FALSE;
			usingThemes = FALSE;
		}
	} flags;

	//////////////////////////////////////////////////////////////////////
	/// \name Item management
	///
	//@{
	#ifdef USE_STL
		/// \brief <em>A list of all item IDs</em>
		///
		/// Holds the unique IDs of all list box items in the control. The item's index in the control equals
		/// its index in the vector.
		///
		/// \sa OnAddString, OnInsertString, OnResetContent, OnDeleteString, ListBoxItem::get_ID
		std::vector<LONG> itemIDs;
	#else
		/// \brief <em>A list of all items</em>
		///
		/// Holds the unique IDs of all list box items in the control. The item's index in the control equals
		/// its index in the vector.
		///
		/// \sa OnAddString, OnInsertString, OnResetContent, OnDeleteString, ListBoxItem::get_ID
		CAtlArray<LONG> itemIDs;
	#endif
	/// \brief <em>Retrieves a new unique item ID at each call</em>
	///
	/// \return A new unique item ID.
	///
	/// \sa itemIDs, ListBoxItem::get_ID
	LONG GetNewItemID(void);
	#ifdef USE_STL
		/// \brief <em>A map of all \c ListBoxItemContainer objects that we've created</em>
		///
		/// Holds pointers to all \c ListBoxItemContainer objects that we've created. We use this map to
		/// inform the containers of item deletions. The container's ID is stored as key; the container's
		/// \c IItemContainer implementation is stored as value.
		///
		/// \sa CreateItemContainer, RegisterItemContainer, ListBoxItemContainer
		std::unordered_map<DWORD, IItemContainer*> itemContainers;
	#else
		/// \brief <em>A map of all \c ListBoxItemContainer objects that we've created</em>
		///
		/// Holds pointers to all \c ListBoxItemContainer objects that we've created. We use this map to
		/// inform the containers of item deletions. The container's ID is stored as key; the container's
		/// \c IItemContainer implementation is stored as value.
		///
		/// \sa CreateItemContainer, RegisterItemContainer, ListBoxItemContainer
		CAtlMap<DWORD, IItemContainer*> itemContainers;
	#endif
	/// \brief <em>Registers the specified \c ListBoxItemContainer collection</em>
	///
	/// Registers the specified \c ListBoxItemContainer collection so that it is informed of item deletions.
	///
	/// \param[in] pContainer The container's \c IItemContainer implementation.
	///
	/// \sa DeregisterItemContainer, itemContainers, RemoveItemFromItemContainers
	void RegisterItemContainer(IItemContainer* pContainer);
	/// \brief <em>De-registers the specified \c ListBoxItemContainer collection</em>
	///
	/// De-registers the specified \c ListBoxItemContainer collection so that it no longer is informed of
	/// item deletions.
	///
	/// \param[in] containerID The container's ID.
	///
	/// \sa RegisterItemContainer, itemContainers
	void DeregisterItemContainer(DWORD containerID);
	/// \brief <em>Removes the specified item from all registered \c ListBoxItemContainer collections</em>
	///
	/// \param[in] itemIdentifier <strong>Non-virtual mode:</strong> The unique ID of the item to remove.
	///            <strong>Virtual mode:</strong> The zero-based index of the item to remove. If -1, all
	///            items are removed.
	///
	/// \sa itemContainers, RegisterItemContainer, OnDeleteString, OnResetContent,
	///     Raise_DestroyedControlWindow
	void RemoveItemFromItemContainers(LONG itemIdentifier);
	/// \brief <em>Replaces the specified item ID in all registered \c ListBoxItemContainer collections</em>
	///
	/// \param[in] oldItemID The old item ID.
	/// \param[in] newItemID The new item ID.
	///
	/// \remarks This method is not supported in virtual mode.
	///
	/// \sa itemContainers, RegisterItemContainer, SetItemID
	void UpdateItemIDInItemContainers(LONG oldItemID, LONG newItemID);
	///@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the index of the current caret item</em>
	int currentCaretItem;
	/// \brief <em>Holds the index of the item below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deMouseEvents being set.
	int itemUnderMouse;

	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
	protected:
		/// \brief <em>Holds all mouse buttons that may cause a click event in the immediate future</em>
		///
		/// A bit field of \c SHORT values representing those mouse buttons that are currently pressed and
		/// may cause a click event in the immediate future.
		///
		/// \sa StoreClickCandidate, IsClickCandidate, RemoveClickCandidate, Raise_Click, Raise_MClick,
		///     Raise_RClick, Raise_XClick
		SHORT clickCandidates;

	public:
		/// \brief <em>If \c TRUE, the \c MouseEnter event already was raised</em>
		///
		/// \sa Raise_MouseEnter
		UINT enteredControl : 1;
		/// \brief <em>If \c TRUE, the \c MouseHover event already was raised</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa Raise_MouseHover
		UINT hoveredControl : 1;
		/// \brief <em>Holds the mouse cursor's last position</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		POINT lastPosition;
		/// \brief <em>Holds the index of the last clicked item</em>
		///
		/// Holds the index of the last clicked item. We use this to ensure that the \c *DblClick events
		/// are not raised accidently.
		///
		/// \attention This member is not reliable with \c deClickEvents being set.
		///
		/// \sa Raise_Click, Raise_DblClick, Raise_MClick, Raise_MDblClick, Raise_RClick, Raise_RDblClick,
		///     Raise_XClick, Raise_XDblClick
		int lastClickedItem;
		/// \brief <em>Holds the details of the \c WM_LBUTTONDOWN message that shall be sent to \c DefWindowProc in \c OnLButtonUp</em>
		///
		/// If drag'n'drop is allowed, we intercept \c WM_LBUTTONDOWN messages if the button has been pressed
		/// over a selected item. If we wouldn't do this, all items would get unselected making it impossible
		/// to drag multiple items. In \c OnLButtonUp we call \c DefWindowProc for the intercepted message so
		/// that all items except the clicked one get unselected if no item dragging has been detected.
		///
		/// \sa OnLButtonDown, OnLButtonUp, OnMouseMove
		MSG mouseDownMessageToSendOnMouseUp;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedItem = -1;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
			lastClickedItem = -1;
		}

		/// \brief <em>Changes flags to indicate the \c MouseHover event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl, LeaveControl
		void HoverControl(void)
		{
			enteredControl = TRUE;
			hoveredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseLeave event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl
		void LeaveControl(void)
		{
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedItem = -1;
		}

		/// \brief <em>Stores a mouse button as click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa clickCandidates, IsClickCandidate, RemoveClickCandidate
		void StoreClickCandidate(SHORT button)
		{
			// avoid combined click events
			if(clickCandidates == 0) {
				clickCandidates |= button;
			}
		}

		/// \brief <em>Retrieves whether a mouse button is a click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a click candidate; otherwise \c FALSE.
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa clickCandidates, StoreClickCandidate, RemoveClickCandidate
		BOOL IsClickCandidate(SHORT button)
		{
			return (clickCandidates & button);
		}

		/// \brief <em>Removes a mouse button from the list of click candidates</em>
		///
		/// \param[in] button The mouse button to remove.
		///
		/// \sa clickCandidates, RemoveAllClickCandidates, StoreClickCandidate, IsClickCandidate
		void RemoveClickCandidate(SHORT button)
		{
			clickCandidates &= ~button;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa clickCandidates, RemoveClickCandidate, StoreClickCandidate, IsClickCandidate
		void RemoveAllClickCandidates(void)
		{
			clickCandidates = 0;
		}
	} MouseStatus;

	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus;

	/// \brief <em>Holds the position of the control's insertion mark</em>
	struct InsertMark
	{
		/// \brief <em>If set to \c TRUE, the insertion mark won't be drawn</em>
		///
		/// \sa OnScroll, DrawInsertionMark
		UINT hidden : 1;
		/// \brief <em>The zero-based index of the item at which the insertion mark is placed</em>
		int itemIndex;
		/// \brief <em>The insertion mark's position relative to the item</em>
		UINT afterItem : 1;
		/// \brief <em>The insertion mark's color</em>
		COLORREF color;

		InsertMark()
		{
			Reset();
		}

		/// \brief <em>Resets all member variables</em>
		void Reset(void)
		{
			hidden = FALSE;
			itemIndex = -1;
			afterItem = FALSE;
			color = RGB(0, 0, 0);
		}
	} insertMark;

	//////////////////////////////////////////////////////////////////////
	/// \name Drag'n'Drop
	///
	//@{
	/// \brief <em>The \c CLSID_WICImagingFactory object used to create WIC objects that are required during drag image creation</em>
	///
	/// \sa OnGetDragImage, CreateThumbnail
	CComPtr<IWICImagingFactory> pWICImagingFactory;
	/// \brief <em>Creates a thumbnail of the specified icon in the specified size</em>
	///
	/// \param[in] hIcon The icon to create the thumbnail for.
	/// \param[in] size The thumbnail's size in pixels.
	/// \param[in,out] pBits The thumbnail's DIB bits.
	/// \param[in] doAlphaChannelPostProcessing WIC has problems to handle the alpha channel of the icon
	///            specified by \c hIcon. If this parameter is set to \c TRUE, some post-processing is done
	///            to correct the pixel failures. Otherwise the failures are not corrected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnGetDragImage, pWICImagingFactory
	HRESULT CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing);

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>The \c IListBoxItemContainer implementation of the collection of the dragged items</em>
		IListBoxItemContainer* pDraggedItems;
		/// \brief <em>The handle of the image list containing the drag image</em>
		///
		/// \sa get_hDragImageList
		HIMAGELIST hDragImageList;
		/// \brief <em>Enables or disables auto-destruction of \c hDragImageList</em>
		///
		/// Controls whether the image list defined by \c hDragImageList is destroyed automatically. If set to
		/// \c TRUE, it is destroyed in \c EndDrag; otherwise not.
		///
		/// \sa hDragImageList, EndDrag
		UINT autoDestroyImgLst : 1;
		/// \brief <em>Indicates whether the drag image is visible or hidden</em>
		///
		/// If this value is 0, the drag image is visible; otherwise not.
		///
		/// \sa get_hDragImageList, get_ShowDragImage, put_ShowDragImage, ShowDragImage, HideDragImage,
		///     IsDragImageVisible
		int dragImageIsHidden;
		/// \brief <em>The zero-based index of the last drop target</em>
		int lastDropTarget;

		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>The currently dragged data for the case that the we're the drag source</em>
		CComPtr<IDataObject> pSourceDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the mouse button (as \c MK_* constant) that the drag'n'drop operation is performed with</em>
		DWORD draggingMouseButton;
		/// \brief <em>If \c TRUE, we'll hide and re-show the drag image in \c IDropTarget::DragEnter so that the item count label is displayed</em>
		///
		/// \sa DragEnter, OLEDrag
		UINT useItemCountLabelHack : 1;
		/// \brief <em>Holds the \c IDataObject to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		IDataObject* drop_pDataObject;
		/// \brief <em>Holds the mouse position to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		POINT drop_mousePosition;
		/// \brief <em>Holds the drop effect to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		DWORD drop_effect;
		//@}
		//////////////////////////////////////////////////////////////////////

		/// \brief <em>Holds data and flags related to auto-scrolling</em>
		///
		/// \sa AutoScroll
		struct AutoScrolling
		{
			/// \brief <em>Holds the current speed multiplier used for horizontal auto-scrolling</em>
			LONG currentHScrollVelocity;
			/// \brief <em>Holds the current speed multiplier used for vertical auto-scrolling</em>
			LONG currentVScrollVelocity;
			/// \brief <em>Holds the current interval of the auto-scroll timer</em>
			LONG currentTimerInterval;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled downwards</em>
			DWORD lastScroll_Down;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the left</em>
			DWORD lastScroll_Left;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the right</em>
			DWORD lastScroll_Right;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled upwardly</em>
			DWORD lastScroll_Up;

			AutoScrolling()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				currentHScrollVelocity = 0;
				currentVScrollVelocity = 0;
				currentTimerInterval = 0;
				lastScroll_Down = 0;
				lastScroll_Left = 0;
				lastScroll_Right = 0;
				lastScroll_Up = 0;
			}
		} autoScrolling;

		/// \brief <em>Holds data required for drag-detection</em>
		struct Candidate
		{
			/// \brief <em>The zero-based index of the item, that might be dragged</em>
			int itemIndex;
			/// \brief <em>The position in pixels at which dragging might start</em>
			POINT position;
			/// \brief <em>The mouse button (as \c MK_* constant) that the drag'n'drop operation might be performed with</em>
			int button;

			Candidate()
			{
				itemIndex = -1;
			}
		} candidate;

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pSourceDataObject = NULL;
			pDropTargetHelper = NULL;
			draggingMouseButton = 0;

			pDraggedItems = NULL;
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
			lastDropTarget = -1;
			drop_pDataObject = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
			}
			ATLASSERT(!pDraggedItems);
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pDraggedItems) {
				this->pDraggedItems->Release();
				this->pDraggedItems = NULL;
			}
			if(hDragImageList && autoDestroyImgLst) {
				ImageList_Destroy(hDragImageList);
			}
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
			lastDropTarget = -1;

			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			if(this->pSourceDataObject) {
				this->pSourceDataObject = NULL;
			}
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;
			drop_pDataObject = NULL;
		}

		/// \brief <em>Decrements the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, HideDragImage, IsDragImageVisible
		void ShowDragImage(BOOL commonDragDropOnly)
		{
			if(hDragImageList) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					ImageList_DragShowNolock(TRUE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					pDropTargetHelper->Show(TRUE);
				}
			}
		}

		/// \brief <em>Increments the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, ShowDragImage, IsDragImageVisible
		void HideDragImage(BOOL commonDragDropOnly)
		{
			if(hDragImageList) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					ImageList_DragShowNolock(FALSE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					pDropTargetHelper->Show(FALSE);
				}
			}
		}

		/// \brief <em>Retrieves whether we're currently displaying a drag image</em>
		///
		/// \return \c TRUE, if we're displaying a drag image; otherwise \c FALSE.
		///
		/// \sa dragImageIsHidden, ShowDragImage, HideDragImage
		BOOL IsDragImageVisible(void)
		{
			return (dragImageIsHidden == 0);
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation started</em>
		///
		/// \param[in] hWndListBox The list box window, that the method will work on to calculate the position
		///            of the drag image's hotspot.
		/// \param[in] pDraggedItems The \c IListBoxItemContainer implementation of the collection of
		///            the dragged items.
		/// \param[in] hDragImageList The image list containing the drag image that shall be used to
		///            visualize the drag'n'drop operation. If -1, the method will create the drag image
		///            itself; if \c NULL, no drag image will be displayed.
		/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImageList parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImageList parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImageList parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImageList parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa EndDrag
		HRESULT BeginDrag(HWND hWndListBox, IListBoxItemContainer* pDraggedItems, HIMAGELIST hDragImageList, PINT pXHotSpot, PINT pYHotSpot)
		{
			ATLASSUME(pDraggedItems);
			if(!pDraggedItems) {
				return E_INVALIDARG;
			}

			UINT b = FALSE;
			if(hDragImageList == static_cast<HIMAGELIST>(LongToHandle(-1))) {
				OLE_HANDLE h = NULL;
				OLE_XPOS_PIXELS xUpperLeft = 0;
				OLE_YPOS_PIXELS yUpperLeft = 0;
				if(FAILED(pDraggedItems->CreateDragImage(&xUpperLeft, &yUpperLeft, &h))) {
					return E_FAIL;
				}
				hDragImageList = static_cast<HIMAGELIST>(LongToHandle(h));
				b = TRUE;

				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				::ScreenToClient(hWndListBox, &mousePosition);
				if(CWindow(hWndListBox).GetExStyle() & WS_EX_LAYOUTRTL) {
					SIZE dragImageSize = {0};
					ImageList_GetIconSize(hDragImageList, reinterpret_cast<PINT>(&dragImageSize.cx), reinterpret_cast<PINT>(&dragImageSize.cy));
					*pXHotSpot = xUpperLeft + dragImageSize.cx - mousePosition.x;
				} else {
					*pXHotSpot = mousePosition.x - xUpperLeft;
				}
				*pYHotSpot = mousePosition.y - yUpperLeft;
			}

			if(this->hDragImageList && this->autoDestroyImgLst) {
				ImageList_Destroy(this->hDragImageList);
			}

			this->autoDestroyImgLst = b;
			this->hDragImageList = hDragImageList;
			if(this->pDraggedItems) {
				this->pDraggedItems->Release();
				this->pDraggedItems = NULL;
			}
			pDraggedItems->Clone(&this->pDraggedItems);
			ATLASSUME(this->pDraggedItems);
			this->lastDropTarget = -1;

			dragImageIsHidden = 1;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation ended</em>
		///
		/// \sa BeginDrag
		void EndDrag(void)
		{
			if(pDraggedItems) {
				pDraggedItems->Release();
				pDraggedItems = NULL;
			}
			if(autoDestroyImgLst && hDragImageList) {
				ImageList_Destroy(hDragImageList);
			}
			hDragImageList = NULL;
			dragImageIsHidden = 1;
			lastDropTarget = -1;
			autoScrolling.Reset();
		}

		/// \brief <em>Retrieves whether we're in drag'n'drop mode</em>
		///
		/// \return \c TRUE if we're in drag'n'drop mode; otherwise \c FALSE.
		///
		/// \sa BeginDrag, EndDrag
		BOOL IsDragging(void)
		{
			return (pDraggedItems != NULL);
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			lastDropTarget = -1;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			lastDropTarget = -1;
			autoScrolling.Reset();
		}
	} dragDropStatus;
	///@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds data and flags related to tooltip support</em>
	struct ToolTipStatus
	{
		/// \brief <em>The tooltip control to use</em>
		HWND hWndToolTip;
		/// \brief <em>The item for which the tooltip is currently displayed</em>
		int toolTipItem;
		/// \brief <em>If \c TRUE, the tooltip is displayed over the item text</em>
		UINT inplaceToolTip : 1;
		/// \brief <em>Holds the current tooltip text</em>
		LPSTR pCurrentToolTipA;
		/// \brief <em>Holds the current tooltip text</em>
		LPWSTR pCurrentToolTipW;

		ToolTipStatus()
		{
			hWndToolTip = NULL;
			toolTipItem = -1;
			inplaceToolTip = FALSE;
			pCurrentToolTipA = NULL;
			pCurrentToolTipW = NULL;
		}

		~ToolTipStatus()
		{
			if(::IsWindow(hWndToolTip)) {
				::DestroyWindow(hWndToolTip);
			}
			
			typedef BOOL WINAPI Str_SetPtrAFn(__deref_inout_opt LPSTR*, __in_opt LPCSTR);
			Str_SetPtrAFn* pfnStr_SetPtrA = NULL;
			HMODULE hComctl32DLL = LoadLibrary(TEXT("comctl32.dll"));
			if(hComctl32DLL) {
				pfnStr_SetPtrA = reinterpret_cast<Str_SetPtrAFn*>(GetProcAddress(hComctl32DLL, MAKEINTRESOURCEA(234)));
				FreeLibrary(hComctl32DLL);
			}
			if(pfnStr_SetPtrA) {
				pfnStr_SetPtrA(&pCurrentToolTipA, NULL);
			}
			Str_SetPtrW(&pCurrentToolTipW, NULL);
		}

		/// \brief <em>Relays a message to the tooltip control</em>
		///
		/// \param[in] hWnd The original recipient of the message to relay.
		/// \param[in] message The message to relay.
		/// \param[in] wParam The first parameter of the message to relay.
		/// \param[in] lParam The second parameter of the message to relay.
		void RelayToToolTip(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
		{
			if(hWndToolTip) {
				MSG msg;
				msg.lParam = lParam;
				msg.wParam = wParam;
				msg.message = message;
				msg.hwnd = hWnd;
				SendMessage(hWndToolTip, TTM_RELAYEVENT, 0, reinterpret_cast<LPARAM>(&msg));
			}
		}

		/// \brief <em>Invalidates the \c toolTipItem member</em>
		void InvalidateToolTipItem(void)
		{
			toolTipItem = -1;
			if(pCurrentToolTipA && pCurrentToolTipA != LPSTR_TEXTCALLBACKA) {
				pCurrentToolTipA[0] = 0;
			}
			if(pCurrentToolTipW && pCurrentToolTipW != LPSTR_TEXTCALLBACKW) {
				pCurrentToolTipW[0] = 0;
			}
		}
	} toolTipStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;
		/// \brief <em>The ID of the timer that is used to auto-scroll the control window during drag'n'drop</em>
		static const UINT_PTR ID_DRAGSCROLL = 13;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
	} timers;

	/// \brief <em>The last unique ID assigned to an item</em>
	///
	/// \sa DecrementNextItemID
	LONG lastItemID;

private:
};     // ListBox

OBJECT_ENTRY_AUTO(__uuidof(ListBox), ListBox)