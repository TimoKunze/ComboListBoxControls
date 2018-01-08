//////////////////////////////////////////////////////////////////////
/// \class ComboBox
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c ComboBox</em>
///
/// This class superclasses \c ComboBox and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IComboBox
/// \else
///   \sa CBLCtlsLibA::IComboBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IComboBoxEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "ComboBoxItem.h"
#include "VirtualComboBoxItem.h"
#include "ComboBoxItems.h"
#include "CommonProperties.h"
#include "StringProperties.h"
#include "TargetOLEDataObject.h"


class ATL_NO_VTABLE ComboBox : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IComboBox, &IID_IComboBox, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IComboBox, &IID_IComboBox, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<ComboBox>,
    public IOleControlImpl<ComboBox>,
    public IOleObjectImpl<ComboBox>,
    public IOleInPlaceActiveObjectImpl<ComboBox>,
    public IViewObjectExImpl<ComboBox>,
    public IOleInPlaceObjectWindowlessImpl<ComboBox>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ComboBox>,
    public Proxy_IComboBoxEvents<ComboBox>,
    public IPersistStorageImpl<ComboBox>,
    public IPersistPropertyBagImpl<ComboBox>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<ComboBox>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_ComboBox, &__uuidof(_IComboBoxEvents), &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_ComboBox, &__uuidof(_IComboBoxEvents), &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<ComboBox>,
    public CComCoClass<ComboBox, &CLSID_ComboBox>,
    public CComControl<ComboBox>,
    public IPerPropertyBrowsingImpl<ComboBox>,
    public IDropTarget,
    public ICategorizeProperties,
    public ICreditsProvider
{
	friend class ComboBoxItem;
	friend class ComboBoxItems;

public:
	/// \brief <em>The contained drop-down list box control</em>
	CContainedWindow containedListBox;
	/// \brief <em>The contained edit control</em>
	CContainedWindow containedEdit;

	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	ComboBox();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_COMBOBOX)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("ComboBoxU"), WC_COMBOBOXW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("ComboBoxA"), WC_COMBOBOXA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(ComboBox)
			COM_INTERFACE_ENTRY(IComboBox)
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
		END_COM_MAP()

		BEGIN_PROP_MAP(ComboBox)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AcceptNumbersOnly", DISPID_CB_ACCEPTNUMBERSONLY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_CB_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutoHorizontalScrolling", DISPID_CB_AUTOHORIZONTALSCROLLING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("BackColor", DISPID_CB_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_CB_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CharacterConversion", DISPID_CB_CHARACTERCONVERSION, CLSID_NULL, VT_I4)
			//PROP_ENTRY_TYPE("CueBanner", DISPID_CB_CUEBANNER, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_CB_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_CB_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DoOEMConversion", DISPID_CB_DOOEMCONVERSION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragDropDownTime", DISPID_CB_DRAGDROPDOWNTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DropDownKey", DISPID_CB_DROPDOWNKEY, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Enabled", DISPID_CB_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_CB_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("ForeColor", DISPID_CB_FORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("HasStrings", DISPID_CB_HASSTRINGS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("HoverTime", DISPID_CB_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IMEMode", DISPID_CB_IMEMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IntegralHeight", DISPID_CB_INTEGRALHEIGHT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ItemHeight", DISPID_CB_ITEMHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ListAlwaysShowVerticalScrollBar", DISPID_CB_LISTALWAYSSHOWVERTICALSCROLLBAR, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ListBackColor", DISPID_CB_LISTBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("ListDragScrollTimeBase", DISPID_CB_LISTDRAGSCROLLTIMEBASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ListForeColor", DISPID_CB_LISTFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("ListHeight", DISPID_CB_LISTHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ListInsertMarkColor", DISPID_CB_LISTINSERTMARKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("ListScrollableWidth", DISPID_CB_LISTSCROLLABLEWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ListWidth", DISPID_CB_LISTWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Locale", DISPID_CB_LOCALE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MaxTextLength", DISPID_CB_MAXTEXTLENGTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MinVisibleItems", DISPID_CB_MINVISIBLEITEMS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_CB_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_CB_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("OwnerDrawItems", DISPID_CB_OWNERDRAWITEMS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_CB_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_CB_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_CB_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SelectionFieldHeight", DISPID_CB_SELECTIONFIELDHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Sorted", DISPID_CB_SORTED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Style", DISPID_CB_STYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_CB_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("Text", DISPID_CB_TEXT, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_CB_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(ComboBox)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IComboBoxEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(ComboBox)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			INDEXED_MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_CTLCOLOREDIT, OnCtlColorEdit)
			MESSAGE_HANDLER(WM_CTLCOLORLISTBOX, OnCtlColorListBox)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_INPUTLANGCHANGE, OnInputLangChange)
			INDEXED_MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			INDEXED_MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
			INDEXED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTEXT, OnSetText)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			INDEXED_MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			INDEXED_MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(OCM_COMPAREITEM, OnReflectedCompareItem)
			MESSAGE_HANDLER(OCM_CTLCOLOREDIT, OnCtlColorEdit)
			MESSAGE_HANDLER(OCM_DRAWITEM, OnReflectedDrawItem)
			MESSAGE_HANDLER(OCM_MEASUREITEM, OnReflectedMeasureItem)

			MESSAGE_HANDLER(CB_ADDSTRING, OnAddString)
			MESSAGE_HANDLER(CB_DELETESTRING, OnDeleteString)
			MESSAGE_HANDLER(CB_INSERTSTRING, OnInsertString)
			MESSAGE_HANDLER(CB_RESETCONTENT, OnResetContent)
			MESSAGE_HANDLER(CB_SETCUEBANNER, OnSetCueBanner)
			MESSAGE_HANDLER(CB_SETCURSEL, OnSetCurSel)

			REFLECTED_COMMAND_CODE_HANDLER(CBN_CLOSEUP, OnReflectedCloseUp)
			REFLECTED_COMMAND_CODE_HANDLER(CBN_DROPDOWN, OnReflectedDropDown)
			REFLECTED_COMMAND_CODE_HANDLER(CBN_EDITCHANGE, OnReflectedEditChange)
			REFLECTED_COMMAND_CODE_HANDLER(CBN_EDITUPDATE, OnReflectedEditUpdate)
			REFLECTED_COMMAND_CODE_HANDLER(CBN_ERRSPACE, OnReflectedErrSpace)
			REFLECTED_COMMAND_CODE_HANDLER(CBN_SELCHANGE, OnReflectedSelChange)
			REFLECTED_COMMAND_CODE_HANDLER(CBN_SELENDCANCEL, OnReflectedSelEndCancel)
			REFLECTED_COMMAND_CODE_HANDLER(CBN_SELENDOK, OnReflectedSelEndOk)
			REFLECTED_COMMAND_CODE_HANDLER(CBN_SETFOCUS, OnReflectedSetFocus)

			COMMAND_CODE_HANDLER(EN_ALIGN_LTR_EC, OnAlign)
			COMMAND_CODE_HANDLER(EN_ALIGN_RTL_EC, OnAlign)
			COMMAND_CODE_HANDLER(EN_MAXTEXT, OnMaxText)
			COMMAND_CODE_HANDLER(EN_SETFOCUS, OnReflectedSetFocus)

			COMMAND_CODE_HANDLER(LBN_SETFOCUS, OnReflectedSetFocus)

			CHAIN_MSG_MAP(CComControl<ComboBox>)
			ALT_MSG_MAP(1)
			MESSAGE_HANDLER(WM_DESTROY, OnListBoxDestroy)
			MESSAGE_HANDLER(WM_HSCROLL, OnListBoxScroll)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnListBoxLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnListBoxLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnListBoxMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnListBoxMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnListBoxMouseWheel)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnListBoxMouseMove)
			MESSAGE_HANDLER(WM_MOUSEWHEEL, OnListBoxMouseWheel)
			MESSAGE_HANDLER(WM_PAINT, OnListBoxPaint)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnListBoxRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnListBoxRButtonUp)
			MESSAGE_HANDLER(WM_VSCROLL, OnListBoxScroll)
			ALT_MSG_MAP(2)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			INDEXED_MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_DESTROY, OnEditDestroy)
			INDEXED_MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			INDEXED_MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
			INDEXED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			INDEXED_MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			INDEXED_MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)
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
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Load to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Load and read directly from the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] pErrorLog The caller's \c IErrorLog implementation which will receive any errors
	///            that occur during property loading.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768206.aspx">IPersistPropertyBag::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog);
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Save to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Save and write directly into the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	/// \param[in] saveAllProperties Flag indicating whether the control should save all its properties
	///            (\c TRUE) or only those that have changed from the default value (\c FALSE).
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768207.aspx">IPersistPropertyBag::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties);
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
	/// \name Implementation of IComboBox
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AcceptNumbersOnly property</em>
	///
	/// Retrieves whether the contained edit control accepts all kind of text or only numbers. If set to
	/// \c VARIANT_TRUE, only numbers, otherwise all text is accepted.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AcceptNumbersOnly, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_AcceptNumbersOnly(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AcceptNumbersOnly property</em>
	///
	/// Sets whether the contained edit control accepts all kind of text or only numbers. If set to
	/// \c VARIANT_TRUE, only numbers, otherwise all text is accepted.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AcceptNumbersOnly, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_AcceptNumbersOnly(VARIANT_BOOL newValue);
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
	/// \brief <em>Retrieves the current setting of the \c AutoHorizontalScrolling property</em>
	///
	/// Retrieves whether the control scrolls automatically in horizontal direction, if the caret reaches the
	/// borders of the control's client area. If set to \c VARIANT_TRUE, the control scrolls automatically;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_AutoHorizontalScrolling, get_Style
	virtual HRESULT STDMETHODCALLTYPE get_AutoHorizontalScrolling(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoHorizontalScrolling property</em>
	///
	/// Sets whether the control scrolls automatically in horizontal direction, if the caret reaches the
	/// borders of the control's client area. If set to \c VARIANT_TRUE, the control scrolls automatically;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_AutoHorizontalScrolling, put_Style
	virtual HRESULT STDMETHODCALLTYPE put_AutoHorizontalScrolling(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property is ignored if the \c Style property
	///          is set to \c sDropDownList.
	///
	/// \sa put_BackColor, get_ForeColor, get_ListBackColor, get_Style
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property is ignored if the \c Style property
	///          is set to \c sDropDownList.
	///
	/// \sa get_BackColor, put_ForeColor, put_ListBackColor, put_Style
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
	/// \brief <em>Retrieves the current setting of the \c CharacterConversion property</em>
	///
	/// Retrieves the kind of conversion that is applied to characters that are typed into the control. Any
	/// of the values defined by the \c CharacterConversionConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_CharacterConversion, get_DoOEMConversion, get_Text,
	///       CBLCtlsLibU::CharacterConversionConstants
	/// \else
	///   \sa put_CharacterConversion, get_DoOEMConversion, get_Text,
	///       CBLCtlsLibA::CharacterConversionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CharacterConversion(CharacterConversionConstants* pValue);
	/// \brief <em>Sets the \c CharacterConversion property</em>
	///
	/// Sets the kind of conversion that is applied to characters that are typed into the control. Any
	/// of the values defined by the \c CharacterConversionConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_CharacterConversion, put_DoOEMConversion, put_Text,
	///       CBLCtlsLibU::CharacterConversionConstants
	/// \else
	///   \sa get_CharacterConversion, put_DoOEMConversion, put_Text,
	///       CBLCtlsLibA::CharacterConversionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_CharacterConversion(CharacterConversionConstants newValue);
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
	/// \brief <em>Retrieves the current setting of the \c ComboItems property</em>
	///
	/// Retrieves a collection object wrapping the control's combo box items.
	///
	/// \param[out] ppItems Receives the collection object's \c IComboBoxItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ComboBoxItems
	virtual HRESULT STDMETHODCALLTYPE get_ComboItems(IComboBoxItems** ppItems);
	/// \brief <em>Retrieves the current setting of the \c CueBanner property</em>
	///
	/// Retrieves the control's textual cue.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to an bug in Windows XP and Windows Server 2003, cue banners won't work on those
	///          systems if East Asian language and complex script support is installed.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa put_CueBanner, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_CueBanner(BSTR* pValue);
	/// \brief <em>Sets the \c CueBanner property</em>
	///
	/// Sets the control's textual cue.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to an bug in Windows XP and Windows Server 2003, cue banners won't work on those
	///          systems if East Asian language and complex script support is installed.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa get_CueBanner, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_CueBanner(BSTR newValue);
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
	/// \brief <em>Retrieves the current setting of the \c DoOEMConversion property</em>
	///
	/// Retrieves whether the control's text is converted from the Windows character set to the OEM character
	/// set and then back to the Windows character set. Such a conversion ensures proper character conversion
	/// when the application calls the \c CharToOem function to convert a Windows string in the control to
	/// OEM characters. This property is most useful if the control contains file names that will be used on
	/// file systems that do not support Unicode.\n
	/// If set to \c VARIANT_TRUE, the conversion is performed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DoOEMConversion, get_CharacterConversion, get_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647473.aspx">CharToOem</a>
	virtual HRESULT STDMETHODCALLTYPE get_DoOEMConversion(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DoOEMConversion property</em>
	///
	/// Sets whether the control's text is converted from the Windows character set to the OEM character
	/// set and then back to the Windows character set. Such a conversion ensures proper character conversion
	/// when the application calls the \c CharToOem function to convert a Windows string in the control to
	/// OEM characters. This property is most useful if the control contains file names that will be used on
	/// file systems that do not support Unicode.\n
	/// If set to \c VARIANT_TRUE, the conversion is performed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_DoOEMConversion, put_CharacterConversion, put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647473.aspx">CharToOem</a>
	virtual HRESULT STDMETHODCALLTYPE put_DoOEMConversion(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DragDropDownTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be placed over the drop-down button
	/// during a drag'n'drop operation before the drop-down list box control will be opened automatically.
	/// If set to 0, automatic drop-down is disabled. If set to -1, the system's double-click time,
	/// multiplied with 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragDropDownTime, get_RegisterForOLEDragDrop, get_ListDragScrollTimeBase,
	///     Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragDropDownTime(LONG* pValue);
	/// \brief <em>Sets the \c DragDropDownTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be placed over the drop-down button
	/// during a drag'n'drop operation before the drop-down list box control will be opened automatically.
	/// If set to 0, automatic drop-down is disabled. If set to -1, the system's double-click time,
	/// multiplied with 4, is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragDropDownTime, put_RegisterForOLEDragDrop, put_ListDragScrollTimeBase,
	///     Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragDropDownTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DropDownButtonObjectState property</em>
	///
	/// Retrieves the accessibility object state of the checkbox which is displayed if the \c Style property
	/// is set to \c sComboDropDownList or \c sDropDownList. For a list of possible object states see the
	/// <a href="https://msdn.microsoft.com/en-us/library/ms697270.aspx">MSDN article</a>.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Style, get_IsDroppedDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms697270.aspx">Accessibility Object State Constants</a>
	virtual HRESULT STDMETHODCALLTYPE get_DropDownButtonObjectState(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c DropDownKey property</em>
	///
	/// Retrieves the key that opens the drop-down window when pressed. Any of the values defined by the
	/// \c DropDownKeyConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DropDownKey, CBLCtlsLibU::DropDownKeyConstants
	/// \else
	///   \sa put_DropDownKey, CBLCtlsLibA::DropDownKeyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DropDownKey(DropDownKeyConstants* pValue);
	/// \brief <em>Sets the \c DropDownKey property</em>
	///
	/// Sets the key that opens the drop-down window when pressed. Any of the values defined by the
	/// \c DropDownKeyConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DropDownKey, CBLCtlsLibU::DropDownKeyConstants
	/// \else
	///   \sa get_DropDownKey, CBLCtlsLibA::DropDownKeyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DropDownKey(DropDownKeyConstants newValue);
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
	/// Retrieves the first combo box item, that is entirely located within the drop-down list box control's
	/// client area and therefore visible to the user.
	///
	/// \param[out] ppFirstItem Receives the first visible item's \c IComboBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_FirstVisibleItem, get_ComboItems
	virtual HRESULT STDMETHODCALLTYPE get_FirstVisibleItem(IComboBoxItem** ppFirstItem);
	/// \brief <em>Sets the \c FirstVisibleItem property</em>
	///
	/// Sets the first combo box item, that is entirely located within the drop-down list box control's
	/// client area and therefore visible to the user.
	///
	/// \param[in] pNewFirstItem The new first visible item's \c IComboBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstVisibleItem, get_ComboItems
	virtual HRESULT STDMETHODCALLTYPE putref_FirstVisibleItem(IComboBoxItem* pNewFirstItem);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the control's content.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont, put_ForeColor, get_ListForeColor,
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
	/// \sa get_Font, putref_Font, put_UseSystemFont, put_ForeColor, put_ListForeColor,
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
	/// \sa get_Font, put_Font, put_UseSystemFont, put_ForeColor, put_ListForeColor,
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
	/// \sa put_ForeColor, get_BackColor, get_ListForeColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ForeColor, put_BackColor, put_ListForeColor
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
	/// \sa put_HasStrings, get_OwnerDrawItems, ComboBoxItem::get_ItemData, ComboBoxItem::get_Text
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
	/// \sa get_HasStrings, put_OwnerDrawItems, ComboBoxItem::put_ItemData, ComboBoxItem::put_Text
	virtual HRESULT STDMETHODCALLTYPE put_HasStrings(VARIANT_BOOL newValue);
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
	/// \sa get_hWndEdit, get_hWndListBox, Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndEdit property</em>
	///
	/// Retrieves the contained edit control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWnd, get_hWndListBox, Raise_CreatedEditControlWindow, Raise_DestroyedEditControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWndEdit(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndListBox property</em>
	///
	/// Retrieves the drop-down list box control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWnd, get_hWndEdit, Raise_CreatedListBoxControlWindow, Raise_DestroyedListBoxControlWindow,
	///     Raise_ListDropDown, Raise_ListCloseUp
	virtual HRESULT STDMETHODCALLTYPE get_hWndListBox(OLE_HANDLE* pValue);
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
	/// \brief <em>Retrieves the current setting of the \c IsDroppedDown property</em>
	///
	/// Retrieves whether the drop-down list box control is currently displayed. If \c VARIANT_TRUE, the list
	/// box is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Style, get_DropDownButtonObjectState
	virtual HRESULT STDMETHODCALLTYPE get_IsDroppedDown(VARIANT_BOOL* pValue);
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
	///          is ignored. Use the \c IComboBoxItem::Height property instead.
	///
	/// \sa put_ItemHeight, get_OwnerDrawItems, get_IntegralHeight, ComboBoxItem::get_Height,
	///     get_SelectionFieldHeight
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
	///          is ignored. Use the \c IComboBoxItem::Height property instead.
	///
	/// \sa get_ItemHeight, put_OwnerDrawItems, put_IntegralHeight, ComboBoxItem::put_Height,
	///     put_SelectionFieldHeight
	virtual HRESULT STDMETHODCALLTYPE put_ItemHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ListAlwaysShowVerticalScrollBar property</em>
	///
	/// Retrieves whether the vertical scroll bar in the drop-down list box control is disabled instead of
	/// hidden if the control does not contain enough items to scroll. If set to \c VARIANT_TRUE, the scroll
	/// bar is disabled; otherwise it is hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListAlwaysShowVerticalScrollBar, get_ComboItems, get_ListScrollableWidth
	virtual HRESULT STDMETHODCALLTYPE get_ListAlwaysShowVerticalScrollBar(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ListAlwaysShowVerticalScrollBar property</em>
	///
	/// Sets whether the vertical scroll bar in the drop-down list box control is disabled instead of
	/// hidden if the control does not contain enough items to scroll. If set to \c VARIANT_TRUE, the scroll
	/// bar is disabled; otherwise it is hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_ListAlwaysShowVerticalScrollBar, get_ComboItems, put_ListScrollableWidth
	virtual HRESULT STDMETHODCALLTYPE put_ListAlwaysShowVerticalScrollBar(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ListBackColor property</em>
	///
	/// Retrieves the drop-down list box control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListBackColor, get_ListForeColor, get_BackColor, get_ListInsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE get_ListBackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ListBackColor property</em>
	///
	/// Sets the drop-down list box control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListBackColor, put_ListForeColor, put_BackColor, put_ListInsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE put_ListBackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c ListDragScrollTimeBase property</em>
	///
	/// Retrieves the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time, divided by 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListDragScrollTimeBase, get_RegisterForOLEDragDrop, Raise_ListOLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_ListDragScrollTimeBase(LONG* pValue);
	/// \brief <em>Sets the \c ListDragScrollTimeBase property</em>
	///
	/// Sets the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time divided by 4 is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListDragScrollTimeBase, put_RegisterForOLEDragDrop, Raise_ListOLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_ListDragScrollTimeBase(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ListForeColor property</em>
	///
	/// Retrieves the drop-down list box control's text color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListForeColor, get_ListBackColor, get_ForeColor, get_ListInsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE get_ListForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ListForeColor property</em>
	///
	/// Sets the drop-down list box control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListForeColor, put_ListBackColor, put_ForeColor, put_ListInsertMarkColor
	virtual HRESULT STDMETHODCALLTYPE put_ListForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c ListHeight property</em>
	///
	/// Retrieves the height in pixels of the drop-down list box. If set to -1, the default setting is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListHeight, get_ListWidth, get_MinVisibleItems
	virtual HRESULT STDMETHODCALLTYPE get_ListHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ListHeight property</em>
	///
	/// Sets the height in pixels of the drop-down list box. If set to -1, the default setting is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListHeight, put_ListWidth, put_MinVisibleItems
	virtual HRESULT STDMETHODCALLTYPE put_ListHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ListInsertMarkColor property</em>
	///
	/// Retrieves the color that is the control's insertion mark is drawn in.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListInsertMarkColor, GetListInsertMarkPosition, get_ListBackColor, get_ListForeColor
	virtual HRESULT STDMETHODCALLTYPE get_ListInsertMarkColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ListInsertMarkColor property</em>
	///
	/// Sets the color that is the control's insertion mark is drawn in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListInsertMarkColor, SetListInsertMarkPosition, put_ListBackColor, put_ListForeColor
	virtual HRESULT STDMETHODCALLTYPE put_ListInsertMarkColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c ListScrollableWidth property</em>
	///
	/// Retrieves the width in pixels, by which the drop-down list box can be scrolled horizontally. If the
	/// width of the control is greater than this value, a horizontal scroll bar is displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListScrollableWidth, get_ListAlwaysShowVerticalScrollBar, get_ListWidth
	virtual HRESULT STDMETHODCALLTYPE get_ListScrollableWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ListScrollableWidth property</em>
	///
	/// Sets the width in pixels, by which the drop-down list box can be scrolled horizontally. If the
	/// width of the control is greater than this value, a horizontal scroll bar is displayed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListScrollableWidth, put_ListAlwaysShowVerticalScrollBar, put_ListWidth
	virtual HRESULT STDMETHODCALLTYPE put_ListScrollableWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ListWidth property</em>
	///
	/// Retrieves the width in pixels of the drop-down list box.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The minimum width of the drop-down list box is the combo box width.
	///
	/// \sa put_ListWidth, get_ListHeight, get_ListScrollableWidth
	virtual HRESULT STDMETHODCALLTYPE get_ListWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ListWidth property</em>
	///
	/// Sets the width in pixels of the drop-down list box.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The minimum width of the drop-down list box is the combo box width.
	///
	/// \sa get_ListWidth, put_ListHeight, put_ListScrollableWidth
	virtual HRESULT STDMETHODCALLTYPE put_ListWidth(OLE_XSIZE_PIXELS newValue);
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
	/// \brief <em>Retrieves the current setting of the \c MaxTextLength property</em>
	///
	/// Retrieves the maximum number of characters, that the user can type into the control. If set to -1,
	/// the system's default setting is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Text, that is set through the \c Text property may exceed this limit.
	///
	/// \sa put_MaxTextLength, get_TextLength, get_Text, Raise_TruncatedText
	virtual HRESULT STDMETHODCALLTYPE get_MaxTextLength(LONG* pValue);
	/// \brief <em>Sets the \c MaxTextLength property</em>
	///
	/// Sets the maximum number of characters, that the user can type into the control. If set to -1,
	/// the system's default setting is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Text, that is set through the \c Text property may exceed this limit.
	///
	/// \sa get_MaxTextLength, get_TextLength, put_Text, Raise_TruncatedText
	virtual HRESULT STDMETHODCALLTYPE put_MaxTextLength(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MinVisibleItems property</em>
	///
	/// Retrieves the minimum number of visible items in the drop-down list box. The list box will be made
	/// large enough to display the specified number of items, even if the height specified by the
	/// \c ListHeight property is smaller.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c IntegralHeight property is set to \c VARIANT_FALSE.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa put_MinVisibleItems, get_ListHeight, get_IntegralHeight
	virtual HRESULT STDMETHODCALLTYPE get_MinVisibleItems(LONG* pValue);
	/// \brief <em>Sets the \c MinVisibleItems property</em>
	///
	/// Sets the minimum number of visible items in the drop-down list box. The list box will be made
	/// large enough to display the specified number of items, even if the height specified by the
	/// \c ListHeight property is smaller.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c IntegralHeight property is set to \c VARIANT_FALSE.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa get_MinVisibleItems, put_ListHeight, put_IntegralHeight
	virtual HRESULT STDMETHODCALLTYPE put_MinVisibleItems(LONG newValue);
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
	/// \brief <em>Retrieves the current setting of the \c SelectedItem property</em>
	///
	/// Retrieves the control's currently selected item.
	///
	/// \param[out] ppSelectedItem Receives the selected item's \c IComboBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_SelectedItem, ComboBoxItem::get_Selected, SelectItemByText, Raise_SelectionChanged
	virtual HRESULT STDMETHODCALLTYPE get_SelectedItem(IComboBoxItem** ppSelectedItem);
	/// \brief <em>Sets the \c SelectedItem property</em>
	///
	/// Sets the control's currently selected item.
	///
	/// \param[in] pNewSelectedItem The new selected item's \c IComboBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_SelectedItem, ComboBoxItem::get_Selected, SelectItemByText, Raise_SelectionChanged
	virtual HRESULT STDMETHODCALLTYPE putref_SelectedItem(IComboBoxItem* pNewSelectedItem);
	/// \brief <em>Retrieves the current setting of the \c SelectionFieldHeight property</em>
	///
	/// Retrieves the height of the part of the control that displays the currently selected item. If set to
	/// -1, the default setting is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_SelectionFieldHeight, get_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE get_SelectionFieldHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c SelectionFieldHeight property</em>
	///
	/// Sets the height of the part of the control that displays the currently selected item. If set to
	/// -1, the default setting is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_SelectionFieldHeight, put_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE put_SelectionFieldHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowDragImage property</em>
	///
	/// Retrieves whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible;
	/// otherwise hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowDragImage, get_SupportOLEDragImages, Raise_OLEDragMouseMove
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
	/// \sa get_ShowDragImage, put_SupportOLEDragImages, Raise_OLEDragMouseMove
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
	/// \sa put_Sorted, get_ComboItems, get_Locale, Raise_CompareItems
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
	/// \sa get_Sorted, get_ComboItems, put_Locale, Raise_CompareItems
	virtual HRESULT STDMETHODCALLTYPE put_Sorted(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Style property</em>
	///
	/// Retrieves which kind of user input is possible. Any of the values defined by the \c StyleConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Style, CBLCtlsLibU::StyleConstants
	/// \else
	///   \sa put_Style, CBLCtlsLibA::StyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Style(StyleConstants* pValue);
	/// \brief <em>Sets the \c Style property</em>
	///
	/// Sets which kind of user input is possible. Any of the values defined by the \c StyleConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_Style, CBLCtlsLibU::StyleConstants
	/// \else
	///   \sa get_Style, CBLCtlsLibA::StyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Style(StyleConstants newValue);
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
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
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
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
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
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the text currently displayed in the combo portion of the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, get_TextLength, get_MaxTextLength, get_AcceptNumbersOnly, get_CueBanner, get_ForeColor,
	///     get_Font, Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the text currently displayed in the combo portion of the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, get_TextLength, put_MaxTextLength, put_AcceptNumbersOnly, put_CueBanner, put_ForeColor,
	///     put_Font, Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c TextLength property</em>
	///
	/// Retrieves the length of the text specified by the \c Text property.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MaxTextLength, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_TextLength(LONG* pValue);
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

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Closes the drop-down list box control</em>
	///
	/// Closes the drop-down list box control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OpenDropDownWindow, Raise_ListCloseUp, get_hWndListBox
	virtual HRESULT STDMETHODCALLTYPE CloseDropDownWindow(void);
	/// \brief <em>Finds an item by its \c ItemData property</em>
	///
	/// Searches the combo box control for the first item that has the \c ItemData property set to the
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
	/// \sa FindItemByText, SelectItemByItemData, ComboBoxItem::get_ItemData
	virtual HRESULT STDMETHODCALLTYPE FindItemByItemData(LONG itemData, VARIANT startAfterItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IComboBoxItem** ppFoundItem = NULL);
	/// \brief <em>Finds an item by its \c ItemData property</em>
	///
	/// Searches the combo box control for the first item that has the \c ItemData property set to the
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
	/// \sa FindItemByItemData, SelectItemByText, ComboBoxItem::get_Text
	virtual HRESULT STDMETHODCALLTYPE FindItemByText(BSTR searchString, VARIANT_BOOL exactMatch = VARIANT_TRUE, VARIANT startAfterItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IComboBoxItem** ppFoundItem = NULL);
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
	///            insertion mark position. It must be relative to the drop-down list box control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the drop-down list box control's
	///            upper-left corner.
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified item. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppComboItem Receives the \c IComboBoxItem implementation of the item, at which
	///             the insertion mark should be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetListInsertMarkPosition, GetListInsertMarkPosition, CBLCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetListInsertMarkPosition, GetListInsertMarkPosition, CBLCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetClosestListInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, IComboBoxItem** ppComboItem);
	/// \brief <em>Retrieves the bounding rectangle of the drop-down button that is displayed if \c Style is set to \c sComboDropDownList or \c sDropDownList</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's upper-left corner) of the
	/// drop-down button which is displayed if the \c Style property is set to \c sComboDropDownList or
	/// \c sDropDownList.
	///
	/// \param[out] pLeft Receives the x-coordinate (in pixels) of the upper-left corner of the button's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pTop Receives the y-coordinate (in pixels) of the upper-left corner of the button's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pRight Receives the x-coordinate (in pixels) of the lower-right corner of the button's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pBottom Receives the y-coordinate (in pixels) of the lower-right corner of the button's
	///             bounding rectangle relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Style, GetSelectionFieldRectangle, GetDroppedStateRectangle
	virtual HRESULT STDMETHODCALLTYPE GetDropDownButtonRectangle(OLE_XPOS_PIXELS* pLeft = NULL, OLE_YPOS_PIXELS* pTop = NULL, OLE_XPOS_PIXELS* pRight = NULL, OLE_YPOS_PIXELS* pBottom = NULL);
	/// \brief <em>Retrieves the bounding rectangle of the control including the drop-down list box</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the screen's upper-left corner) of the
	/// control when it is dropped down, i. e. including the drop-down list box.
	///
	/// \param[out] pLeft Receives the x-coordinate (in pixels) of the upper-left corner of the overall
	///             bounding rectangle relative to the screen's upper-left corner.
	/// \param[out] pTop Receives the y-coordinate (in pixels) of the upper-left corner of the overall
	///             bounding rectangle relative to the screen's upper-left corner.
	/// \param[out] pRight Receives the x-coordinate (in pixels) of the lower-right corner of the overall
	///             bounding rectangle relative to the screen's upper-left corner.
	/// \param[out] pBottom Receives the y-coordinate (in pixels) of the lower-right corner of the overall
	///             bounding rectangle relative to the screen's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_IsDroppedDown, GetDropDownButtonRectangle, GetSelectionFieldRectangle
	virtual HRESULT STDMETHODCALLTYPE GetDroppedStateRectangle(OLE_XPOS_PIXELS* pLeft = NULL, OLE_YPOS_PIXELS* pTop = NULL, OLE_XPOS_PIXELS* pRight = NULL, OLE_YPOS_PIXELS* pBottom = NULL);
	/// \brief <em>Retrieves the position of the control's insertion mark</em>
	///
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified item. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppComboItem Receives the \c IComboBoxItem implementation of the item, at which the
	///             insertion mark is displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetListInsertMarkPosition, GetClosestListInsertMarkPosition, GetListInsertMarkRectangle,
	///       CBLCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetListInsertMarkPosition, GetClosestListInsertMarkPosition, GetListInsertMarkRectangle,
	///       CBLCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetListInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, IComboBoxItem** ppComboItem);
	/// \brief <em>Retrieves the bounding rectangle of the control's insertion mark</em>
	///
	/// \param[out] pXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///             relative to the drop-down list box control's upper-left corner.
	/// \param[out] pYTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///             relative to the drop-down list box control's upper-left corner.
	/// \param[out] pXRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///             relative to the drop-down list box control's upper-left corner.
	/// \param[out] pYBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///             relative to the drop-down list box control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetListInsertMarkPosition, SetListInsertMarkPosition
	virtual HRESULT STDMETHODCALLTYPE GetListInsertMarkRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Retrieves the contained edit control's current selection's start and end</em>
	///
	/// Retrieves the zero-based character indices of the contained edit control's current selection's start
	/// and end.
	///
	/// \param[out] pSelectionStart The zero-based index of the character at which the selection starts.
	/// \param[out] pSelectionEnd The zero-based index of the first unselected character after the end of the
	///             selection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetSelection
	virtual HRESULT STDMETHODCALLTYPE GetSelection(LONG* pSelectionStart = NULL, LONG* pSelectionEnd = NULL);
	/// \brief <em>Retrieves the bounding rectangle of the part of the control that displays the currently selected item</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's upper-left corner) of the
	/// part of the control that displays the currently selected item.
	///
	/// \param[out] pLeft Receives the x-coordinate (in pixels) of the upper-left corner of the item field's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pTop Receives the y-coordinate (in pixels) of the upper-left corner of the item field's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pRight Receives the x-coordinate (in pixels) of the lower-right corner of the item
	///             field's bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pBottom Receives the y-coordinate (in pixels) of the lower-right corner of the item
	///             field's bounding rectangle relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Style, GetDropDownButtonRectangle, GetDroppedStateRectangle
	virtual HRESULT STDMETHODCALLTYPE GetSelectionFieldRectangle(OLE_XPOS_PIXELS* pLeft = NULL, OLE_YPOS_PIXELS* pTop = NULL, OLE_XPOS_PIXELS* pRight = NULL, OLE_YPOS_PIXELS* pBottom = NULL);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pHitTestDetails Receives a value specifying the exact part of the control the specified
	///             point lies in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ListHitTest, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa ListHitTest, CBLCtlsLibA::HitTestConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the drop-down list box control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[out] pHitTestDetails Receives a value specifying the exact part of the drop-down list box
	///             control the specified point lies in. Any of the values defined by the \c HitTestConstants
	///             enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa HitTest, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa HitTest, CBLCtlsLibA::HitTestConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE ListHitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IComboBoxItem** ppHitItem);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Opens the drop-down list box control</em>
	///
	/// Opens the drop-down list box control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CloseDropDownWindow, Raise_ListDropDown, get_hWndListBox
	virtual HRESULT STDMETHODCALLTYPE OpenDropDownWindow(void);
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
	/// \sa ComboBoxItems::Add
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
	/// Searches the combo box control for the first item that has the \c ItemData property set to the
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
	/// \sa SelectItemByText, FindItemByItemData, ComboBoxItem::get_ItemData, putref_SelectedItem
	virtual HRESULT STDMETHODCALLTYPE SelectItemByItemData(LONG itemData, VARIANT startAfterItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IComboBoxItem** ppFoundItem = NULL);
	/// \brief <em>Finds and selects an item by its \c ItemData property</em>
	///
	/// Searches the combo box control for the first item that has the \c ItemData property set to the
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
	/// \sa SelectItemByItemData, FindItemByText, ComboBoxItem::get_Text, putref_SelectedItem
	virtual HRESULT STDMETHODCALLTYPE SelectItemByText(BSTR searchString, VARIANT_BOOL exactMatch = VARIANT_TRUE, VARIANT startAfterItem = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IComboBoxItem** ppFoundItem = NULL);
	/// \brief <em>Sets the position of the control's insertion mark</em>
	///
	/// \param[in] relativePosition The insertion mark's position relative to the specified item. Any of
	///            the values defined by the \c InsertMarkPositionConstants enumeration is valid.
	/// \param[in] pComboItem The \c IComboBoxItem implementation of the item, at which the insertion
	///            mark will be displayed. If set to \c NULL, the insertion mark will be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetListInsertMarkPosition, GetClosestListInsertMarkPosition, get_ListInsertMarkColor,
	///       get_RegisterForOLEDragDrop, CBLCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa GetListInsertMarkPosition, GetClosestListInsertMarkPosition, get_ListInsertMarkColor,
	///       get_RegisterForOLEDragDrop, CBLCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetListInsertMarkPosition(InsertMarkPositionConstants relativePosition, IComboBoxItem* pComboItem);
	/// \brief <em>Sets the contained edit control's selection's start and end</em>
	///
	/// Sets the zero-based character indices of the contained edit control's selection's start and end.
	///
	/// \param[in] selectionStart The zero-based index of the character at which the selection starts. If set
	///            to -1, the current selection is cleared.
	/// \param[in] selectionEnd The zero-based index of the first unselected character after the end of the
	///            selection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks To select all text in the  contained edit control, set \c selectionStart to 0 and
	///          \c selectionEnd to -1.
	///
	/// \sa GetSelection
	virtual HRESULT STDMETHODCALLTYPE SetSelection(LONG selectionStart, LONG selectionEnd);
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
	/// \sa CommonProperties, StringProperties,
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
	LRESULT OnContextMenu(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to configure the control window and to raise the \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CTLCOLOREDIT handler</em>
	///
	/// Will be called if the control is asked to configure the specified device context for drawing the
	/// contained edit control, i. e. to setup the colors and brushes.
	/// We use this handler for the \c BackColor and \c ForeColor properties.
	///
	/// \sa get_BackColor, get_ForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761691.aspx">WM_CTLCOLOREDIT</a>
	LRESULT OnCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CTLCOLORLISTBOX handler</em>
	///
	/// Will be called if the control is asked to configure the specified device context for drawing the
	/// drop-down list box control, i. e. to setup the colors and brushes.
	/// We use this handler for the \c ListBackColor and \c ListForeColor properties.
	///
	/// \sa get_ListBackColor, get_ListForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761360.aspx">WM_CTLCOLORLISTBOX</a>
	LRESULT OnCtlColorListBox(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, OnEditDestroy, OnListBoxDestroy, Raise_DestroyedControlWindow,
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
	LRESULT OnKeyDown(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
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
	LRESULT OnLButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnMButtonDown, OnRButtonDown, OnXButtonDown, OnListBoxLButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, OnListBoxLButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, OnListBoxMButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, OnListBoxMButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
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
	LRESULT OnMouseHover(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, OnMouseWheel, OnListBoxMouseMove, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnMouseMove, OnListBoxMouseWheel, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnMouseWheel(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
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
	LRESULT OnRButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, OnListBoxRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, OnListBoxRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
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
	/// \brief <em>\c WM_SETTEXT handler</em>
	///
	/// Will be called if the control's text shall be changed.
	/// We use this handler to support data-binding for the \c Text property.
	///
	/// \sa OnReflectedEditChange, get_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632644.aspx">WM_SETTEXT</a>
	LRESULT OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
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
	/// \sa Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_COMPAREITEM handler</em>
	///
	/// Will be called if the control has the \c LBS_SORT style and is owner-drawn and an item is inserted.
	/// The handler of this message must specify the relative order of the two specified items.
	/// We use this handler to raise the \c CompareItems event.
	///
	/// \sa Raise_CompareItems,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775921.aspx">WM_COMPAREITEM</a>
	LRESULT OnReflectedCompareItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DRAWITEM handler</em>
	///
	/// Will be called if the control's parent window is asked to draw a combo box item.
	/// We use this handler to raise the \c OwnerDrawItem event.
	///
	/// \sa OnReflectedMeasureItem, get_OwnerDrawItems, Raise_OwnerDrawItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775923.aspx">WM_DRAWITEM</a>
	LRESULT OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MEASUREITEM handler</em>
	///
	/// Will be called if the control's parent window is asked for a combo box item's height.
	/// We use this handler to raise the \c MeasureItem event.
	///
	/// \sa OnReflectedDrawItem, get_OwnerDrawItems, Raise_MeasureItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775925.aspx">WM_MEASUREITEM</a>
	LRESULT OnReflectedMeasureItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c CB_ADDSTRING handler</em>
	///
	/// Will be called if an item shall be inserted. We use this handler to raise the \c InsertingItem
	/// and \c InsertedItem events.
	///
	/// \sa OnInsertString, Raise_InsertingItem, Raise_InsertedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775828.aspx">CB_ADDSTRING</a>
	LRESULT OnAddString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c CB_DELETESTRING handler</em>
	///
	/// Will be called if an item shall be removed. We use this handler to raise the \c FreeItemData,
	/// \c RemovingItem and \c RemovedItem events.
	///
	/// \sa OnResetContent, Raise_FreeItemData, Raise_RemovingItem, Raise_RemovedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775830.aspx">CB_DELETESTRING</a>
	LRESULT OnDeleteString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c CB_INSERTSTRING handler</em>
	///
	/// Will be called if an item shall be inserted. We use this handler to raise the \c InsertingItem
	/// and \c InsertedItem events.
	///
	/// \sa OnAddString, Raise_InsertingItem, Raise_InsertedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775875.aspx">CB_INSERTSTRING</a>
	LRESULT OnInsertString(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c CB_RESETCONTENT handler</em>
	///
	/// Will be called if all items shall be removed. We use this handler to raise the \c FreeItemData,
	/// \c RemovingItem and \c RemovedItem events.
	///
	/// \sa OnDeleteString, Raise_RemovingItem, Raise_RemovedItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775878.aspx">CB_RESETCONTENT</a>
	LRESULT OnResetContent(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c CB_SETCUEBANNER handler</em>
	///
	/// Will be called if the control's textual cue shall be set.
	/// We use this handler to synchronize our settings.
	///
	/// \sa get_CueBanner,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775897.aspx">CB_SETCUEBANNER</a>
	LRESULT OnSetCueBanner(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c CB_SETCURSEL handler</em>
	///
	/// Will be called if the control's currently selected item shall be set.
	/// We use this handler to raise the \c SelectionChanged event.
	///
	/// \sa get_SelectedItem, Raise_SelectionChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775899.aspx">CB_SETCURSEL</a>
	LRESULT OnSetCurSel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the drop-down list box control is being destroyed.
	/// We use this handler to raise the \c DestroyedListBoxControlWindow event.
	///
	/// \sa OnDestroy, OnEditDestroy, Raise_DestroyedListBoxControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnListBoxDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the drop-down list box control's client area.
	/// We use this handler to raise the \c ListMouseDown event.
	///
	/// \sa OnListBoxMButtonDown, OnListBoxRButtonDown, OnLButtonDown, Raise_ListMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnListBoxLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the drop-down list box control's client area.
	/// We use this handler to raise the \c ListMouseUp event.
	///
	/// \sa OnListBoxMButtonUp, OnListBoxRButtonUp, OnLButtonUp, Raise_ListMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnListBoxLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the drop-down list box control's client area.
	/// We use this handler to raise the \c ListMouseDown event.
	///
	/// \sa OnListBoxLButtonDown, OnListBoxRButtonDown, OnMButtonDown, Raise_ListMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnListBoxMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the drop-down list box control's client area.
	/// We use this handler to raise the \c ListMouseUp event.
	///
	/// \sa OnListBoxLButtonUp, OnListBoxRButtonUp, OnMButtonUp, Raise_ListMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnListBoxMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the drop-down
	/// list box control's client area.
	/// We use this handler to raise the \c ListMouseMove event.
	///
	/// \sa OnListBoxLButtonDown, OnListBoxLButtonUp, OnListBoxMouseWheel, OnMouseMove, Raise_ListMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnListBoxMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// drop-down list box control's client area.
	/// We use this handler to raise the \c ListMouseWheel event.
	///
	/// \sa OnListBoxMouseMove, OnMouseWheel, Raise_ListMouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnListBoxMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT handler</em>
	///
	/// Will be called if the drop-down list box control needs to be drawn.
	/// We use this handler to draw the insertion mark.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>
	LRESULT OnListBoxPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the drop-down list box control's client area.
	/// We use this handler mainly to raise the \c ListMouseDown event.
	///
	/// \sa OnListBoxLButtonDown, OnListBoxMButtonDown, OnRButtonDown, Raise_ListMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnListBoxRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the drop-down list box control's client area.
	/// We use this handler to raise the \c ListMouseUp event.
	///
	/// \sa OnListBoxLButtonUp, OnListBoxMButtonUp, OnRButtonUp, Raise_ListMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnListBoxRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_*SCROLL handler</em>
	///
	/// Will be called if the drop-down list box control shall be scrolled.
	/// We use this handler to redraw the insertion mark.
	///
	/// \sa ListBoxDrawInsertionMark,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651283.aspx">WM_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651284.aspx">WM_VSCROLL</a>
	LRESULT OnListBoxScroll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the contained edit control is being destroyed.
	/// We use this handler to raise the \c DestroyedEditControlWindow event.
	///
	/// \sa OnDestroy, OnListBoxDestroy, Raise_DestroyedEditControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnEditDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c CBN_CLOSEUP handler</em>
	///
	/// Will be called if the control's parent window is notified, that the drop-down list box control
	/// is about to be displayed.
	/// We use this handler to raise the \c ListCloseUp event.
	///
	/// \sa OnReflectedDropDown, Raise_ListCloseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775806.aspx">CBN_CLOSEUP</a>
	LRESULT OnReflectedCloseUp(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c CBN_DROPDOWN handler</em>
	///
	/// Will be called if the control's parent window is notified, that the drop-down list box control
	/// is about to be hidden.
	/// We use this handler to raise the \c ListDropDown event.
	///
	/// \sa OnReflectedCloseUp, Raise_ListDropDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775810.aspx">CBN_DROPDOWN</a>
	LRESULT OnReflectedDropDown(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c CBN_EDITCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the contained edit control's content
	/// has changed.
	/// We use this handler to raise the \c TextChanged event.
	///
	/// \sa OnSetText, get_Text, Raise_TextChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775812.aspx">CBN_EDITCHANGE</a>
	LRESULT OnReflectedEditChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c CBN_EDITUPDATE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the contained edit control's content
	/// is about to be drawn.
	/// We use this handler to raise the \c BeforeDrawText event.
	///
	/// \sa Raise_BeforeDrawText,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775814.aspx">CBN_EDITUPDATE</a>
	LRESULT OnReflectedEditUpdate(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c CBN_ERRSPACE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control couldn't allocate enough
	/// memory to meet a specific request.
	/// We use this handler to raise the \c OutOfMemory event.
	///
	/// \sa Raise_OutOfMemory,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775816.aspx">CBN_ERRSPACE</a>
	LRESULT OnReflectedErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c CBN_SELCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the selection in the control has
	/// changed.
	/// We use this handler to raise the \c SelectionChanged event.
	///
	/// \sa OnReflectedSelEndOk, OnReflectedSelEndCancel, Raise_SelectionChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775821.aspx">CBN_SELCHANGE</a>
	LRESULT OnReflectedSelChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c CBN_SELENDCANCEL handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user has canceled changing the
	/// selection in the control.
	/// We use this handler to raise the \c SelectionCanceled event.
	///
	/// \sa OnReflectedSelEndOk, OnReflectedSelChange, Raise_SelectionCanceled,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775823.aspx">CBN_SELENDCANCEL</a>
	LRESULT OnReflectedSelEndCancel(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c CBN_SELENDOK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user has finished changing the
	/// selection in the control.
	/// We use this handler to raise the \c SelectionChanging event.
	///
	/// \sa OnReflectedSelEndCancel, OnReflectedSelChange, Raise_SelectionChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775825.aspx">CBN_SELENDOK</a>
	LRESULT OnReflectedSelEndOk(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c CBN_SETFOCUS handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has gained the keyboard
	/// focus.
	/// We use this handler to initialize IME.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775827.aspx">CBN_SETFOCUS</a>
	LRESULT OnReflectedSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_ALIGN_*_EC handler</em>
	///
	/// Will be called if the combo box is notified, that the user has changed the contained edit control's
	/// writing direction.
	/// We use this handler to raise the \c WritingDirectionChanged event.
	///
	/// \sa get_RightToLeft, Raise_WritingDirectionChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672108.aspx">EN_ALIGN_LTR_EC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672109.aspx">EN_ALIGN_RTL_EC</a>
	LRESULT OnAlign(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_MAXTEXT handler</em>
	///
	/// Will be called if the combo box is notified, that the text, that was entered into the contained edit
	/// control, got truncated.
	/// We use this handler to raise the \c TruncatedText event.
	///
	/// \sa Raise_TruncatedText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672115.aspx">EN_MAXTEXT</a>
	LRESULT OnMaxText(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c BeforeDrawText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_BeforeDrawText, CBLCtlsLibU::_IComboBoxEvents::BeforeDrawText
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_BeforeDrawText, CBLCtlsLibA::_IComboBoxEvents::BeforeDrawText
	/// \endif
	inline HRESULT Raise_BeforeDrawText(void);
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
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_Click, CBLCtlsLibU::_IComboBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_Click, CBLCtlsLibA::_IComboBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	///          currently being inserted. In this case all properties of \c IComboBoxItem except \c Index
	///          and \c ItemData will fail and the \c ItemData property is read-only.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_CompareItems, CBLCtlsLibU::_IComboBoxEvents::CompareItems,
	///       CBLCtlsLibU::CompareResultConstants, get_OwnerDrawItems, get_Sorted, get_Locale,
	///       ComboBoxItem::get_Index, ComboBoxItem::get_ItemData
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_CompareItems, CBLCtlsLibA::_IComboBoxEvents::CompareItems,
	///       CBLCtlsLibA::CompareResultConstants, get_OwnerDrawItems, get_Sorted, get_Locale,
	///       ComboBoxItem::get_Index, ComboBoxItem::get_ItemData
	/// \endif
	inline HRESULT Raise_CompareItems(IComboBoxItem* pFirstItem, IComboBoxItem* pSecondItem, LONG locale, CompareResultConstants* pResult);
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
	/// \param[in] hitTestDetails The part of the control that the menu's proposed position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the control from
	///                showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ContextMenu, CBLCtlsLibU::_IComboBoxEvents::ContextMenu,
	///       Raise_RClick, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ContextMenu, CBLCtlsLibA::_IComboBoxEvents::ContextMenu,
	///       Raise_RClick, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu);
	/// \brief <em>Raises the \c CreatedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_CreatedEditControlWindow,
	///       CBLCtlsLibU::_IComboBoxEvents::CreatedEditControlWindow,
	///       Raise_DestroyedEditControlWindow, get_hWndEdit
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_CreatedEditControlWindow,
	///       CBLCtlsLibA::_IComboBoxEvents::CreatedEditControlWindow,
	///       Raise_DestroyedEditControlWindow, get_hWndEdit
	/// \endif
	inline HRESULT Raise_CreatedEditControlWindow(HWND hWndEdit);
	/// \brief <em>Raises the \c CreatedListBoxControlWindow event</em>
	///
	/// \param[in] hWndListBox The list box control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_CreatedListBoxControlWindow,
	///       CBLCtlsLibU::_IComboBoxEvents::CreatedListBoxControlWindow,
	///       Raise_DestroyedListBoxControlWindow, Raise_ListDropDown, get_hWndListBox
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_CreatedListBoxControlWindow,
	///       CBLCtlsLibA::_IComboBoxEvents::CreatedListBoxControlWindow,
	///       Raise_DestroyedListBoxControlWindow, Raise_ListDropDown, get_hWndListBox
	/// \endif
	inline HRESULT Raise_CreatedListBoxControlWindow(HWND hWndListBox);
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
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_DblClick, CBLCtlsLibU::_IComboBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_DblClick, CBLCtlsLibA::_IComboBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_DestroyedControlWindow,
	///       CBLCtlsLibU::_IComboBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow, get_hWnd
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_DestroyedControlWindow,
	///       CBLCtlsLibA::_IComboBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow, get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(HWND hWnd);
	/// \brief <em>Raises the \c DestroyedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_DestroyedEditControlWindow,
	///       CBLCtlsLibU::_IComboBoxEvents::DestroyedEditControlWindow,
	///       Raise_CreatedEditControlWindow, get_hWndEdit
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_DestroyedEditControlWindow,
	///       CBLCtlsLibA::_IComboBoxEvents::DestroyedEditControlWindow,
	///       Raise_CreatedEditControlWindow, get_hWndEdit
	/// \endif
	inline HRESULT Raise_DestroyedEditControlWindow(HWND hWndEdit);
	/// \brief <em>Raises the \c DestroyedListBoxControlWindow event</em>
	///
	/// \param[in] hWndListBox The list box control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_DestroyedListBoxControlWindow,
	///       CBLCtlsLibU::_IComboBoxEvents::DestroyedListBoxControlWindow,
	///       Raise_CreatedListBoxControlWindow, Raise_ListCloseUp, get_hWndListBox
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_DestroyedListBoxControlWindow,
	///       CBLCtlsLibA::_IComboBoxEvents::DestroyedListBoxControlWindow,
	///       Raise_CreatedListBoxControlWindow, Raise_ListCloseUp, get_hWndListBox
	/// \endif
	inline HRESULT Raise_DestroyedListBoxControlWindow(HWND hWndListBox);
	/// \brief <em>Raises the \c FreeItemData event</em>
	///
	/// \param[in] pComboItem The item whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_FreeItemData, CBLCtlsLibU::_IComboBoxEvents::FreeItemData,
	///       Raise_RemovingItem, Raise_RemovedItem, ComboBoxItem::put_ItemData
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_FreeItemData, CBLCtlsLibA::_IComboBoxEvents::FreeItemData,
	///       Raise_RemovingItem, Raise_RemovedItem, ComboBoxItem::put_ItemData
	/// \endif
	inline HRESULT Raise_FreeItemData(IComboBoxItem* pComboItem);
	/// \brief <em>Raises the \c InsertedItem event</em>
	///
	/// \param[in] pComboItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_InsertedItem, CBLCtlsLibU::_IComboBoxEvents::InsertedItem,
	///       Raise_InsertingItem, Raise_RemovedItem
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_InsertedItem, CBLCtlsLibA::_IComboBoxEvents::InsertedItem,
	///       Raise_InsertingItem, Raise_RemovedItem
	/// \endif
	inline HRESULT Raise_InsertedItem(IComboBoxItem* pComboItem);
	/// \brief <em>Raises the \c InsertingItem event</em>
	///
	/// \param[in] pComboItem The item being inserted. If the combo box is sorted, this object's \c Index
	///            property may be wrong.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort insertion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_InsertingItem, CBLCtlsLibU::_IComboBoxEvents::InsertingItem,
	///       Raise_InsertedItem, Raise_RemovingItem
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_InsertingItem, CBLCtlsLibA::_IComboBoxEvents::InsertingItem,
	///       Raise_InsertedItem, Raise_RemovingItem
	/// \endif
	inline HRESULT Raise_InsertingItem(IVirtualComboBoxItem* pComboItem, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ItemMouseEnter event</em>
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
	///            position lies in. Most of the values defined by the \c HitTestConstants enumeration are
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ItemMouseEnter, CBLCtlsLibU::_IComboBoxEvents::ItemMouseEnter,
	///       Raise_ItemMouseLeave, Raise_ListMouseMove, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ItemMouseEnter, CBLCtlsLibA::_IComboBoxEvents::ItemMouseEnter,
	///       Raise_ItemMouseLeave, Raise_ListMouseMove, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemMouseEnter(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c ItemMouseLeave event</em>
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
	///            position lies in. Most of the values defined by the \c HitTestConstants enumeration are
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ItemMouseLeave, CBLCtlsLibU::_IComboBoxEvents::ItemMouseLeave,
	///       Raise_ItemMouseEnter, Raise_ListMouseMove, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ItemMouseLeave, CBLCtlsLibA::_IComboBoxEvents::ItemMouseLeave,
	///       Raise_ItemMouseEnter, Raise_ListMouseMove, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ItemMouseLeave(IComboBoxItem* pComboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	///   \sa Proxy_IComboBoxEvents::Fire_KeyDown, CBLCtlsLibU::_IComboBoxEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_KeyDown, CBLCtlsLibA::_IComboBoxEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress
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
	///   \sa Proxy_IComboBoxEvents::Fire_KeyPress, CBLCtlsLibU::_IComboBoxEvents::KeyPress, Raise_KeyDown,
	///       Raise_KeyUp
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_KeyPress, CBLCtlsLibA::_IComboBoxEvents::KeyPress, Raise_KeyDown,
	///       Raise_KeyUp
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
	///   \sa Proxy_IComboBoxEvents::Fire_KeyUp, CBLCtlsLibU::_IComboBoxEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_KeyUp, CBLCtlsLibA::_IComboBoxEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c ListCloseUp event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ListCloseUp, CBLCtlsLibU::_IComboBoxEvents::ListCloseUp,
	///       Raise_ListDropDown, CloseDropDownWindow, get_hWndListBox
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListCloseUp, CBLCtlsLibA::_IComboBoxEvents::ListCloseUp,
	///       Raise_ListDropDown, CloseDropDownWindow, get_hWndListBox
	/// \endif
	inline HRESULT Raise_ListCloseUp(void);
	/// \brief <em>Raises the \c ListDropDown event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ListDropDown, CBLCtlsLibU::_IComboBoxEvents::ListDropDown,
	///       Raise_ListCloseUp, OpenDropDownWindow, get_hWndListBox
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListDropDown, CBLCtlsLibA::_IComboBoxEvents::ListDropDown,
	///       Raise_ListCloseUp, OpenDropDownWindow, get_hWndListBox
	/// \endif
	inline HRESULT Raise_ListDropDown(void);
	/// \brief <em>Raises the \c ListMouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ListMouseDown, CBLCtlsLibU::_IComboBoxEvents::ListMouseDown,
	///       Raise_ListMouseUp, Raise_MouseDown
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListMouseDown, CBLCtlsLibA::_IComboBoxEvents::ListMouseDown,
	///       Raise_ListMouseUp, Raise_MouseDown
	/// \endif
	inline HRESULT Raise_ListMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ListMouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ListMouseMove, CBLCtlsLibU::_IComboBoxEvents::ListMouseMove,
	///       Raise_ListMouseDown, Raise_ListMouseUp, Raise_ListMouseWheel, Raise_MouseMove
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListMouseMove, CBLCtlsLibA::_IComboBoxEvents::ListMouseMove,
	///       Raise_ListMouseDown, Raise_ListMouseUp, Raise_ListMouseWheel, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_ListMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ListMouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ListMouseUp, CBLCtlsLibU::_IComboBoxEvents::ListMouseUp,
	///       Raise_ListMouseDown, Raise_MouseUp
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListMouseUp, CBLCtlsLibA::_IComboBoxEvents::ListMouseUp,
	///       Raise_ListMouseDown, Raise_MouseUp
	/// \endif
	inline HRESULT Raise_ListMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ListMouseWheel event</em>
	///
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ListMouseWheel, CBLCtlsLibU::_IComboBoxEvents::ListMouseWheel,
	///       Raise_ListMouseMove, Raise_MouseWheel, CBLCtlsLibU::ExtendedMouseButtonConstants,
	///       CBLCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListMouseWheel, CBLCtlsLibA::_IComboBoxEvents::ListMouseWheel,
	///       Raise_ListMouseMove, Raise_MouseWheel, CBLCtlsLibA::ExtendedMouseButtonConstants,
	///       CBLCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_ListMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c ListOLEDragDrop event</em>
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
	///   \sa Proxy_IComboBoxEvents::Fire_ListOLEDragDrop, CBLCtlsLibU::_IComboBoxEvents::ListOLEDragDrop,
	///       Raise_ListOLEDragEnter, Raise_ListOLEDragMouseMove, Raise_ListOLEDragLeave, Raise_ListMouseUp,
	///       Raise_OLEDragDrop, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListOLEDragDrop, CBLCtlsLibA::_IComboBoxEvents::ListOLEDragDrop,
	///       Raise_ListOLEDragEnter, Raise_ListOLEDragMouseMove, Raise_ListOLEDragLeave, Raise_ListMouseUp,
	///       Raise_OLEDragDrop, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_ListOLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c ListOLEDragEnter event</em>
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
	///   \sa Proxy_IComboBoxEvents::Fire_ListOLEDragEnter, CBLCtlsLibU::_IComboBoxEvents::ListOLEDragEnter,
	///       Raise_ListOLEDragMouseMove, Raise_ListOLEDragLeave, Raise_ListOLEDragDrop, Raise_OLEDragEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListOLEDragEnter, CBLCtlsLibA::_IComboBoxEvents::ListOLEDragEnter,
	///       Raise_ListOLEDragMouseMove, Raise_ListOLEDragLeave, Raise_ListOLEDragDrop, Raise_OLEDragEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_ListOLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c ListOLEDragLeave event</em>
	///
	/// \param[in] fakedLeave If \c FALSE, the method releases the cached data object; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ListOLEDragLeave, CBLCtlsLibU::_IComboBoxEvents::ListOLEDragLeave,
	///       Raise_ListOLEDragEnter, Raise_ListOLEDragMouseMove, Raise_ListOLEDragDrop, Raise_OLEDragLeave
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListOLEDragLeave, CBLCtlsLibA::_IComboBoxEvents::ListOLEDragLeave,
	///       Raise_ListOLEDragEnter, Raise_ListOLEDragMouseMove, Raise_ListOLEDragDrop, Raise_OLEDragLeave
	/// \endif
	inline HRESULT Raise_ListOLEDragLeave(BOOL fakedLeave);
	/// \brief <em>Raises the \c ListOLEDragMouseMove event</em>
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
	///   \sa Proxy_IComboBoxEvents::Fire_ListOLEDragMouseMove,
	///       CBLCtlsLibU::_IComboBoxEvents::ListOLEDragMouseMove, Raise_ListOLEDragEnter,
	///       Raise_ListOLEDragLeave, Raise_ListOLEDragDrop, Raise_ListMouseMove, Raise_OLEDragMouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ListOLEDragMouseMove,
	///       CBLCtlsLibA::_IComboBoxEvents::ListOLEDragMouseMove, Raise_ListOLEDragEnter,
	///       Raise_ListOLEDragLeave, Raise_ListOLEDragDrop, Raise_ListMouseMove, Raise_OLEDragMouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_ListOLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition);
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
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MClick, CBLCtlsLibU::_IComboBoxEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MClick, CBLCtlsLibA::_IComboBoxEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MDblClick, CBLCtlsLibU::_IComboBoxEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MDblClick, CBLCtlsLibA::_IComboBoxEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MeasureItem event</em>
	///
	/// \param[in] pComboItem The item for which the size is required. If the event is raised for the
	///            selection field, this parameter will be \c NULL.
	/// \param[out] pItemHeight Must be set to the item's height in pixels by the client app.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MeasureItem, CBLCtlsLibU::_IComboBoxEvents::MeasureItem,
	///       Raise_OwnerDrawItem, ComboBox::put_ItemHeight, ComboBoxItem::put_Height,
	///       ComboBox::get_OwnerDrawItems
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MeasureItem, CBLCtlsLibA::_IComboBoxEvents::MeasureItem,
	///       Raise_OwnerDrawItem, ComboBox::put_ItemHeight, ComboBoxItem::put_Height,
	///       ComboBox::get_OwnerDrawItems
	/// \endif
	inline HRESULT Raise_MeasureItem(IComboBoxItem* pListItem, OLE_YSIZE_PIXELS* pItemHeight);
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
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MouseDown, CBLCtlsLibU::_IComboBoxEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick, Raise_ListMouseDown,
	///       CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MouseDown, CBLCtlsLibA::_IComboBoxEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick, Raise_ListMouseDown,
	///       CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MouseEnter, CBLCtlsLibU::_IComboBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MouseEnter, CBLCtlsLibA::_IComboBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MouseHover, CBLCtlsLibU::_IComboBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime,
	///       CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MouseHover, CBLCtlsLibA::_IComboBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime,
	///       CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MouseLeave, CBLCtlsLibU::_IComboBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MouseLeave, CBLCtlsLibA::_IComboBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MouseMove, CBLCtlsLibU::_IComboBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       Raise_ListMouseMove, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MouseMove, CBLCtlsLibA::_IComboBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       Raise_ListMouseMove, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MouseUp, CBLCtlsLibU::_IComboBoxEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick, Raise_ListMouseUp,
	///       CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MouseUp, CBLCtlsLibA::_IComboBoxEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick, Raise_ListMouseUp,
	///       CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_MouseWheel, CBLCtlsLibU::_IComboBoxEvents::MouseWheel,
	///       Raise_MouseMove, Raise_ListMouseWheel, CBLCtlsLibU::ExtendedMouseButtonConstants,
	///       CBLCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_MouseWheel, CBLCtlsLibA::_IComboBoxEvents::MouseWheel,
	///       Raise_MouseMove, Raise_ListMouseWheel, CBLCtlsLibA::ExtendedMouseButtonConstants,
	///       CBLCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta, HitTestConstants hitTestDetails);
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
	///   \sa Proxy_IComboBoxEvents::Fire_OLEDragDrop, CBLCtlsLibU::_IComboBoxEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       Raise_ListOLEDragDrop, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_OLEDragDrop, CBLCtlsLibA::_IComboBoxEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       Raise_ListOLEDragDrop, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
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
	///   \sa Proxy_IComboBoxEvents::Fire_OLEDragEnter, CBLCtlsLibU::_IComboBoxEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       Raise_ListOLEDragEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_OLEDragEnter, CBLCtlsLibA::_IComboBoxEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       Raise_ListOLEDragEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \param[in] fakedLeave If \c FALSE, the method releases the cached data object; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_OLEDragLeave, CBLCtlsLibU::_IComboBoxEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave,
	///       Raise_ListOLEDragLeave
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_OLEDragLeave, CBLCtlsLibA::_IComboBoxEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave,
	///       Raise_ListOLEDragLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(BOOL fakedLeave);
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
	///   \sa Proxy_IComboBoxEvents::Fire_OLEDragMouseMove, CBLCtlsLibU::_IComboBoxEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       Raise_ListOLEDragMouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_OLEDragMouseMove, CBLCtlsLibA::_IComboBoxEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       Raise_ListOLEDragMouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OutOfMemory event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_OutOfMemory, CBLCtlsLibU::_IComboBoxEvents::OutOfMemory
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_OutOfMemory, CBLCtlsLibA::_IComboBoxEvents::OutOfMemory
	/// \endif
	inline HRESULT Raise_OutOfMemory(void);
	/// \brief <em>Raises the \c OwnerDrawItem event</em>
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
	///   \sa Proxy_IComboBoxEvents::Fire_OwnerDrawItem, CBLCtlsLibU::_IComboBoxEvents::OwnerDrawItem,
	///       Raise_MeasureItem, ComboBox::put_ItemHeight, ComboBoxItem::put_Height,
	///       ComboBox::get_OwnerDrawItems, CBLCtlsLibU::RECTANGLE, CBLCtlsLibU::OwnerDrawActionConstants,
	///       CBLCtlsLibU::OwnerDrawItemStateConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_OwnerDrawItem, CBLCtlsLibA::_IComboBoxEvents::OwnerDrawItem,
	///       Raise_MeasureItem, ComboBox::put_ItemHeight, ComboBoxItem::put_Height,
	///       ComboBox::get_OwnerDrawItems, CBLCtlsLibA::RECTANGLE, CBLCtlsLibA::OwnerDrawActionConstants,
	///       CBLCtlsLibA::OwnerDrawItemStateConstants
	/// \endif
	inline HRESULT Raise_OwnerDrawItem(IComboBoxItem* pComboItem, OwnerDrawActionConstants requiredAction, OwnerDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle);
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
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_RClick, CBLCtlsLibU::_IComboBoxEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_RClick, CBLCtlsLibA::_IComboBoxEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_RDblClick, CBLCtlsLibU::_IComboBoxEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_RDblClick, CBLCtlsLibA::_IComboBoxEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick, CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_RecreatedControlWindow,
	///       CBLCtlsLibU::_IComboBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow, get_hWnd
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_RecreatedControlWindow,
	///       CBLCtlsLibA::_IComboBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow, get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(HWND hWnd);
	/// \brief <em>Raises the \c RemovedItem event</em>
	///
	/// \param[in] pComboItem The removed item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_RemovedItem, CBLCtlsLibU::_IComboBoxEvents::RemovedItem,
	///       Raise_RemovingItem, Raise_InsertedItem
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_RemovedItem, CBLCtlsLibA::_IComboBoxEvents::RemovedItem,
	///       Raise_RemovingItem, Raise_InsertedItem
	/// \endif
	inline HRESULT Raise_RemovedItem(IVirtualComboBoxItem* pComboItem);
	/// \brief <em>Raises the \c RemovingItem event</em>
	///
	/// \param[in] pComboItem The item being removed.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_RemovingItem, CBLCtlsLibU::_IComboBoxEvents::RemovingItem,
	///       Raise_RemovedItem, Raise_InsertingItem
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_RemovingItem, CBLCtlsLibA::_IComboBoxEvents::RemovingItem,
	///       Raise_RemovedItem, Raise_InsertingItem
	/// \endif
	inline HRESULT Raise_RemovingItem(IComboBoxItem* pComboItem, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_ResizedControlWindow,
	///       CBLCtlsLibU::_IComboBoxEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_ResizedControlWindow,
	///       CBLCtlsLibA::_IComboBoxEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c SelectionCanceled event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_SelectionCanceled, CBLCtlsLibU::_IComboBoxEvents::SelectionCanceled,
	///       get_SelectedItem, Raise_SelectionChanging, Raise_SelectionChanged
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_SelectionCanceled, CBLCtlsLibA::_IComboBoxEvents::SelectionCanceled,
	///       get_SelectedItem, Raise_SelectionChanging, Raise_SelectionChanged
	/// \endif
	inline HRESULT Raise_SelectionCanceled(void);
	/// \brief <em>Raises the \c SelectionChanged event</em>
	///
	/// \param[in] previousSelectedItem The previous selected item.
	/// \param[in] newSelectedItem The new selected item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_SelectionChanged, CBLCtlsLibU::_IComboBoxEvents::SelectionChanged,
	///       get_SelectedItem, Raise_SelectionCanceled, Raise_SelectionChanging
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_SelectionChanged, CBLCtlsLibA::_IComboBoxEvents::SelectionChanged,
	///       get_SelectedItem, Raise_SelectionCanceled, Raise_SelectionChanging
	/// \endif
	inline HRESULT Raise_SelectionChanged(int previousSelectedItem, int newSelectedItem);
	/// \brief <em>Raises the \c SelectionChanging event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_SelectionChanging, CBLCtlsLibU::_IComboBoxEvents::SelectionChanging,
	///       get_SelectedItem, Raise_SelectionChanged, Raise_SelectionCanceled
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_SelectionChanging, CBLCtlsLibA::_IComboBoxEvents::SelectionChanging,
	///       get_SelectedItem, Raise_SelectionChanged, Raise_SelectionCanceled
	/// \endif
	inline HRESULT Raise_SelectionChanging(void);
	/// \brief <em>Raises the \c TextChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_TextChanged, CBLCtlsLibU::_IComboBoxEvents::TextChanged
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_TextChanged, CBLCtlsLibA::_IComboBoxEvents::TextChanged
	/// \endif
	inline HRESULT Raise_TextChanged(void);
	/// \brief <em>Raises the \c TruncatedText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_TruncatedText, CBLCtlsLibU::_IComboBoxEvents::TruncatedText
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_TruncatedText, CBLCtlsLibA::_IComboBoxEvents::TruncatedText
	/// \endif
	inline HRESULT Raise_TruncatedText(void);
	/// \brief <em>Raises the \c WritingDirectionChanged event</em>
	///
	/// \param[in] newWritingDirection The control's new writing direction. Any of the values defined by the
	///            \c WritingDirectionConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to limitations of Microsoft Windows, this event is not raised if the writing direction
	///          is changed using the control's default context menu.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_WritingDirectionChanged,
	///       CBLCtlsLibU::_IComboBoxEvents::WritingDirectionChanged, CBLCtlsLibU::WritingDirectionConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_WritingDirectionChanged,
	///       CBLCtlsLibA::_IComboBoxEvents::WritingDirectionChanged, CBLCtlsLibA::WritingDirectionConstants
	/// \endif
	inline HRESULT Raise_WritingDirectionChanged(WritingDirectionConstants newWritingDirection);
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
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_XClick, CBLCtlsLibU::_IComboBoxEvents::XClick, Raise_XDblClick,
	///       Raise_Click, Raise_MClick, Raise_RClick, CBLCtlsLibU::ExtendedMouseButtonConstants,
	///       CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_XClick, CBLCtlsLibA::_IComboBoxEvents::XClick, Raise_XDblClick,
	///       Raise_Click, Raise_MClick, Raise_RClick, CBLCtlsLibA::ExtendedMouseButtonConstants,
	///       CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IComboBoxEvents::Fire_XDblClick, CBLCtlsLibU::_IComboBoxEvents::XDblClick, Raise_XClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_RDblClick, CBLCtlsLibU::ExtendedMouseButtonConstants,
	///       CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IComboBoxEvents::Fire_XDblClick, CBLCtlsLibA::_IComboBoxEvents::XDblClick, Raise_XClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_RDblClick, CBLCtlsLibA::ExtendedMouseButtonConstants,
	///       CBLCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
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
	/// Sends \c WM_* and \c CB_* messages to the control window to make it match the current property
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
	/// \return The requested item's index if successful; -2 otherwise.
	///
	/// \sa itemIDs, ItemIndexToID, ComboBoxItems::get_Item, ComboBoxItems::Remove
	int IDToItemIndex(LONG ID);
	/// \brief <em>Translates an item's index to its unique ID</em>
	///
	/// \param[in] itemIndex The index of the item whose unique ID is requested.
	///
	/// \return The requested item's unique ID if successful; -1 otherwise.
	///
	/// \sa itemIDs, IDToItemIndex, ComboBoxItem::get_ID
	LONG ItemIndexToID(int itemIndex);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the item that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[out] pFlags A bit field of \c HitTestConstants flags, that holds further details about the
	///             drop-down list box control's part below the specified point.
	/// \param[in] autoScroll If \c TRUE and the specified point lies directly above or below the list box,
	///            the control is scrolled vertically by 1 line.
	///
	/// \return The "hit" item's zero-based index.
	///
	/// \if UNICODE
	///   \sa ListHitTest, CBLCtlsLibU::HitTestConstants
	/// \else
	///   \sa ListHitTest, CBLCtlsLibA::HitTestConstants
	/// \endif
	int ListBoxHitTest(LONG x, LONG y, __in HitTestConstants* pFlags, BOOL autoScroll = FALSE);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);
	/// \brief <em>Auto-scrolls the control</em>
	///
	/// \sa OnTimer, DragDropStatus::AutoScrolling
	void ListBoxAutoScroll(void);
	/// \brief <em>Draws the insertion mark</em>
	///
	/// This method is called by \c OnPaint to draw the drop-down list box control's insertion mark.
	///
	/// \param[in] targetDC The device context to be used for drawing.
	///
	/// \sa OnPaint
	void ListBoxDrawInsertionMark(CDCHandle targetDC);

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
	///     ComboBoxItem::put_Text, Raise_InsertingItem, Raise_InsertedItem
	void EnterSilentItemInsertionSection(void);
	/// \brief <em>Decrements the \c silentItemInsertions flag</em>
	///
	/// \sa EnterSilentItemInsertionSection, Flags::silentItemInsertions, LeaveSilentItemDeletionSection,
	///     ComboBoxItem::put_Text, Raise_InsertingItem, Raise_InsertedItem
	void LeaveSilentItemInsertionSection(void);
	/// \brief <em>Increments the \c silentItemDeletions flag</em>
	///
	/// \sa LeaveSilentItemDeletionSection, Flags::silentItemDeletions, EnterSilentItemInsertionSection,
	///     ComboBoxItem::put_Text, Raise_RemovingItem, Raise_RemovedItem
	void EnterSilentItemDeletionSection(void);
	/// \brief <em>Decrements the \c silentItemDeletions flag</em>
	///
	/// \sa EnterSilentItemDeletionSection, Flags::silentItemDeletions, LeaveSilentItemInsertionSection,
	///     ComboBoxItem::put_Text, Raise_RemovingItem, Raise_RemovedItem
	void LeaveSilentItemDeletionSection(void);
	/// \brief <em>Changes an item's unique ID</em>
	///
	/// \param[in] itemIndex The zero-based index of the item for which to change the ID.
	/// \param[in] itemID The item's new unique ID.
	///
	/// \sa ComboBoxItem::put_Text
	void SetItemID(int itemIndex, LONG itemID);
	/// \brief <em>Decrements \c lastItemID by 1</em>
	///
	/// \sa lastItemID, ComboBoxItem::put_Text
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
			CComObject< PropertyNotifySinkImpl<ComboBox> >* pPropertyNotifySink;

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
			HRESULT InitializePropertyWatcher(ComboBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ComboBox> >::CreateInstance(&pPropertyNotifySink);
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
			CComObject< PropertyNotifySinkImpl<ComboBox> >* pPropertyNotifySink;

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
			HRESULT InitializePropertyWatcher(ComboBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ComboBox> >::CreateInstance(&pPropertyNotifySink);
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
		/// The default width (in pixels) of the border around the drop-down list box control's client area,
		/// that's sensitive for auto-scrolling during a drag'n'drop operation. If the mouse cursor's position
		/// lies within this area during a drag'n'drop operation, the drop-down list box control will be
		/// auto-scrolled.
		///
		/// \sa listDragScrollTimeBase, Raise_ListOLEDragMouseMove
		static const int DRAGSCROLLZONEWIDTH = 16;

		/// \brief <em>Holds the \c AcceptNumbersOnly property's setting</em>
		///
		/// \sa get_AcceptNumbersOnly, put_AcceptNumbersOnly
		UINT acceptNumbersOnly : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c AutoHorizontalScrolling property's setting</em>
		///
		/// \sa get_AutoHorizontalScrolling, put_AutoHorizontalScrolling
		UINT autoHorizontalScrolling : 1;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c CharacterConversion property's setting</em>
		///
		/// \sa get_CharacterConversion, put_CharacterConversion
		CharacterConversionConstants characterConversion;
		/// \brief <em>Holds the \c CueBanner property's setting</em>
		///
		/// \sa get_CueBanner, put_CueBanner
		CComBSTR cueBanner;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DontRedraw property's setting</em>
		///
		/// \sa get_DontRedraw, put_DontRedraw
		UINT dontRedraw : 1;
		/// \brief <em>Holds the \c DoOEMConversion property's setting</em>
		///
		/// \sa get_DoOEMConversion, put_DoOEMConversion
		UINT doOEMConversion : 1;
		/// \brief <em>Holds the \c DragDropDownTime property's setting</em>
		///
		/// \sa get_DragDropDownTime, put_DragDropDownTime
		long dragDropDownTime;
		/// \brief <em>Holds the \c DropDownKey property's setting</em>
		///
		/// \sa get_DropDownKey, put_DropDownKey
		DropDownKeyConstants dropDownKey;
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
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c IMEMode property's setting</em>
		///
		/// \sa get_IMEMode, put_IMEMode
		IMEModeConstants IMEMode;
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
		/// \brief <em>Holds the \c ListAlwaysShowVerticalScrollBar property's setting</em>
		///
		/// \sa get_ListAlwaysShowVerticalScrollBar, put_ListAlwaysShowVerticalScrollBar
		UINT listAlwaysShowVerticalScrollBar : 1;
		/// \brief <em>Holds the \c ListBackColor property's setting</em>
		///
		/// \sa get_ListBackColor, put_ListBackColor
		OLE_COLOR listBackColor;
		/// \brief <em>Holds the \c ListDragScrollTimeBase property's setting</em>
		///
		/// \sa get_ListDragScrollTimeBase, put_ListDragScrollTimeBase
		long listDragScrollTimeBase;
		/// \brief <em>Holds the \c ListForeColor property's setting</em>
		///
		/// \sa get_ListForeColor, put_ListForeColor
		OLE_COLOR listForeColor;
		/// \brief <em>Holds the \c ListHeight property's setting</em>
		///
		/// \sa get_ListHeight, put_ListHeight
		OLE_YSIZE_PIXELS listHeight;
		/// \brief <em>Holds the \c ListInsertMarkColor property's setting</em>
		///
		/// \sa get_ListInsertMarkColor, put_ListInsertMarkColor
		OLE_COLOR listInsertMarkColor;
		/// \brief <em>Holds the \c ListScrollableWidth property's setting</em>
		///
		/// \sa get_ListScrollableWidth, put_ListScrollableWidth
		OLE_XSIZE_PIXELS listScrollableWidth;
		/// \brief <em>Holds the \c ListWidth property's setting</em>
		///
		/// \sa get_ListWidth, put_ListWidth
		OLE_XSIZE_PIXELS listWidth;
		/// \brief <em>Holds the \c Locale property's setting</em>
		///
		/// \sa get_Locale, put_Locale
		long locale;
		/// \brief <em>Holds the \c MaxTextLength property's setting</em>
		///
		/// \sa get_MaxTextLength, put_MaxTextLength
		long maxTextLength;
		/// \brief <em>Holds the \c MinVisibleItems property's setting</em>
		///
		/// \sa get_MinVisibleItems, put_MinVisibleItems
		long minVisibleItems;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c OwnerDrawItems property's setting</em>
		///
		/// \sa get_OwnerDrawItems, put_OwnerDrawItems
		OwnerDrawItemsConstants ownerDrawItems;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		UINT registerForOLEDragDrop : 1;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c SelectionFieldHeight property's setting</em>
		///
		/// \sa get_SelectionFieldHeight, put_SelectionFieldHeight
		OLE_YSIZE_PIXELS selectionFieldHeight;
		/// \brief <em>Holds the value that the \c SelectionFieldHeight property was set to</em>
		///
		/// \sa get_SelectionFieldHeight, put_SelectionFieldHeight
		OLE_YSIZE_PIXELS setSelectionFieldHeight;
		/// \brief <em>Holds the \c Sorted property's setting</em>
		///
		/// \sa get_Sorted, put_Sorted
		UINT sorted : 1;
		/// \brief <em>Holds the \c Style property's setting</em>
		///
		/// \sa get_Style, put_Style
		StyleConstants style;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c Text property's setting</em>
		///
		/// \sa get_Text, put_Text
		CComBSTR text;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;

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
			acceptNumbersOnly = FALSE;
			appearance = aDefault;
			autoHorizontalScrolling = TRUE;
			backColor = 0x80000000 | COLOR_WINDOW;
			borderStyle = bsNone;
			characterConversion = ccNone;
			cueBanner = L"";
			disabledEvents = static_cast<DisabledEventsConstants>(deMouseEvents | deListBoxMouseEvents | deClickEvents | deKeyboardEvents | deItemInsertionEvents | deItemDeletionEvents | deFreeItemData | deBeforeDrawText | deTextChangedEvents | deSelectionChangingEvents);
			dontRedraw = FALSE;
			doOEMConversion = FALSE;
			dragDropDownTime = -1;
			dropDownKey = ddkF4;
			enabled = TRUE;
			foreColor = 0x80000000 | COLOR_WINDOWTEXT;
			hasStrings = TRUE;
			hoverTime = -1;
			IMEMode = imeInherit;
			integralHeight = TRUE;
			itemHeight = -1;
			setItemHeight = -1;
			listAlwaysShowVerticalScrollBar = FALSE;
			listBackColor = backColor;
			listDragScrollTimeBase = -1;
			listForeColor = foreColor;
			listHeight = -1;
			listInsertMarkColor = RGB(0, 0, 0);
			listScrollableWidth = 0;
			listWidth = 0;
			locale = LOCALE_USER_DEFAULT;
			maxTextLength = -1;
			minVisibleItems = 30;
			mousePointer = mpDefault;
			ownerDrawItems = odiDontOwnerDraw;
			processContextMenuKeys = TRUE;
			registerForOLEDragDrop = FALSE;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			selectionFieldHeight = -1;
			setSelectionFieldHeight = -1;
			sorted = FALSE;
			style = sDropDownList;
			supportOLEDragImages = TRUE;
			text = L"";
			useSystemFont = TRUE;
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
		/// \c brief <em>If \c TRUE, \c get_Text will use \c CB_GETLBTEXT instead of \c GetWindowText</em>
		///
		/// When we raise the \c TextChanged event in response to the \c CBN_SELCHANGE notification, the text
		/// returned by \c GetWindowText is still the one of the previously selected item. Therefore we set
		/// this flag and in \c get_Text we use \c CB_GETLBTEXT to retrieve the text.
		///
		/// \sa get_Text, Raise_TextChanged, OnReflectedSelChange,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775862.aspx">CB_GETLBTEXT</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775821.aspx">CBN_SELCHANGE</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms633520.aspx">GetWindowText</a>
		UINT windowTextIsOutdated : 1;
		/// \brief <em>Holds the height of the contained list box window that it had on the last \c WM_WINDOWPOSCHANGED message</em>
		///
		/// This member is used with comctl32.dll version 5.x to work around a problem with high dpi settings,
		/// where the contained list box has a size of 0.
		///
		/// \sa OnWindowPosChanged, Raise_ListDropDown
		int lastListBoxHeight;

		Flags()
		{
			dontRecreate = FALSE;
			silentItemInsertions = 0;
			silentItemDeletions = 0;
			uiActivationPending = FALSE;
			usingThemes = FALSE;
			windowTextIsOutdated = FALSE;
			lastListBoxHeight = 0;
		}
	} flags;

	//////////////////////////////////////////////////////////////////////
	/// \name Item management
	///
	//@{
	#ifdef USE_STL
		/// \brief <em>A list of all item IDs</em>
		///
		/// Holds the unique IDs of all combo box items in the control. The item's index in the control equals
		/// its index in the vector.
		///
		/// \sa OnAddString, OnInsertString, OnResetContent, OnDeleteString, ComboBoxItem::get_ID
		std::vector<LONG> itemIDs;
	#else
		/// \brief <em>A list of all items</em>
		///
		/// Holds the unique IDs of all combo box items in the control. The item's index in the control equals
		/// its index in the vector.
		///
		/// \sa OnAddString, OnInsertString, OnResetContent, OnDeleteString, ComboBoxItem::get_ID
		CAtlArray<LONG> itemIDs;
	#endif
	/// \brief <em>Retrieves a new unique item ID at each call</em>
	///
	/// \return A new unique item ID.
	///
	/// \sa itemIDs, ComboBoxItem::get_ID
	LONG GetNewItemID(void);
	///@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the index of the currently selected item</em>
	int currentSelectedItem;
	/// \brief <em>Holds the index of the item below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deListBoxMouseEvents being set.
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

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
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

	/// \brief <em>Holds the contained edit control's mouse status</em>
	MouseStatus mouseStatus_Edit;
	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus_ComboBox;
	/// \brief <em>Holds the drop-down list box control's mouse status</em>
	MouseStatus	mouseStatus_ListBox;

	/// \brief <em>Holds the position of an insertion mark</em>
	typedef struct InsertMark
	{
		/// \brief <em>If set to \c TRUE, the insertion mark won't be drawn</em>
		///
		/// \sa OnListBoxScroll, ListBoxDrawInsertionMark
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
	} InsertMark;
	/// \brief <em>Holds the position of the drop-down list box control's insertion mark</em>
	InsertMark listInsertMark;

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>Indicates whether the drag image is visible or hidden</em>
		///
		/// If this value is 0, the drag image is visible; otherwise not.
		///
		/// \sa get_ShowDragImage, put_ShowDragImage, ShowDragImage, HideDragImage,
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
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>If \c TRUE, the current OLE drag'n'drop operation affects the drop-down list box control</em>
		UINT isOverListBox : 1;
		/// \brief <em>If \c TRUE, the \c ID_DRAGDROPDOWN timer is ticking</em>
		UINT dropDownTimerIsRunning : 1;
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

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pDropTargetHelper = NULL;
			dragImageIsHidden = 1;
			lastDropTarget = -1;
			isOverListBox = FALSE;
			dropDownTimerIsRunning = FALSE;
			drop_pDataObject = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
			}
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			dragImageIsHidden = 1;
			lastDropTarget = -1;
			isOverListBox = FALSE;
			dropDownTimerIsRunning = FALSE;
			drop_pDataObject = NULL;
		}

		/// \brief <em>Decrements the \c dragImageIsHidden flag</em>
		///
		/// \sa dragImageIsHidden, HideDragImage, IsDragImageVisible
		void ShowDragImage(void)
		{
			if(pDropTargetHelper) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					pDropTargetHelper->Show(TRUE);
				}
			}
		}

		/// \brief <em>Increments the \c dragImageIsHidden flag</em>
		///
		/// \sa dragImageIsHidden, ShowDragImage, IsDragImageVisible
		void HideDragImage(void)
		{
			if(pDropTargetHelper) {
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

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa ListOLEDragLeaveOrDrop, OLEDragEnter
		HRESULT ListOLEDragEnter(void)
		{
			isOverListBox = TRUE;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa ListOLEDragEnter, OLEDragLeaveOrDrop
		void ListOLEDragLeaveOrDrop(void)
		{
			isOverListBox = FALSE;
			autoScrolling.Reset();
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop, ListOLEDragEnter
		HRESULT OLEDragEnter(void)
		{
			lastDropTarget = -1;
			isOverListBox = FALSE;
			dropDownTimerIsRunning = FALSE;
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter, ListOLEDragLeaveOrDrop
		void OLEDragLeaveOrDrop(void)
		{
			lastDropTarget = -1;
		}
	} dragDropStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;
		/// \brief <em>The ID of the timer that is used to auto-scroll the drop-down list box control window during drag'n'drop</em>
		static const UINT_PTR ID_DRAGSCROLL = 13;
		/// \brief <em>The ID of the timer that is used to open the drop-down list box control automatically during drag'n'drop</em>
		static const UINT_PTR ID_DRAGDROPDOWN = 14;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
	} timers;

	/// \brief <em>The last unique ID assigned to an item</em>
	///
	/// \sa DecrementNextItemID, GetNewItemID
	LONG lastItemID;

private:
};     // ComboBox

OBJECT_ENTRY_AUTO(__uuidof(ComboBox), ComboBox)