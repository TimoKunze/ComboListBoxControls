//////////////////////////////////////////////////////////////////////
/// \class ListBoxItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing list box item</em>
///
/// This class is a wrapper around a list box item that - unlike an item wrapped by
/// \c VirtualListBoxItem - really exists in the control.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IListBoxItem, VirtualListBoxItem, ListBoxItems, ListBox
/// \else
///   \sa CBLCtlsLibA::IListBoxItem, VirtualListBoxItem, ListBoxItems, ListBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IListBoxItemEvents_CP.h"
#include "helpers.h"
#include "ListBox.h"


class ATL_NO_VTABLE ListBoxItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListBoxItem, &CLSID_ListBoxItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListBoxItem>,
    public Proxy_IListBoxItemEvents<ListBoxItem>, 
    #ifdef UNICODE
    	public IDispatchImpl<IListBoxItem, &IID_IListBoxItem, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListBoxItem, &IID_IListBoxItem, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ListBox;
	friend class ListBoxItems;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTBOXITEM)

		BEGIN_COM_MAP(ListBoxItem)
			COM_INTERFACE_ENTRY(IListBoxItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListBoxItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IListBoxItemEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
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
	/// \name Implementation of IListBoxItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Anchor property</em>
	///
	/// Retrieves whether the item is the control's anchor item, i. e. it's the item with which
	/// range-selection begins. If it is the anchor item, this property is set to \c VARIANT_TRUE;
	/// otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Caret, get_Selected, ListBox::get_MultiSelect, ListBox::get_AnchorItem
	virtual HRESULT STDMETHODCALLTYPE get_Anchor(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the item is the control's caret item, i. e. it has the focus. If it is the
	/// caret item, this property is set to \c VARIANT_TRUE; otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Anchor, get_Selected, ListBox::get_CaretItem
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Height property</em>
	///
	/// Retrieves the item's height in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c OwnerDrawItems property is not set to \c odiOwnerDrawVariableHeight, this property
	///          is read-only. Use the \c IListBox::ItemHeight property instead.
	///
	/// \sa put_Height, GetRectangle, ListBox::get_OwnerDrawItems, ListBox::get_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE get_Height(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c Height property</em>
	///
	/// Sets the item's height in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c OwnerDrawItems property is not set to \c odiOwnerDrawVariableHeight, this property
	///          is read-only. Use the \c IListBox::ItemHeight property instead.
	///
	/// \sa get_Height, GetRectangle, ListBox::put_OwnerDrawItems, ListBox::put_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE put_Height(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks An item's ID will never change.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Index, CBLCtlsLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_Index, CBLCtlsLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing items changes other items' indexes, the index is the best
	///          (and fastest) option to identify an index.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ID, CBLCtlsLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_ID, CBLCtlsLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemData property</em>
	///
	/// Retrieves the \c LONG value associated with the item. Use this property to associate any data
	/// with the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ItemData, ListBox::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE get_ItemData(LONG* pValue);
	/// \brief <em>Sets the \c ItemData property</em>
	///
	/// Sets the \c LONG value associated with the item. Use this property to associate any data
	/// with the item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ItemData, ListBox::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE put_ItemData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the item is drawn as a selected item, i. e. whether its background is
	/// highlighted. If this property is set to \c VARIANT_TRUE, the item is highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Selected, ListBox::get_MultiSelect, ListBox::DeselectItems, ListBox::SelectItems, get_Anchor,
	///     get_Caret, ListBox::get_SelectedItem
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Selected property</em>
	///
	/// Sets whether the item is drawn as a selected item, i. e. whether its background is highlighted.
	/// If this property is set to \c VARIANT_TRUE, the item is highlighted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Selected, ListBox::get_MultiSelect, ListBox::DeselectItems, ListBox::SelectItems,
	///     ListBox::putref_SelectedItem, ListBox::putref_CaretItem
	virtual HRESULT STDMETHODCALLTYPE put_Selected(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, ListBox::get_ProcessTabs
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the item's text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, ListBox::put_ProcessTabs
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);

	/// \brief <em>Retrieves an image list containing the item's drag image</em>
	///
	/// Retrieves the handle to an image list containing a bitmap that can be used to visualize
	/// dragging of this item.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] phImageList The image list containing the drag image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the image list.
	///
	/// \sa ListBox::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	/// \brief <em>Retrieves the bounding rectangle of either the item or a part of it</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of either the
	/// item or a part of it.
	///
	/// \param[in] rectangleType The rectangle to retrieve. Any of the values defined by the
	///            \c ItemRectangleTypeConstants enumeration is valid.
	/// \param[out] PXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
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
	/// \if UNICODE
	///   \sa get_Height, CBLCtlsLibU::ItemRectangleTypeConstants
	/// \else
	///   \sa get_Height, CBLCtlsLibA::ItemRectangleTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(ItemRectangleTypeConstants /*rectangleType*/, OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given item</em>
	///
	/// Attaches this object to a given item, so that the item's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] itemIndex The item to attach to.
	/// \param[in] itemData Specifies the value that will be returned by the \c ItemData property, if the
	///            specified item index is -1.
	///
	/// \sa Detach, get_ItemData
	void Attach(int itemIndex, ULONG_PTR itemData = NULL);
	/// \brief <em>Detaches this object from an item</em>
	///
	/// Detaches this object from the item it currently wraps, so that it doesn't wrap any item anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	HRESULT SaveState(VirtualListBoxItem* pTarget);
	/// \brief <em>Sets the owner of this item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerLBox
	void SetOwner(__in_opt ListBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ListBox object that owns this item</em>
		///
		/// \sa SetOwner
		ListBox* pOwnerLBox;
		/// \brief <em>The index of the item wrapped by this object</em>
		int itemIndex;
		/// \brief <em>Specifies the value that will be returned by the \c ItemData property, if the specified item index is -1</em>
		///
		/// \sa get_ItemData
		ULONG_PTR itemData;

		Properties()
		{
			pOwnerLBox = NULL;
			itemIndex = -2;
			itemData = NULL;
		}

		~Properties();

		/// \brief <em>Retrieves the owning list box' window handle</em>
		///
		/// \return The window handle of the list box that contains this item.
		///
		/// \sa pOwnerLBox
		HWND GetLBoxHWnd(void);
	} properties;
};     // ListBoxItem

OBJECT_ENTRY_AUTO(__uuidof(ListBoxItem), ListBoxItem)