//////////////////////////////////////////////////////////////////////
/// \class ComboBoxItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing combo item</em>
///
/// This class is a wrapper around a combo box item that - unlike an item wrapped by
/// \c VirtualComboBoxItem - really exists in the control.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IComboBoxItem, VirtualComboBoxItem, ComboBoxItems, ComboBox
/// \else
///   \sa CBLCtlsLibA::IComboBoxItem, VirtualComboBoxItem, ComboBoxItems, ComboBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IComboBoxItemEvents_CP.h"
#include "helpers.h"
#include "ComboBox.h"


class ATL_NO_VTABLE ComboBoxItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ComboBoxItem, &CLSID_ComboBoxItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ComboBoxItem>,
    public Proxy_IComboBoxItemEvents<ComboBoxItem>, 
    #ifdef UNICODE
    	public IDispatchImpl<IComboBoxItem, &IID_IComboBoxItem, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IComboBoxItem, &IID_IComboBoxItem, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ComboBox;
	friend class ComboBoxItems;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_COMBOBOXITEM)

		BEGIN_COM_MAP(ComboBoxItem)
			COM_INTERFACE_ENTRY(IComboBoxItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ComboBoxItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IComboBoxItemEvents))
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
	/// \name Implementation of IComboBoxItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Height property</em>
	///
	/// Retrieves the item's height in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c OwnerDrawItems property is not set to \c odiOwnerDrawVariableHeight, this property
	///          is read-only. Use the \c IComboBox::ItemHeight property instead.
	///
	/// \sa put_Height, GetRectangle, ComboBox::get_OwnerDrawItems, ComboBox::get_ItemHeight
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
	///          is read-only. Use the \c IComboBox::ItemHeight property instead.
	///
	/// \sa get_Height, GetRectangle, ComboBox::put_OwnerDrawItems, ComboBox::put_ItemHeight
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
	/// \sa put_ItemData, ComboBox::Raise_FreeItemData
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
	/// \sa get_ItemData, ComboBox::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE put_ItemData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the item is the currently selected item. If this property is set to
	/// \c VARIANT_TRUE, the item is selected; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ComboBox::get_SelectedItem
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the item's text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);

	/// \brief <em>Retrieves the bounding rectangle of either the item or a part of it</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the drop-down list box control's client area)
	/// of either the item or a part of it.
	///
	/// \param[in] rectangleType The rectangle to retrieve. Any of the values defined by the
	///            \c ItemRectangleTypeConstants enumeration is valid.
	/// \param[out] PXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
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
	HRESULT SaveState(VirtualComboBoxItem* pTarget);
	/// \brief <em>Sets the owner of this item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerCBox
	void SetOwner(__in_opt ComboBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ComboBox object that owns this item</em>
		///
		/// \sa SetOwner
		ComboBox* pOwnerCBox;
		/// \brief <em>The index of the item wrapped by this object</em>
		int itemIndex;
		/// \brief <em>Specifies the value that will be returned by the \c ItemData property, if the specified item index is -1</em>
		///
		/// \sa get_ItemData
		ULONG_PTR itemData;

		Properties()
		{
			pOwnerCBox = NULL;
			itemIndex = -2;
			itemData = NULL;
		}

		~Properties();

		/// \brief <em>Retrieves the owning combo box' window handle</em>
		///
		/// \return The window handle of the combo box that contains this item.
		///
		/// \sa pOwnerCBox, GetLBoxHWnd
		HWND GetCBoxHWnd(void);
		/// \brief <em>Retrieves the owning combo box' drop-down list box window handle</em>
		///
		/// \return The window handle of the drop-down list box of the combo box that contains this item.
		///
		/// \sa pOwnerCBox, GetCBoxHWnd
		HWND GetLBoxHWnd(void);
	} properties;
};     // ComboBoxItem

OBJECT_ENTRY_AUTO(__uuidof(ComboBoxItem), ComboBoxItem)