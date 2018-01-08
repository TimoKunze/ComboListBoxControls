//////////////////////////////////////////////////////////////////////
/// \class DriveComboBoxItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing drive combo item</em>
///
/// This class is a wrapper around a combo box item that - unlike an item wrapped by
/// \c VirtualDriveComboBoxItem - really exists in the control.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IDriveComboBoxItem, VirtualDriveComboBoxItem, DriveComboBoxItems, DriveComboBox
/// \else
///   \sa CBLCtlsLibA::IDriveComboBoxItem, VirtualDriveComboBoxItem, DriveComboBoxItems, DriveComboBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IDriveComboBoxItemEvents_CP.h"
#include "helpers.h"
#include "DriveComboBox.h"


class ATL_NO_VTABLE DriveComboBoxItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<DriveComboBoxItem, &CLSID_DriveComboBoxItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<DriveComboBoxItem>,
    public Proxy_IDriveComboBoxItemEvents<DriveComboBoxItem>, 
    #ifdef UNICODE
    	public IDispatchImpl<IDriveComboBoxItem, &IID_IDriveComboBoxItem, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IDriveComboBoxItem, &IID_IDriveComboBoxItem, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class DriveComboBox;
	friend class DriveComboBoxItems;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_DRIVECOMBOBOXITEM)

		BEGIN_COM_MAP(DriveComboBoxItem)
			COM_INTERFACE_ENTRY(IDriveComboBoxItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(DriveComboBoxItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IDriveComboBoxItemEvents))
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
	/// \name Implementation of IDriveComboBoxItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c DriveType property</em>
	///
	/// Retrieves the drive's type. Any of the values defined by the \c DriveTypeConstants enumeration
	/// is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Path, CBLCtlsLibU::DriveTypeConstants
	/// \else
	///   \sa get_Path, CBLCtlsLibA::DriveTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DriveType(DriveTypeConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c Height property</em>
	///
	/// Retrieves the item's height in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa GetRectangle, DriveComboBox::get_ItemHeight
	virtual HRESULT STDMETHODCALLTYPE get_Height(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's icon in the control's \c ilItems image list.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IconIndex, DriveComboBox::get_hImageList, get_SelectedIconIndex, get_OverlayIndex,
	///       CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_IconIndex, DriveComboBox::get_hImageList, get_SelectedIconIndex, get_OverlayIndex,
	///       CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the item's icon in the control's \c ilItems image list. If set to -1, the
	/// control will fire the \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IconIndex, DriveComboBox::put_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo,
	///       put_SelectedIconIndex, put_OverlayIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_IconIndex, DriveComboBox::put_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo,
	///       put_SelectedIconIndex, put_OverlayIndex, BLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
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
	/// \brief <em>Retrieves the current setting of the \c Indent property</em>
	///
	/// Retrieves the item's indentation in steps of 10 pixels. If set to 1, the item's indentation will be
	/// 10 pixels; if set to 2, it will be 20 pixels and so on. If set to -1, the control will fire the
	/// \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Indent, DriveComboBox::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Indent(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c Indent property</em>
	///
	/// Sets the item's indentation in steps of 10 pixels. If set to 1, the item's indentation will be
	/// 10 pixels; if set to 2, it will be 20 pixels and so on. If set to -1, the control will fire the
	/// \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Indent, DriveComboBox::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE put_Indent(OLE_XSIZE_PIXELS newValue);
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
	/// \sa put_ItemData, DriveComboBox::Raise_FreeItemData
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
	/// \sa get_ItemData, DriveComboBox::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE put_ItemData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c OverlayIndex property</em>
	///
	/// Retrieves the one-based index of the item's overlay icon in the control's \c ilItems image list. If
	/// set to -1, the control will fire the \c ItemGetDisplayInfo event each time this property's value is
	/// required. An index of 0 means that no overlay is drawn for this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OverlayIndex, DriveComboBox::get_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo,
	///       get_IconIndex, get_SelectedIconIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_OverlayIndex, DriveComboBox::get_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo,
	///       get_IconIndex, get_SelectedIconIndex, CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OverlayIndex(LONG* pValue);
	/// \brief <em>Sets the \c OverlayIndex property</em>
	///
	/// Sets the one-based index of the item's overlay icon in the control's \c ilItems image list. If
	/// set to -1, the control will fire the \c ItemGetDisplayInfo event each time this property's value is
	/// required. An index of 0 means that no overlay is drawn for this item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OverlayIndex, DriveComboBox::put_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo,
	///       put_IconIndex, put_SelectedIconIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_OverlayIndex, DriveComboBox::put_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo,
	///       put_IconIndex, put_SelectedIconIndex, CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OverlayIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Path property</em>
	///
	/// Retrieves the drive's path.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_DriveType, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_Path(BSTR* pValue);
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
	/// \sa DriveComboBox::get_SelectedItem
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c SelectedIconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's selected icon in the control's \c ilItems image list.
	/// The selected icon is used instead of the normal icon identified by the \c IconIndex property if the
	/// item is the caret item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SelectedIconIndex, DriveComboBox::get_hImageList, get_IconIndex,
	///       get_OverlayIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_SelectedIconIndex, DriveComboBox::get_hImageList, get_IconIndex,
	///       get_OverlayIndex, CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SelectedIconIndex(LONG* pValue);
	/// \brief <em>Sets the \c SelectedIconIndex property</em>
	///
	/// Sets the zero-based index of the item's selected icon in the control's \c ilItems image list.
	/// The selected icon is used instead of the normal icon identified by the \c IconIndex property if the
	/// item is the caret item. If set to -1, the control will fire the \c ItemGetDisplayInfo event each time
	/// this property's value is needed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SelectedIconIndex, DriveComboBox::put_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo,
	///       put_IconIndex, put_OverlayIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_SelectedIconIndex, DriveComboBox::put_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo,
	///       put_IconIndex, put_OverlayIndex, CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SelectedIconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text. If set to \c NULL, the control will fire the \c ItemGetDisplayInfo
	/// event each time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, get_Path, DriveComboBox::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the item's text. If set to \c NULL, the control will fire the \c ItemGetDisplayInfo
	/// event each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, get_Path, DriveComboBox::Raise_ItemGetDisplayInfo
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
	/// \sa DriveComboBox::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
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
	void Attach(int itemIndex);
	/// \brief <em>Detaches this object from an item</em>
	///
	/// Detaches this object from the item it currently wraps, so that it doesn't wrap any item anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndDCBox The drive combo box window the method will work on.
	///
	/// \return An \c HRESULT error code.
	HRESULT SaveState(PCOMBOBOXEXITEM pTarget, HWND hWndDCBox = NULL);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \overload
	HRESULT SaveState(VirtualDriveComboBoxItem* pTarget);
	/// \brief <em>Sets the owner of this item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerDCBox
	void SetOwner(__in_opt DriveComboBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c DriveComboBox object that owns this item</em>
		///
		/// \sa SetOwner
		DriveComboBox* pOwnerDCBox;
		/// \brief <em>The index of the item wrapped by this object</em>
		int itemIndex;

		Properties()
		{
			pOwnerDCBox = NULL;
			itemIndex = -2;
		}

		~Properties();

		/// \brief <em>Retrieves the owning drive combo box' window handle</em>
		///
		/// \return The window handle of the drive combo box that contains this item.
		///
		/// \sa pOwnerDCBox, GetLBoxHWnd
		HWND GetDCBoxHWnd(void);
		/// \brief <em>Retrieves the owning drive combo box' drop-down list box window handle</em>
		///
		/// \return The window handle of the drop-down list box of the drive combo box that contains this item.
		///
		/// \sa pOwnerDCBox, GetDCBoxHWnd
		HWND GetLBoxHWnd(void);
		/// \brief <em>Invalidates the combo box control contained in the owning drive combo box</em>
		///
		/// \sa InvalidateListBox, pOwnerDCBox
		void InvalidateComboBox(void);
		/// \brief <em>Invalidates the owning drive combo box' drop-down list box control</em>
		///
		/// \sa InvalidateComboBox, pOwnerDCBox
		void InvalidateListBox(void);
	} properties;

	/// \brief <em>Writes a given object's settings to a given target</em>
	///
	/// \param[in] itemIndex The item for which to save the settings.
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndDCBox The drive combo box window the method will work on.
	///
	/// \return An \c HRESULT error code.
	HRESULT SaveState(int itemIndex, PCOMBOBOXEXITEM pTarget, HWND hWndDCBox = NULL);
};     // DriveComboBoxItem

OBJECT_ENTRY_AUTO(__uuidof(DriveComboBoxItem), DriveComboBoxItem)