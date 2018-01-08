//////////////////////////////////////////////////////////////////////
/// \class VirtualDriveComboBoxItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing drive combo box item</em>
///
/// This class is a wrapper around an drive combo box item that does not yet or not anymore exist in the
/// control.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IVirtualDriveComboBoxItem, DriveComboBoxItem, DriveComboBox
/// \else
///   \sa CBLCtlsLibA::IVirtualDriveComboBoxItem, DriveComboBoxItem, DriveComboBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IVirtualDriveComboBoxItemEvents_CP.h"
#include "helpers.h"
#include "DriveComboBox.h"


class ATL_NO_VTABLE VirtualDriveComboBoxItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualDriveComboBoxItem, &CLSID_VirtualDriveComboBoxItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualDriveComboBoxItem>,
    public Proxy_IVirtualDriveComboBoxItemEvents< VirtualDriveComboBoxItem>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualDriveComboBoxItem, &IID_IVirtualDriveComboBoxItem, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualDriveComboBoxItem, &IID_IVirtualDriveComboBoxItem, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class DriveComboBox;
	friend class DriveComboBoxItem;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALDRIVECOMBOBOXITEM)

		BEGIN_COM_MAP(VirtualDriveComboBoxItem)
			COM_INTERFACE_ENTRY(IVirtualDriveComboBoxItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualDriveComboBoxItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualDriveComboBoxItemEvents))
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
	/// \name Implementation of IVirtualDriveComboBoxItem
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
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's icon in the control's \c ilItems image list. If set to
	/// -1, the control will fire the \c ItemGetDisplayInfo event each time this property's value is
	/// required. If set to -2, no icon will be or was displayed for this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa DriveComboBox::get_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo, get_SelectedIconIndex,
	///       get_OverlayIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa DriveComboBox::get_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo, get_SelectedIconIndex,
	///       get_OverlayIndex, CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
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
	/// \remarks This property is read-only.
	///
	/// \sa DriveComboBox::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Indent(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index that will identify or has identified the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemData property</em>
	///
	/// Retrieves the \c LONG value that will be or was associated with the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa DriveComboBox::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE get_ItemData(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c OverlayIndex property</em>
	///
	/// Retrieves the one-based index of the item's overlay icon in the control's \c ilItems image list. If
	/// set to -1, the control will fire the \c ItemGetDisplayInfo event each time this property's value is
	/// required. An index of 0 means that no overlay will be or was drawn for this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa DriveComboBox::get_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo, get_IconIndex,
	///       get_SelectedIconIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa DriveComboBox::get_hImageList, DriveComboBox::Raise_ItemGetDisplayInfo, get_IconIndex,
	///       get_SelectedIconIndex, CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OverlayIndex(LONG* pValue);
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
	/// \brief <em>Retrieves the current setting of the \c SelectedIconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's selected icon in the control's \c ilItems image list.
	/// The selected icon is used instead of the normal icon identified by the \c IconIndex property if the
	/// item is the selected item. If set to -1, the control will fire the \c ItemGetDisplayInfo event each
	/// time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa DriveComboBox::get_hImageList, get_IconIndex, get_OverlayIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa DriveComboBox::get_hImageList, get_IconIndex, get_OverlayIndex, CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SelectedIconIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text. If set to \c NULL, the control will fire the \c ItemGetDisplayInfo
	/// event each time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Path, DriveComboBox::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Initializes this object with given data</em>
	///
	/// Initializes this object with the settings from a given source.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(__in PCOMBOBOXEXITEM pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(__in PCOMBOBOXEXITEM pTarget);
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
		/// \brief <em>The zero-based index of the drive with 0 being drive A:</em>
		int driveIndex;
		/// \brief <em>A structure holding this item's settings</em>
		COMBOBOXEXITEM settings;

		Properties()
		{
			pOwnerDCBox = NULL;
			driveIndex = -1;
			ZeroMemory(&settings, sizeof(settings));
			settings.iItem = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning drive combo box' window handle</em>
		///
		/// \return The window handle of the drive combo box that contains this item.
		///
		/// \sa pOwnerDCBox
		HWND GetDCBoxHWnd(void);
	} properties;
};     // VirtualDriveComboBoxItem

OBJECT_ENTRY_AUTO(__uuidof(VirtualDriveComboBoxItem), VirtualDriveComboBoxItem)