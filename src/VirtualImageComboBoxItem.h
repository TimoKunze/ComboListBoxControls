//////////////////////////////////////////////////////////////////////
/// \class VirtualImageComboBoxItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing image combo box item</em>
///
/// This class is a wrapper around an image combo box item that does not yet or not anymore exist in the
/// control.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IVirtualImageComboBoxItem, ImageComboBoxItem, ImageComboBox
/// \else
///   \sa CBLCtlsLibA::IVirtualImageComboBoxItem, ImageComboBoxItem, ImageComboBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IVirtualImageComboBoxItemEvents_CP.h"
#include "helpers.h"
#include "ImageComboBox.h"


class ATL_NO_VTABLE VirtualImageComboBoxItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualImageComboBoxItem, &CLSID_VirtualImageComboBoxItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualImageComboBoxItem>,
    public Proxy_IVirtualImageComboBoxItemEvents< VirtualImageComboBoxItem>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualImageComboBoxItem, &IID_IVirtualImageComboBoxItem, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualImageComboBoxItem, &IID_IVirtualImageComboBoxItem, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ImageComboBox;
	friend class ImageComboBoxItem;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALIMAGECOMBOBOXITEM)

		BEGIN_COM_MAP(VirtualImageComboBoxItem)
			COM_INTERFACE_ENTRY(IVirtualImageComboBoxItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualImageComboBoxItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualImageComboBoxItemEvents))
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
	/// \name Implementation of IVirtualImageComboBoxItem
	///
	//@{
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
	///   \sa ImageComboBox::get_hImageList, ImageComboBox::Raise_ItemGetDisplayInfo, get_SelectedIconIndex,
	///       get_OverlayIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa ImageComboBox::get_hImageList, ImageComboBox::Raise_ItemGetDisplayInfo, get_SelectedIconIndex,
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
	/// \sa ImageComboBox::Raise_ItemGetDisplayInfo
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
	/// \sa ImageComboBox::Raise_FreeItemData
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
	///   \sa ImageComboBox::get_hImageList, ImageComboBox::Raise_ItemGetDisplayInfo, get_IconIndex,
	///       get_SelectedIconIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa ImageComboBox::get_hImageList, ImageComboBox::Raise_ItemGetDisplayInfo, get_IconIndex,
	///       get_SelectedIconIndex, CBLCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OverlayIndex(LONG* pValue);
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
	///   \sa ImageComboBox::get_hImageList, get_IconIndex, get_OverlayIndex, CBLCtlsLibU::ImageListConstants
	/// \else
	///   \sa ImageComboBox::get_hImageList, get_IconIndex, get_OverlayIndex, CBLCtlsLibA::ImageListConstants
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
	/// \sa ImageComboBox::Raise_ItemGetDisplayInfo
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
	/// \sa Properties::pOwnerICBox
	void SetOwner(__in_opt ImageComboBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ImageComboBox object that owns this item</em>
		///
		/// \sa SetOwner
		ImageComboBox* pOwnerICBox;
		/// \brief <em>A structure holding this item's settings</em>
		COMBOBOXEXITEM settings;

		Properties()
		{
			pOwnerICBox = NULL;
			ZeroMemory(&settings, sizeof(settings));
			settings.iItem = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning image combo box' window handle</em>
		///
		/// \return The window handle of the image combo box that contains this item.
		///
		/// \sa pOwnerICBox
		HWND GetICBoxHWnd(void);
	} properties;
};     // VirtualImageComboBoxItem

OBJECT_ENTRY_AUTO(__uuidof(VirtualImageComboBoxItem), VirtualImageComboBoxItem)