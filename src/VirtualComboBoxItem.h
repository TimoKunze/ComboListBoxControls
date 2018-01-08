//////////////////////////////////////////////////////////////////////
/// \class VirtualComboBoxItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing combo box item</em>
///
/// This class is a wrapper around a combo box item that does not yet or not anymore exist in the control.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IVirtualComboBoxItem, ComboBoxItem, ComboBox
/// \else
///   \sa CBLCtlsLibA::IVirtualComboBoxItem, ComboBoxItem, ComboBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IVirtualComboBoxItemEvents_CP.h"
#include "helpers.h"
#include "ComboBox.h"


class ATL_NO_VTABLE VirtualComboBoxItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualComboBoxItem, &CLSID_VirtualComboBoxItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualComboBoxItem>,
    public Proxy_IVirtualComboBoxItemEvents< VirtualComboBoxItem>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualComboBoxItem, &IID_IVirtualComboBoxItem, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualComboBoxItem, &IID_IVirtualComboBoxItem, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ComboBox;
	friend class ComboBoxItem;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALCOMBOBOXITEM)

		BEGIN_COM_MAP(VirtualComboBoxItem)
			COM_INTERFACE_ENTRY(IVirtualComboBoxItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualComboBoxItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualComboBoxItemEvents))
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
	/// \name Implementation of IVirtualComboBoxItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index that will identify or has identified the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ComboBoxItems::Add
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
	/// \sa ComboBox::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE get_ItemData(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Initializes this object with given data</em>
	///
	/// Initializes this object with the settings from a given source.
	///
	/// \param[in] itemIndex The item's zero-based index.
	/// \param[in] data The item's data or the address of its text buffer.
	/// \param[in] itemData The data associated with the item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(int itemIndex, LPARAM data, LONG itemData);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[out] itemIndex The item's zero-based index.
	/// \param[out] data The item's data or the address of its text buffer.
	/// \param[out] itemData The data associated with the item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(int& itemIndex, LPARAM& data, LONG& itemData);
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
		/// \brief <em>The item's zero-based index</em>
		int itemIndex;
		/// \brief <em>The item's data or the address of its text buffer</em>
		LPARAM data;
		/// \brief <em>If \c TRUE, \c data holds the address of a string; otherwise it's just an integer value</em>
		UINT dataIsString : 1;
		/// \brief <em>The data associated with the item</em>
		LONG itemData;

		Properties()
		{
			pOwnerCBox = NULL;
			itemIndex = -1;
			data = 0;
			dataIsString = FALSE;
			itemData = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning combo box' window handle</em>
		///
		/// \return The window handle of the combo box that contains this item.
		///
		/// \sa pOwnerCBox
		HWND GetCBoxHWnd(void);
	} properties;
};     // VirtualComboBoxItem

OBJECT_ENTRY_AUTO(__uuidof(VirtualComboBoxItem), VirtualComboBoxItem)