//////////////////////////////////////////////////////////////////////
/// \class ImageComboBoxItems
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ImageComboBoxItem objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ImageComboBoxItem
/// objects. A \c ImageComboBoxItems object is used to group items that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IImageComboBoxItems, ImageComboBoxItem, ImageComboBox
/// \else
///   \sa CBLCtlsLibA::IImageComboBoxItems, ImageComboBoxItem, ImageComboBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IImageComboBoxItemsEvents_CP.h"
#include "helpers.h"
#include "CWindowEx.h"
#include "ImageComboBox.h"
#include "ImageComboBoxItem.h"


class ATL_NO_VTABLE ImageComboBoxItems : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ImageComboBoxItems, &CLSID_ImageComboBoxItems>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ImageComboBoxItems>,
    public Proxy_IImageComboBoxItemsEvents<ImageComboBoxItems>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IImageComboBoxItems, &IID_IImageComboBoxItems, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IImageComboBoxItems, &IID_IImageComboBoxItems, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ImageComboBox;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_IMAGECOMBOBOXITEMS)

		BEGIN_COM_MAP(ImageComboBoxItems)
			COM_INTERFACE_ENTRY(IImageComboBoxItems)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ImageComboBoxItems)
			CONNECTION_POINT_ENTRY(__uuidof(_IImageComboBoxItemsEvents))
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the items</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ComboBoxItem objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x items</em>
	///
	/// Retrieves the next \c numberOfMaxItems items from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of items the array identified by \c pItems can
	///            contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a item's \c IImageComboBoxItem implementation.
	/// \param[out] pNumberOfItemsReturned The number of items that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ImageComboBoxItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// item in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x items</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip items.
	///
	/// \param[in] numberOfItemsToSkip The number of items to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IImageComboBoxItems
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c CaseSensitiveFilters property</em>
	///
	/// Retrieves whether string comparisons, that are done when applying the filters on an item, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CaseSensitiveFilters, get_Filter, get_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE get_CaseSensitiveFilters(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CaseSensitiveFilters property</em>
	///
	/// Sets whether string comparisons, that are done when applying the filters on an item, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaseSensitiveFilters, put_Filter, put_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE put_CaseSensitiveFilters(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ComparisonFunction property</em>
	///
	/// Retrieves an item filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(T itemProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG or \c BSTR). This function
	/// must compare its arguments and return a non-zero value if the arguments are equal and zero
	/// otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       CBLCtlsLibU::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       CBLCtlsLibA::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue);
	/// \brief <em>Sets the \c ComparisonFunction property</em>
	///
	/// Sets an item filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(T itemProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG or \c BSTR). This function
	/// must compare its arguments and return a non-zero value if the arguments are equal and zero
	/// otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       CBLCtlsLibU::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       CBLCtlsLibA::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Filter property</em>
	///
	/// Retrieves an item filter.\n
	/// An \c IComboBoxItems collection can be filtered by any of \c IImageComboBoxItem's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction, CBLCtlsLibU::FilteredPropertyConstants
	/// \else
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction, CBLCtlsLibA::FilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue);
	/// \brief <em>Sets the \c Filter property</em>
	///
	/// Sets an item filter.\n
	/// An \c IComboBoxItems collection can be filtered by any of \c IImageComboBoxItem's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       CBLCtlsLibU::FilteredPropertyConstants
	/// \else
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       CBLCtlsLibA::FilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c FilterType property</em>
	///
	/// Retrieves an item filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_FilterType, get_Filter, CBLCtlsLibU::FilteredPropertyConstants,
	///       CBLCtlsLibU::FilterTypeConstants
	/// \else
	///   \sa put_FilterType, get_Filter, CBLCtlsLibA::FilteredPropertyConstants,
	///       CBLCtlsLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue);
	/// \brief <em>Sets the \c FilterType property</em>
	///
	/// Sets an item filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_FilterType, put_Filter, IsPartOfCollection, CBLCtlsLibU::FilteredPropertyConstants,
	///       CBLCtlsLibU::FilterTypeConstants
	/// \else
	///   \sa get_FilterType, put_Filter, IsPartOfCollection, CBLCtlsLibA::FilteredPropertyConstants,
	///       CBLCtlsLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue);
	/// \brief <em>Retrieves a \c ImageComboBoxItem object from the collection</em>
	///
	/// Retrieves a \c ImageComboBoxItem object from the collection that wraps the combo box item identified
	/// by \c itemIdentifier.
	///
	/// \param[in] itemIdentifier A value that identifies the combo box item to be retrieved.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppItem Receives the item's \c IImageComboBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa Add, Remove, Contains, CBLCtlsLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, Contains, CBLCtlsLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType = iitIndex, IImageComboBoxItem** ppItem = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ImageComboBoxItem objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator A pointer to the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds an item to the combo box</em>
	///
	/// Adds an item with the specified properties at the specified position in the control and returns a
	/// \c ImageComboBoxItem object wrapping the inserted item.
	///
	/// \param[in] itemText The new item's caption text. If set to \c NULL, the control will fire the
	///            \c ItemGetDisplayInfo event each time this property's value is required.
	/// \param[in] insertAt The new item's zero-based index. If set to -1, the item will be inserted
	///            as the last item.
	/// \param[in] iconIndex The zero-based index of the item's icon in the control's \c ilItems image list.
	///            If set to -1, the control will fire the \c ItemGetDisplayInfo event each time this
	///            property's value is required. A value of -2 means 'not specified' and is valid if
	///            there's no image list associated with the control.
	/// \param[in] selectedIconIndex The zero-based index of the item's selected icon in the control's
	///            \c ilItems image list. This icon will be used instead of the normal icon identified by
	///            \c iconIndex if the item is the selected item. If set to -1, the control will fire the
	///            \c ItemGetDisplayInfo event each time this property's value is required. If set to -2,
	///            the normal icon specified by \c iconIndex will be used.
	/// \param[in] overlayIndex The one-based index of the item's overlay icon in the control's \c ilItems
	///            image list. If set to -1, the control will fire the \c ItemGetDisplayInfo event each time
	///            this property's value is required. If set to 0, no overlay is drawn for this item.
	/// \param[in] itemIndentation The new item's indentation in steps of 10 pixels. If set to 1, the
	///            item's indentation will be 10 pixels; if set to 2, it will be 20 pixels and so on.
	///            If set to -1, the control will fire the \c ItemGetDisplayInfo event each time this
	///            property's value is required.
	/// \param[in] itemData A \c LONG value that will be associated with the item.
	/// \param[out] ppAddedItem Receives the added item's \c IImageComboBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Count, Remove, RemoveAll, ImageComboBoxItem::get_Text, ImageComboBoxItem::get_IconIndex,
	///     ImageComboBoxItem::get_SelectedIconIndex, ImageComboBoxItem::get_OverlayIndex,
	///     ImageComboBoxItem::get_ItemData, ImageComboBox::get_hImageList,
	///     ImageComboBox::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE Add(BSTR itemText, LONG insertAt = -1, LONG iconIndex = -2, LONG selectedIconIndex = -2, LONG overlayIndex = 0, LONG itemIndentation = 0, LONG itemData = 0, IImageComboBoxItem** ppAddedItem = NULL);
	/// \brief <em>Retrieves whether the specified item is part of the item collection</em>
	///
	/// \param[in] itemIdentifier A value that identifies the item to be checked.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the item is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, Add, Remove, CBLCtlsLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_Filter, Add, Remove, CBLCtlsLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType = iitIndex, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the items in the collection</em>
	///
	/// Retrieves the number of \c ImageComboBoxItem objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified item in the collection from the combo box</em>
	///
	/// \param[in] itemIdentifier A value that identifies the combo box item to be removed.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ItemIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, CBLCtlsLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, CBLCtlsLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType = iitIndex);
	/// \brief <em>Removes all items in the collection from the combo box</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerICBox
	void SetOwner(__in_opt ImageComboBox* pOwner);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Filter validation
	///
	//@{
	/// \brief <em>Validates a filter for a \c VARIANT_BOOL type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c VARIANT_BOOL.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidIntegerFilter, IsValidStringFilter, put_Filter
	BOOL IsValidBooleanFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c LONG or compatible.
	///
	/// \param[in] filter The \c VARIANT to check.
	/// \param[in] minValue The minimum value that the corresponding property would accept.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidStringFilter, put_Filter
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c BSTR type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c BSTR.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidIntegerFilter, put_Filter
	BOOL IsValidStringFilter(VARIANT& filter);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Filter appliance
	///
	//@{
	/// \brief <em>Retrieves the control's first item that might be in the collection</em>
	///
	/// \param[in] numberOfItems The number of combo box items in the control.
	/// \param[in] hWndICBox The image combo box window the method will work on.
	///
	/// \return The item being the collection's first candidate. -1 if no item was found.
	///
	/// \sa GetNextItemToProcess, Count, RemoveAll, Next
	int GetFirstItemToProcess(int numberOfItems, HWND /*hWndICBox*/);
	/// \brief <em>Retrieves the control's next item that might be in the collection</em>
	///
	/// \param[in] previousItem The item at which to start the search for a new collection candidate.
	/// \param[in] numberOfItems The number of items in the control.
	/// \param[in] hWndICBox The image combo box window the method will work on.
	///
	/// \return The next item being a candidate for the collection. -1 if no item was found.
	///
	/// \sa GetFirstItemToProcess, Count, RemoveAll, Next
	int GetNextItemToProcess(int previousItem, int numberOfItems, HWND /*hWndICBox*/);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified boolean value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsIntegerInSafeArray, IsStringInSafeArray, get_ComparisonFunction
	BOOL IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified integer value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsStringInSafeArray, get_ComparisonFunction
	BOOL IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified \c BSTR value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsIntegerInSafeArray, get_ComparisonFunction
	BOOL IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether an item is part of the collection (applying the filters)</em>
	///
	/// \param[in] itemIndex The item to check.
	/// \param[in] hWndICBox The image combo box window the method will work on.
	///
	/// \return \c TRUE, if the item is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(int itemIndex, HWND hWndICBox = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Shortens a filter as much as possible</em>
	///
	/// Optimizes a filter by detecting redundancies, tautologies and so on.
	///
	/// \param[in] filteredProperty The filter to optimize. Any of the values defined by the
	///            \c FilteredPropertyConstants enumeration is valid.
	///
	/// \sa put_Filter, put_FilterType
	void OptimizeFilter(FilteredPropertyConstants filteredProperty);
	#ifdef USE_STL
		/// \brief <em>Removes the specified items</em>
		///
		/// \param[in] itemsToRemove A vector containing all items to remove.
		/// \param[in] hWndICBox The image combo box window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveItems(std::list<int>& itemsToRemove, HWND hWndICBox);
	#else
		/// \brief <em>Removes the specified items</em>
		///
		/// \param[in] itemsToRemove A vector containing all items to remove.
		/// \param[in] hWndICBox The image combo box window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveItems(CAtlList<int>& itemsToRemove, HWND hWndICBox);
	#endif

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#define NUMBEROFFILTERS_IMGCMB 15
		/// \brief <em>Holds the \c CaseSensitiveFilters property's setting</em>
		///
		/// \sa get_CaseSensitiveFilters, put_CaseSensitiveFilters
		UINT caseSensitiveFilters : 1;
		/// \brief <em>Holds the \c ComparisonFunction property's setting</em>
		///
		/// \sa get_ComparisonFunction, put_ComparisonFunction
		LPVOID comparisonFunction[NUMBEROFFILTERS_IMGCMB];
		/// \brief <em>Holds the \c Filter property's setting</em>
		///
		/// \sa get_Filter, put_Filter
		VARIANT filter[NUMBEROFFILTERS_IMGCMB];
		/// \brief <em>Holds the \c FilterType property's setting</em>
		///
		/// \sa get_FilterType, put_FilterType
		FilterTypeConstants filterType[NUMBEROFFILTERS_IMGCMB];

		/// \brief <em>The \c ImageComboBox object that owns this collection</em>
		///
		/// \sa SetOwner
		ImageComboBox* pOwnerICBox;
		/// \brief <em>Holds the last enumerated item</em>
		int lastEnumeratedItem;
		/// \brief <em>If \c TRUE, we must filter the items</em>
		///
		/// \sa put_Filter, put_FilterType
		UINT usingFilters : 1;

		Properties()
		{
			caseSensitiveFilters = FALSE;
			pOwnerICBox = NULL;
			lastEnumeratedItem = -1;

			for(int i = 0; i < NUMBEROFFILTERS_IMGCMB; ++i) {
				VariantInit(&filter[i]);
				filterType[i] = ftDeactivated;
				comparisonFunction[i] = NULL;
			}
			usingFilters = FALSE;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);

		/// \brief <em>Retrieves the owning image combo box' window handle</em>
		///
		/// \return The window handle of the image combo box that contains the items in this collection.
		///
		/// \sa pOwnerICBox
		HWND GetICBoxHWnd(void);
	} properties;
};     // ImageComboBoxItems

OBJECT_ENTRY_AUTO(__uuidof(ImageComboBoxItems), ImageComboBoxItems)