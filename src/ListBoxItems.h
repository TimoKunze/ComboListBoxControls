//////////////////////////////////////////////////////////////////////
/// \class ListBoxItems
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ListBoxItem objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ListBoxItem
/// objects. A \c ListBoxItems object is used to group items that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IListBoxItems, ListBoxItem, ListBox
/// \else
///   \sa CBLCtlsLibA::IListBoxItems, ListBoxItem, ListBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IListBoxItemsEvents_CP.h"
#include "helpers.h"
#include "CWindowEx.h"
#include "ListBox.h"
#include "ListBoxItem.h"


class ATL_NO_VTABLE ListBoxItems : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListBoxItems, &CLSID_ListBoxItems>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListBoxItems>,
    public Proxy_IListBoxItemsEvents<ListBoxItems>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IListBoxItems, &IID_IListBoxItems, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListBoxItems, &IID_IListBoxItems, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ListBox;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTBOXITEMS)

		BEGIN_COM_MAP(ListBoxItems)
			COM_INTERFACE_ENTRY(IListBoxItems)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListBoxItems)
			CONNECTION_POINT_ENTRY(__uuidof(_IListBoxItemsEvents))
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
	/// the \c ListBoxItem objects managed by this collection object.
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
	///                the pointer to a item's \c IListBoxItem implementation.
	/// \param[out] pNumberOfItemsReturned The number of items that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ListBoxItem,
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
	/// \name Implementation of IListBoxItems
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
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Some of the
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
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Some of the
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
	/// An \c IListBoxItems collection can be filtered by any of \c IListBoxItem's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Some of the
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
	/// An \c IListBoxItems collection can be filtered by any of \c IListBoxItem's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Some of the
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
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Some of the
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
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Some of the
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
	/// \brief <em>Retrieves a \c ListBoxItem object from the collection</em>
	///
	/// Retrieves a \c ListBoxItem object from the collection that wraps the list box item identified
	/// by \c itemIdentifier.
	///
	/// \param[in] itemIdentifier A value that identifies the list box item to be retrieved.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppItem Receives the item's \c IListBoxItem implementation.
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
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType = iitIndex, IListBoxItem** ppItem = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ListBoxItem objects
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

	/// \brief <em>Adds an item to the list box</em>
	///
	/// Adds an item with the specified properties at the specified position in the control and returns a
	/// \c ListBoxItem object wrapping the inserted item.
	///
	/// \param[in] itemText The new item's caption text.
	/// \param[in] insertAt The new item's zero-based index. If set to -1, the item will be inserted
	///            as the last item.
	/// \param[in] itemData A \c LONG value that will be associated with the item.
	/// \param[out] ppAddedItem Receives the added item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c HasStrings property is set to \c VARIANT_FALSE and the \c OwnerDrawItems property
	///          is set to a value other than \c odiDontOwnerDraw, the parameter \c itemText is ignored.
	///
	/// \sa Count, Remove, RemoveAll, ListBoxItem::get_Text, ListBoxItem::get_ItemData,
	///     ListBox::get_HasStrings, ListBox::get_OwnerDrawItems, ListBox::PrepareForItemInsertions
	virtual HRESULT STDMETHODCALLTYPE Add(BSTR itemText, LONG insertAt = -1, LONG itemData = 0, IListBoxItem** ppAddedItem = NULL);
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
	/// Retrieves the number of \c ListBoxItem objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified item in the collection from the list box</em>
	///
	/// \param[in] itemIdentifier A value that identifies the list box item to be removed.
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
	/// \brief <em>Removes all items in the collection from the list box</em>
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
	/// \sa Properties::pOwnerLBox
	void SetOwner(__in_opt ListBox* pOwner);

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
	/// \param[in] numberOfItems The number of list box items in the control.
	/// \param[in] hWndLBox The list box window the method will work on.
	///
	/// \return The item being the collection's first candidate. -1 if no item was found.
	///
	/// \sa GetNextItemToProcess, Count, RemoveAll, Next
	int GetFirstItemToProcess(int numberOfItems, HWND /*hWndLBox*/);
	/// \brief <em>Retrieves the control's next item that might be in the collection</em>
	///
	/// \param[in] previousItem The item at which to start the search for a new collection candidate.
	/// \param[in] numberOfItems The number of items in the control.
	/// \param[in] hWndLBox The list box window the method will work on.
	///
	/// \return The next item being a candidate for the collection. -1 if no item was found.
	///
	/// \sa GetFirstItemToProcess, Count, RemoveAll, Next
	int GetNextItemToProcess(int previousItem, int numberOfItems, HWND /*hWndLBox*/);
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
	/// \param[in] hWndLBox The list box window the method will work on.
	///
	/// \return \c TRUE, if the item is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(int itemIndex, HWND hWndLBox = NULL);
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
		/// \param[in] hWndLBox The list box window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveItems(std::list<int>& itemsToRemove, HWND hWndLBox);
	#else
		/// \brief <em>Removes the specified items</em>
		///
		/// \param[in] itemsToRemove A vector containing all items to remove.
		/// \param[in] hWndLBox The list box window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveItems(CAtlList<int>& itemsToRemove, HWND hWndLBox);
	#endif

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#define NUMBEROFFILTERS_LST 12
		/// \brief <em>Holds the \c CaseSensitiveFilters property's setting</em>
		///
		/// \sa get_CaseSensitiveFilters, put_CaseSensitiveFilters
		UINT caseSensitiveFilters : 1;
		/// \brief <em>Holds the \c ComparisonFunction property's setting</em>
		///
		/// \sa get_ComparisonFunction, put_ComparisonFunction
		LPVOID comparisonFunction[NUMBEROFFILTERS_LST];
		/// \brief <em>Holds the \c Filter property's setting</em>
		///
		/// \sa get_Filter, put_Filter
		VARIANT filter[NUMBEROFFILTERS_LST];
		/// \brief <em>Holds the \c FilterType property's setting</em>
		///
		/// \sa get_FilterType, put_FilterType
		FilterTypeConstants filterType[NUMBEROFFILTERS_LST];

		/// \brief <em>The \c ListBox object that owns this collection</em>
		///
		/// \sa SetOwner
		ListBox* pOwnerLBox;
		/// \brief <em>Holds the last enumerated item</em>
		int lastEnumeratedItem;
		/// \brief <em>If \c TRUE, we must filter the items</em>
		///
		/// \sa put_Filter, put_FilterType
		UINT usingFilters : 1;

		Properties()
		{
			caseSensitiveFilters = FALSE;
			pOwnerLBox = NULL;
			lastEnumeratedItem = -1;

			for(int i = 0; i < NUMBEROFFILTERS_LST; ++i) {
				VariantInit(&filter[i]);
				filterType[i] = ftDeactivated;
				comparisonFunction[i] = NULL;
			}
			usingFilters = FALSE;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);

		/// \brief <em>Retrieves the owning list box' window handle</em>
		///
		/// \return The window handle of the list box that contains the items in this collection.
		///
		/// \sa pOwnerLBox
		HWND GetLBoxHWnd(void);
	} properties;
};     // ListBoxItems

OBJECT_ENTRY_AUTO(__uuidof(ListBoxItems), ListBoxItems)