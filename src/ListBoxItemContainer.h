//////////////////////////////////////////////////////////////////////
/// \class ListBoxItemContainer
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ListBoxItem objects</em>
///
/// This class provides easy access to collections of \c ListBoxItem objects. While a
/// \c ListBoxItems object is used to group items that have certain properties in common, a
/// \c ListBoxItemContainer object is used to group any items and acts more like a clipboard.
///
/// \if UNICODE
///   \sa CBLCtlsLibU::IListBoxItemContainer, ListBoxItem, ListBoxItems, ListBox
/// \else
///   \sa CBLCtlsLibA::IListBoxItemContainer, ListBoxItem, ListBoxItems, ListBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "CBLCtlsU.h"
#else
	#include "CBLCtlsA.h"
#endif
#include "_IListBoxItemContainerEvents_CP.h"
#include "IItemContainer.h"
#include "helpers.h"
#include "ListBox.h"
#include "ListBoxItem.h"


class ATL_NO_VTABLE ListBoxItemContainer : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListBoxItemContainer, &CLSID_ListBoxItemContainer>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListBoxItemContainer>,
    public Proxy_IListBoxItemContainerEvents<ListBoxItemContainer>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IListBoxItemContainer, &IID_IListBoxItemContainer, &LIBID_CBLCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IListBoxItemContainer, &IID_IListBoxItemContainer, &LIBID_CBLCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IItemContainer
{
	friend class ListBox;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used to generate and set this object's ID.
	///
	/// \sa ~ListBoxItemContainer
	ListBoxItemContainer();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used to deregister the container.
	///
	/// \sa ListBoxItemContainer
	~ListBoxItemContainer();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTBOXITEMCONTAINER)

		BEGIN_COM_MAP(ListBoxItemContainer)
			COM_INTERFACE_ENTRY(IListBoxItemContainer)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListBoxItemContainer)
			CONNECTION_POINT_ENTRY(__uuidof(_IListBoxItemContainerEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportsErrorInfo
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
	/// \name Implementation of IListBoxItemContainer
	///
	//@{
	/// \brief <em>Retrieves a \c ListBoxItem object from the collection</em>
	///
	/// Retrieves a \c ListBoxItem object from the collection that wraps the list box item identified
	/// by \c itemID.
	///
	/// \param[in] itemIdentifier A value that identifies the list box item to retrieve.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. This must be
	///            \c iitID in non-virtual mode and \c iitIndex in virtual mode.
	/// \param[out] ppItem Receives the item's \c IListBoxItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ListBoxItem::get_ID, ListBoxItem::get_Index, ListBox::get_VirtualMode, Add, Remove,
	///       CBLCtlsLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa ListBoxItem::get_ID, ListBoxItem::get_Index, ListBox::get_VirtualMode, Add, Remove,
	///       CBLCtlsLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType = iitID, IListBoxItem** ppItem = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ListBoxItem objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds the specified item(s) to the collection</em>
	///
	/// \param[in] items The item(s) to add. May be either an item ID, a \c ListBoxItem object or a
	///            \c ListBoxItems collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ListBoxItem::get_ID, Count, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT items);
	/// \brief <em>Clones the collection object</em>
	///
	/// Retrieves an exact copy of the collection.
	///
	/// \param[out] ppClone Receives the clone's \c IListBoxItemContainer implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ListBox::CreateItemContainer
	virtual HRESULT STDMETHODCALLTYPE Clone(IListBoxItemContainer** ppClone);
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
	/// \brief <em>Retrieves an imagelist containing the items' common drag image</em>
	///
	/// Retrieves the handle to an imagelist containing a bitmap that can be used to visualize
	/// dragging of the items of this collection.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] phImageList The handle to the imagelist.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa ListBox::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	/// \brief <em>Removes the specified item from the collection</em>
	///
	/// \param[in] itemIdentifier A value that identifies the list box item to remove.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. This must be
	///            \c iitID in non-virtual mode and \c iitIndex in virtual mode.
	/// \param[in] removePhysically If \c VARIANT_TRUE, the item is removed from the control, too.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ListBoxItem::get_ID, ListBoxItem::get_Index, ListBox::get_VirtualMode, Add, Count, RemoveAll,
	///       CBLCtlsLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa ListBoxItem::get_ID, ListBoxItem::get_Index, ListBox::get_VirtualMode, Add, Count, RemoveAll,
	///       CBLCtlsLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType = iitID, VARIANT_BOOL removePhysically = VARIANT_FALSE);
	/// \brief <em>Removes all items from the collection</em>
	///
	/// \param[in] removePhysically If \c VARIANT_TRUE, the items get removed from the control, too.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(VARIANT_BOOL removePhysically = VARIANT_FALSE);
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
	/// \name Implementation of IItemContainer
	///
	//@{
	/// \brief <em>A item was removed and the item container should check its content</em>
	///
	/// \param[in] itemIdentifier <strong>Non-virtual mode:</strong> The unique ID of the removed item.
	///            <strong>Virtual mode:</strong> The zero-based index of the removed item. If -1, all items
	///            were removed.
	///
	/// \sa ListBox::RemoveItemFromItemContainers
	void RemovedItem(LONG itemIdentifier);
	/// \brief <em>An item's unique ID has been changed and the item container should check its content</em>
	///
	/// \param[in] oldItemID The old unique ID of the removed item.
	/// \param[in] newItemID The new unique ID of the removed item.
	///
	/// \remarks This method is not supported in virtual mode.
	///
	/// \sa ListBox::UpdateItemIDInItemContainers
	void ReplacedItemID(LONG oldItemID, LONG newItemID);
	/// \brief <em>Retrieves the container's ID used to identify it</em>
	///
	/// \return The container's ID.
	///
	/// \sa ListBox::DeregisterItemContainer, containerID
	DWORD GetID(void);
	/// \brief <em>Holds the container's ID</em>
	DWORD containerID;
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ListBox object that owns this collection</em>
		///
		/// \sa SetOwner
		ListBox* pOwnerLBox;
		/// \brief <em>If set to \c TRUE, the container works with item indexes instead of item IDs</em>
		UINT useIndexes : 1;
		#ifdef USE_STL
			/// \brief <em>Holds the items that this object contains</em>
			std::vector<LONG> items;
		#else
			/// \brief <em>Holds the items that this object contains</em>
			//CAtlArray<LONG> items;
		#endif
		/// \brief <em>Points to the next enumerated item</em>
		int nextEnumeratedItem;

		Properties()
		{
			pOwnerLBox = NULL;
			useIndexes = FALSE;
			nextEnumeratedItem = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning list box' window handle</em>
		///
		/// \return The window handle of the list box that contains the items in this collection.
		///
		/// \sa pOwnerLBox
		HWND GetLBoxHWnd(void);
	} properties;

	/* TODO: If we move this one into the Properties struct, the compiler complains that the private
	         = operator of CAtlArray cannot be accessed?! */
	#ifndef USE_STL
		/// \brief <em>Holds the items that this object contains</em>
		CAtlArray<LONG> items;
	#endif

	/// \brief <em>Incremented by the constructor and used as the constructed object's ID</em>
	///
	/// \sa ListBoxItemContainer, containerID
	static DWORD nextID;
};     // ListBoxItemContainer

OBJECT_ENTRY_AUTO(__uuidof(ListBoxItemContainer), ListBoxItemContainer)