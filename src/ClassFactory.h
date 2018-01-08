//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa ComboBox, DriveComboBox, ImageComboBox, ListBox
//////////////////////////////////////////////////////////////////////


#pragma once

#include "ComboBoxItem.h"
#include "ComboBoxItems.h"
#include "DriveComboBoxItem.h"
#include "DriveComboBoxItems.h"
#include "ImageComboBoxItem.h"
#include "ImageComboBoxItems.h"
#include "ListBoxItem.h"
#include "ListBoxItems.h"
#include "TargetOLEDataObject.h"


class ClassFactory
{
public:
	/// \brief <em>Creates a \c ComboBoxItem object</em>
	///
	/// Creates a \c ComboBoxItem object that represents a given combo box item and returns
	/// its \c IComboBoxItem implementation as a smart pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pCBox The \c ComboBox object the \c ComboBoxItem object will work on.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	/// \param[in] itemData Specifies the value that the created object will return in its \c ItemData
	///            property, if the specified item index is -1.
	///
	/// \return A smart pointer to the created object's \c IComboBoxItem implementation.
	///
	/// \sa ComboBoxItem::get_ItemData
	static inline CComPtr<IComboBoxItem> InitComboItem(int itemIndex, ComboBox* pCBox, BOOL validateItemIndex = TRUE, ULONG_PTR itemData = NULL)
	{
		CComPtr<IComboBoxItem> pItem = NULL;
		InitComboItem(itemIndex, pCBox, IID_IComboBoxItem, reinterpret_cast<IUnknown**>(&pItem), validateItemIndex, itemData);
		return pItem;
	};

	/// \brief <em>Creates a \c ComboBoxItem object</em>
	///
	/// \overload
	///
	/// Creates a \c ComboBoxItem object that represents a given combo box item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pCBox The \c ComboBox object the \c ComboBoxItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	/// \param[in] itemData Specifies the value that the created object will return in its \c ItemData
	///            property, if the specified item index is -1.
	///
	/// \sa ComboBoxItem::get_ItemData
	static inline void InitComboItem(int itemIndex, ComboBox* pCBox, REFIID requiredInterface, IUnknown** ppItem, BOOL validateItemIndex = TRUE, ULONG_PTR itemData = NULL)
	{
		ATLASSERT_POINTER(ppItem, IUnknown*);
		ATLASSUME(ppItem);

		*ppItem = NULL;
		if(validateItemIndex && !IsValidComboBoxItemIndex(itemIndex, pCBox)) {
			// there's nothing useful we could return
			return;
		}

		// create a ComboBoxItem object and initialize it with the given parameters
		CComObject<ComboBoxItem>* pCBoxItemObj = NULL;
		CComObject<ComboBoxItem>::CreateInstance(&pCBoxItemObj);
		pCBoxItemObj->AddRef();
		pCBoxItemObj->SetOwner(pCBox);
		pCBoxItemObj->Attach(itemIndex, itemData);
		pCBoxItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItem));
		pCBoxItemObj->Release();
	};

	/// \brief <em>Creates a \c DriveComboBoxItem object</em>
	///
	/// Creates a \c DriveComboBoxItem object that represents a given image combo box item and returns
	/// its \c IDriveComboBoxItem implementation as a smart pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pDCBox The \c DriveComboBox object the \c DriveComboBoxItem object will work on.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IDriveComboBoxItem implementation.
	static inline CComPtr<IDriveComboBoxItem> InitDriveComboItem(int itemIndex, DriveComboBox* pDCBox, BOOL validateItemIndex = TRUE)
	{
		CComPtr<IDriveComboBoxItem> pItem = NULL;
		InitDriveComboItem(itemIndex, pDCBox, IID_IDriveComboBoxItem, reinterpret_cast<IUnknown**>(&pItem), validateItemIndex);
		return pItem;
	};

	/// \brief <em>Creates a \c DriveComboBoxItem object</em>
	///
	/// \overload
	///
	/// Creates a \c DriveComboBoxItem object that represents a given image combo box item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pDCBox The \c DriveComboBox object the \c DriveComboBoxItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	static inline void InitDriveComboItem(int itemIndex, DriveComboBox* pDCBox, REFIID requiredInterface, IUnknown** ppItem, BOOL validateItemIndex = TRUE)
	{
		ATLASSERT_POINTER(ppItem, IUnknown*);
		ATLASSUME(ppItem);

		*ppItem = NULL;
		if(validateItemIndex && !IsValidComboBoxItemIndex(itemIndex, pDCBox)) {
			// there's nothing useful we could return
			return;
		}

		// create a DriveComboBoxItem object and initialize it with the given parameters
		CComObject<DriveComboBoxItem>* pDCBoxItemObj = NULL;
		CComObject<DriveComboBoxItem>::CreateInstance(&pDCBoxItemObj);
		pDCBoxItemObj->AddRef();
		pDCBoxItemObj->SetOwner(pDCBox);
		pDCBoxItemObj->Attach(itemIndex);
		pDCBoxItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItem));
		pDCBoxItemObj->Release();
	};

	/// \brief <em>Creates a \c ImageComboBoxItem object</em>
	///
	/// Creates a \c ImageComboBoxItem object that represents a given image combo box item and returns
	/// its \c IImageComboBoxItem implementation as a smart pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pICBox The \c ImageComboBox object the \c ImageComboBoxItem object will work on.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IImageComboBoxItem implementation.
	static inline CComPtr<IImageComboBoxItem> InitImageComboItem(int itemIndex, ImageComboBox* pICBox, BOOL validateItemIndex = TRUE)
	{
		CComPtr<IImageComboBoxItem> pItem = NULL;
		InitImageComboItem(itemIndex, pICBox, IID_IImageComboBoxItem, reinterpret_cast<IUnknown**>(&pItem), validateItemIndex);
		return pItem;
	};

	/// \brief <em>Creates a \c ImageComboBoxItem object</em>
	///
	/// \overload
	///
	/// Creates a \c ImageComboBoxItem object that represents a given image combo box item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pICBox The \c ImageComboBox object the \c ImageComboBoxItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	static inline void InitImageComboItem(int itemIndex, ImageComboBox* pICBox, REFIID requiredInterface, IUnknown** ppItem, BOOL validateItemIndex = TRUE)
	{
		ATLASSERT_POINTER(ppItem, IUnknown*);
		ATLASSUME(ppItem);

		*ppItem = NULL;
		if(validateItemIndex && !IsValidComboBoxItemIndex(itemIndex, pICBox)) {
			// there's nothing useful we could return
			return;
		}

		// create a ImageComboBoxItem object and initialize it with the given parameters
		CComObject<ImageComboBoxItem>* pICBoxItemObj = NULL;
		CComObject<ImageComboBoxItem>::CreateInstance(&pICBoxItemObj);
		pICBoxItemObj->AddRef();
		pICBoxItemObj->SetOwner(pICBox);
		pICBoxItemObj->Attach(itemIndex);
		pICBoxItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItem));
		pICBoxItemObj->Release();
	};

	/// \brief <em>Creates a \c ListBoxItem object</em>
	///
	/// Creates a \c ListBoxItem object that represents a given list box item and returns
	/// its \c IListBoxItem implementation as a smart pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pLBox The \c ListBox object the \c ListBoxItem object will work on.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	/// \param[in] itemData Specifies the value that the created object will return in its \c ItemData
	///            property, if the specified item index is -1.
	///
	/// \return A smart pointer to the created object's \c IListBoxItem implementation.
	///
	/// \sa ListBoxItem::get_ItemData
	static inline CComPtr<IListBoxItem> InitListItem(int itemIndex, ListBox* pLBox, BOOL validateItemIndex = TRUE, ULONG_PTR itemData = NULL)
	{
		CComPtr<IListBoxItem> pItem = NULL;
		InitListItem(itemIndex, pLBox, IID_IListBoxItem, reinterpret_cast<IUnknown**>(&pItem), validateItemIndex, itemData);
		return pItem;
	};

	/// \brief <em>Creates a \c ListBoxItem object</em>
	///
	/// \overload
	///
	/// Creates a \c ListBoxItem object that represents a given list box item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pLBox The \c ListBox object the \c ListBoxItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	/// \param[in] itemData Specifies the value that the created object will return in its \c ItemData
	///            property, if the specified item index is -1.
	///
	/// \sa ListBoxItem::get_ItemData
	static inline void InitListItem(int itemIndex, ListBox* pLBox, REFIID requiredInterface, IUnknown** ppItem, BOOL validateItemIndex = TRUE, ULONG_PTR itemData = NULL)
	{
		ATLASSERT_POINTER(ppItem, IUnknown*);
		ATLASSUME(ppItem);

		*ppItem = NULL;
		if(validateItemIndex && !IsValidListBoxItemIndex(itemIndex, pLBox)) {
			// there's nothing useful we could return
			return;
		}

		// create a ListBoxItem object and initialize it with the given parameters
		CComObject<ListBoxItem>* pLBoxItemObj = NULL;
		CComObject<ListBoxItem>::CreateInstance(&pLBoxItemObj);
		pLBoxItemObj->AddRef();
		pLBoxItemObj->SetOwner(pLBox);
		pLBoxItemObj->Attach(itemIndex, itemData);
		pLBoxItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItem));
		pLBoxItemObj->Release();
	};

	/// \brief <em>Creates a \c ComboBoxItems object</em>
	///
	/// Creates a \c ComboBoxItems object that represents a collection of combo box items and returns
	/// its \c IComboBoxItems implementation as a smart pointer.
	///
	/// \param[in] pCBox The \c ComboBox object the \c ComboBoxItems object will work on.
	///
	/// \return A smart pointer to the created object's \c IComboBoxItems implementation.
	static inline CComPtr<IComboBoxItems> InitComboItems(ComboBox* pCBox)
	{
		CComPtr<IComboBoxItems> pItems = NULL;
		InitComboItems(pCBox, IID_IComboBoxItems, reinterpret_cast<IUnknown**>(&pItems));
		return pItems;
	};

	/// \brief <em>Creates a \c ComboBoxItems object</em>
	///
	/// \overload
	///
	/// Creates a \c ComboBoxItems object that represents a collection of combo box items and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pCBox The \c ComboBox object the \c ComboBoxItems object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItems Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitComboItems(ComboBox* pCBox, REFIID requiredInterface, IUnknown** ppItems)
	{
		ATLASSERT_POINTER(ppItems, IUnknown*);
		ATLASSUME(ppItems);

		*ppItems = NULL;
		// create a ComboBoxItems object and initialize it with the given parameters
		CComObject<ComboBoxItems>* pCBoxItemsObj = NULL;
		CComObject<ComboBoxItems>::CreateInstance(&pCBoxItemsObj);
		pCBoxItemsObj->AddRef();

		pCBoxItemsObj->SetOwner(pCBox);

		pCBoxItemsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItems));
		pCBoxItemsObj->Release();
	};

	/// \brief <em>Creates a \c DriveComboBoxItems object</em>
	///
	/// Creates a \c DriveComboBoxItems object that represents a collection of drive combo box items and
	/// returns its \c IDriveComboBoxItems implementation as a smart pointer.
	///
	/// \param[in] pDCBox The \c DriveComboBox object the \c DriveComboBoxItems object will work on.
	///
	/// \return A smart pointer to the created object's \c IDriveComboBoxItems implementation.
	static inline CComPtr<IDriveComboBoxItems> InitDriveComboItems(DriveComboBox* pDCBox)
	{
		CComPtr<IDriveComboBoxItems> pItems = NULL;
		InitDriveComboItems(pDCBox, IID_IDriveComboBoxItems, reinterpret_cast<IUnknown**>(&pItems));
		return pItems;
	};

	/// \brief <em>Creates a \c DriveComboBoxItems object</em>
	///
	/// \overload
	///
	/// Creates a \c DriveComboBoxItems object that represents a collection of drive combo box items and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDCBox The \c DriveComboBox object the \c DriveComboBoxItems object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItems Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitDriveComboItems(DriveComboBox* pDCBox, REFIID requiredInterface, IUnknown** ppItems)
	{
		ATLASSERT_POINTER(ppItems, IUnknown*);
		ATLASSUME(ppItems);

		*ppItems = NULL;
		// create a DriveComboBoxItems object and initialize it with the given parameters
		CComObject<DriveComboBoxItems>* pDCBoxItemsObj = NULL;
		CComObject<DriveComboBoxItems>::CreateInstance(&pDCBoxItemsObj);
		pDCBoxItemsObj->AddRef();

		pDCBoxItemsObj->SetOwner(pDCBox);

		pDCBoxItemsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItems));
		pDCBoxItemsObj->Release();
	};

	/// \brief <em>Creates a \c ImageComboBoxItems object</em>
	///
	/// Creates a \c ImageComboBoxItems object that represents a collection of image combo box items and
	/// returns its \c IImageComboBoxItems implementation as a smart pointer.
	///
	/// \param[in] pICBox The \c ImageComboBox object the \c ImageComboBoxItems object will work on.
	///
	/// \return A smart pointer to the created object's \c IImageComboBoxItems implementation.
	static inline CComPtr<IImageComboBoxItems> InitImageComboItems(ImageComboBox* pICBox)
	{
		CComPtr<IImageComboBoxItems> pItems = NULL;
		InitImageComboItems(pICBox, IID_IImageComboBoxItems, reinterpret_cast<IUnknown**>(&pItems));
		return pItems;
	};

	/// \brief <em>Creates a \c ImageComboBoxItems object</em>
	///
	/// \overload
	///
	/// Creates a \c ImageComboBoxItems object that represents a collection of image combo box items and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pICBox The \c ImageComboBox object the \c ImageComboBoxItems object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItems Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitImageComboItems(ImageComboBox* pICBox, REFIID requiredInterface, IUnknown** ppItems)
	{
		ATLASSERT_POINTER(ppItems, IUnknown*);
		ATLASSUME(ppItems);

		*ppItems = NULL;
		// create a ImageComboBoxItems object and initialize it with the given parameters
		CComObject<ImageComboBoxItems>* pICBoxItemsObj = NULL;
		CComObject<ImageComboBoxItems>::CreateInstance(&pICBoxItemsObj);
		pICBoxItemsObj->AddRef();

		pICBoxItemsObj->SetOwner(pICBox);

		pICBoxItemsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItems));
		pICBoxItemsObj->Release();
	};

	/// \brief <em>Creates a \c ListBoxItems object</em>
	///
	/// Creates a \c ListBoxItems object that represents a collection of list box items and returns
	/// its \c IListBoxItems implementation as a smart pointer.
	///
	/// \param[in] pLBox The \c ListBox object the \c ListBoxItems object will work on.
	///
	/// \return A smart pointer to the created object's \c IListBoxItems implementation.
	static inline CComPtr<IListBoxItems> InitListItems(ListBox* pLBox)
	{
		CComPtr<IListBoxItems> pItems = NULL;
		InitListItems(pLBox, IID_IListBoxItems, reinterpret_cast<IUnknown**>(&pItems));
		return pItems;
	};

	/// \brief <em>Creates a \c ListBoxItems object</em>
	///
	/// \overload
	///
	/// Creates a \c ListBoxItems object that represents a collection of list box items and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pLBox The \c ListBox object the \c ListBoxItems object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItems Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitListItems(ListBox* pLBox, REFIID requiredInterface, IUnknown** ppItems)
	{
		ATLASSERT_POINTER(ppItems, IUnknown*);
		ATLASSUME(ppItems);

		*ppItems = NULL;
		// create a ListBoxItems object and initialize it with the given parameters
		CComObject<ListBoxItems>* pLBoxItemsObj = NULL;
		CComObject<ListBoxItems>::CreateInstance(&pLBoxItemsObj);
		pLBoxItemsObj->AddRef();

		pLBoxItemsObj->SetOwner(pLBox);

		pLBoxItemsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItems));
		pLBoxItemsObj->Release();
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, LPUNKNOWN* ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, LPUNKNOWN);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};
};     // ClassFactory