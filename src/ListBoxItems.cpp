// ListBoxItems.cpp: Manages a collection of ListBoxItem objects

#include "stdafx.h"
#include "ListBoxItems.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListBoxItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListBoxItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ListBoxItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ListBoxItems>* pLBoxItemsObj = NULL;
	CComObject<ListBoxItems>::CreateInstance(&pLBoxItemsObj);
	pLBoxItemsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pLBoxItemsObj->properties);

	pLBoxItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pLBoxItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ListBoxItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	// check each item in the list box
	int numberOfItems = static_cast<int>(SendMessage(hWndLBox, LB_GETCOUNT, 0, 0));
	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		do {
			if(properties.lastEnumeratedItem >= 0) {
				if(properties.lastEnumeratedItem < numberOfItems) {
					properties.lastEnumeratedItem = GetNextItemToProcess(properties.lastEnumeratedItem, numberOfItems, hWndLBox);
				}
			} else {
				properties.lastEnumeratedItem = GetFirstItemToProcess(numberOfItems, hWndLBox);
			}
			if(properties.lastEnumeratedItem >= numberOfItems) {
				properties.lastEnumeratedItem = -1;
			}
		} while((properties.lastEnumeratedItem != -1) && (!IsPartOfCollection(properties.lastEnumeratedItem, hWndLBox)));

		if(properties.lastEnumeratedItem != -1) {
			ClassFactory::InitListItem(properties.lastEnumeratedItem, properties.pOwnerLBox, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
			pItems[i].vt = VT_DISPATCH;
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ListBoxItems::Reset(void)
{
	properties.lastEnumeratedItem = -1;
	return S_OK;
}

STDMETHODIMP ListBoxItems::Skip(ULONG numberOfItemsToSkip)
{
	VARIANT dummy;
	ULONG numItemsReturned = 0;
	// we could skip all items at once, but it's easier to skip them one after the other
	for(ULONG i = 1; i <= numberOfItemsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numItemsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numItemsReturned == 0) {
			// there're no more items to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


BOOL ListBoxItems::IsValidBooleanFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BOOL);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListBoxItems::IsValidIntegerFilter(VARIANT& filter, int minValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListBoxItems::IsValidIntegerFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_UI4));
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListBoxItems::IsValidStringFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BSTR);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

int ListBoxItems::GetFirstItemToProcess(int numberOfItems, HWND /*hWndLBox*/)
{
	//ATLASSERT(IsWindow(hWndLBox));

	if(numberOfItems == 0) {
		return -1;
	}
	return 0;
}

int ListBoxItems::GetNextItemToProcess(int previousItem, int numberOfItems, HWND /*hWndLBox*/)
{
	//ATLASSERT(IsWindow(hWndLBox));

	if(previousItem < numberOfItems - 1) {
		return previousItem + 1;
	}
	return -1;
}

BOOL ListBoxItems::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(VARIANT_BOOL, VARIANT_BOOL);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.boolVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(element.boolVal == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ListBoxItems::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		int v = 0;
		if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT))) {
			v = element.intVal;
		}
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LONG, LONG);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, v)) {
				foundMatch = TRUE;
			}
		} else {
			if(v == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ListBoxItems::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(BSTR, BSTR);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.bstrVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(properties.caseSensitiveFilters) {
				if(lstrcmpW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			} else {
				if(lstrcmpiW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ListBoxItems::IsPartOfCollection(int itemIndex, HWND hWndLBox/* = NULL*/)
{
	if(!hWndLBox) {
		hWndLBox = properties.GetLBoxHWnd();
	}
	ATLASSERT(IsWindow(hWndLBox));

	if(!IsValidListBoxItemIndex(itemIndex, hWndLBox)) {
		return FALSE;
	}

	BOOL isPart = FALSE;

	if(properties.filterType[fpIndex] != ftDeactivated) {
		if(IsIntegerInSafeArray(properties.filter[fpIndex].parray, itemIndex, properties.comparisonFunction[fpIndex])) {
			if(properties.filterType[fpIndex] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpIndex] == ftIncluding) {
				goto Exit;
			}
		}
	}
	if(properties.filterType[fpSelected] != ftDeactivated) {
		if(IsBooleanInSafeArray(properties.filter[fpSelected].parray, BOOL2VARIANTBOOL(SendMessage(hWndLBox, LB_GETSEL, itemIndex, 0)), properties.comparisonFunction[fpSelected])) {
			if(properties.filterType[fpSelected] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpSelected] == ftIncluding) {
				goto Exit;
			}
		}
	}
	if(properties.filterType[fpItemData] != ftDeactivated) {
		if(IsIntegerInSafeArray(properties.filter[fpItemData].parray, static_cast<int>(SendMessage(hWndLBox, LB_GETITEMDATA, itemIndex, 0)), properties.comparisonFunction[fpItemData])) {
			if(properties.filterType[fpItemData] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpItemData] == ftIncluding) {
				goto Exit;
			}
		}
	}
	if(properties.filterType[fpText] != ftDeactivated) {
		CComBSTR text = L"";

		BOOL expectsString;
		DWORD style = CWindow(hWndLBox).GetStyle();
		if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
			expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
		} else {
			expectsString = TRUE;
		}
		if(expectsString) {
			int textLength = static_cast<int>(SendMessage(hWndLBox, LB_GETTEXTLEN, itemIndex, 0));
			if(textLength != LB_ERR) {
				LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
				if(pBuffer) {
					SendMessage(hWndLBox, LB_GETTEXT, itemIndex, reinterpret_cast<LPARAM>(pBuffer));
					text = pBuffer;
					HeapFree(GetProcessHeap(), 0, pBuffer);
				}
			}
		}

		if(IsStringInSafeArray(properties.filter[fpText].parray, text, properties.comparisonFunction[fpText])) {
			if(properties.filterType[fpText] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpText] == ftIncluding) {
				goto Exit;
			}
		}
	}
	isPart = TRUE;

Exit:
	return isPart;
}

void ListBoxItems::OptimizeFilter(FilteredPropertyConstants filteredProperty)
{
	if(filteredProperty != fpSelected) {
		// currently we optimize boolean filters only
		return;
	}

	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(properties.filter[filteredProperty].parray, 1, &lBound);
	SafeArrayGetUBound(properties.filter[filteredProperty].parray, 1, &uBound);

	// now that we have the bounds, iterate the array
	VARIANT element;
	UINT numberOfTrues = 0;
	UINT numberOfFalses = 0;
	for(LONG i = lBound; i <= uBound; ++i) {
		SafeArrayGetElement(properties.filter[filteredProperty].parray, &i, &element);
		if(element.boolVal == VARIANT_FALSE) {
			++numberOfFalses;
		} else {
			++numberOfTrues;
		}

		VariantClear(&element);
	}

	if(numberOfTrues > 0 && numberOfFalses > 0) {
		// we've something like True Or False Or True - we can deactivate this filter
		properties.filterType[filteredProperty] = ftDeactivated;
		VariantClear(&properties.filter[filteredProperty]);
	} else if(numberOfTrues == 0 && numberOfFalses > 1) {
		// False Or False Or False... is still False, so we need just one False
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_FALSE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	} else if(numberOfFalses == 0 && numberOfTrues > 1) {
		// True Or True Or True... is still True, so we need just one True
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_TRUE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	}
}

#ifdef USE_STL
	HRESULT ListBoxItems::RemoveItems(std::list<int>& itemsToRemove, HWND hWndLBox)
#else
	HRESULT ListBoxItems::RemoveItems(CAtlList<int>& itemsToRemove, HWND hWndLBox)
#endif
{
	ATLASSERT(IsWindow(hWndLBox));

	CWindowEx2(hWndLBox).InternalSetRedraw(FALSE);
	// sort in reverse order
	#ifdef USE_STL
		itemsToRemove.sort(std::greater<int>());
		for(std::list<int>::const_iterator iter = itemsToRemove.begin(); iter != itemsToRemove.end(); ++iter) {
			SendMessage(hWndLBox, LB_DELETESTRING, *iter, 0);
		}
	#else
		// perform a crude bubble sort
		for(size_t j = 0; j < itemsToRemove.GetCount(); ++j) {
			for(size_t i = 0; i < itemsToRemove.GetCount() - 1; ++i) {
				if(itemsToRemove.GetAt(itemsToRemove.FindIndex(i)) < itemsToRemove.GetAt(itemsToRemove.FindIndex(i + 1))) {
					itemsToRemove.SwapElements(itemsToRemove.FindIndex(i), itemsToRemove.FindIndex(i + 1));
				}
			}
		}

		for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
			SendMessage(hWndLBox, LB_DELETESTRING, itemsToRemove.GetAt(itemsToRemove.FindIndex(i)), 0);
		}
	#endif
	CWindowEx2(hWndLBox).InternalSetRedraw(TRUE);

	return S_OK;
}


ListBoxItems::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS_LST; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerLBox) {
		pOwnerLBox->Release();
	}
}

void ListBoxItems::Properties::CopyTo(ListBoxItems::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerLBox = this->pOwnerLBox;
		if(pTarget->pOwnerLBox) {
			pTarget->pOwnerLBox->AddRef();
		}
		pTarget->lastEnumeratedItem = this->lastEnumeratedItem;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS_LST; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND ListBoxItems::Properties::GetLBoxHWnd(void)
{
	ATLASSUME(pOwnerLBox);

	OLE_HANDLE handle = NULL;
	pOwnerLBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ListBoxItems::SetOwner(ListBox* pOwner)
{
	if(properties.pOwnerLBox) {
		properties.pOwnerLBox->Release();
	}
	properties.pOwnerLBox = pOwner;
	if(properties.pOwnerLBox) {
		properties.pOwnerLBox->AddRef();
	}
}


STDMETHODIMP ListBoxItems::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP ListBoxItems::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP ListBoxItems::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIndex:
		case fpItemData:
		case fpSelected:
		case fpText:
			*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListBoxItems::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
{
	switch(filteredProperty) {
		case fpIndex:
		case fpItemData:
		case fpSelected:
		case fpText:
			properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListBoxItems::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIndex:
		case fpItemData:
		case fpSelected:
		case fpText:
			VariantClear(pValue);
			VariantCopy(pValue, &properties.filter[filteredProperty]);
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListBoxItems::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	// check 'newValue'
	switch(filteredProperty) {
		case fpSelected:
			if(!IsValidBooleanFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpIndex:
			if(!IsValidIntegerFilter(newValue, 0)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpItemData:
			if(!IsValidIntegerFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpText:
			if(!IsValidStringFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		default:
			return E_INVALIDARG;
			break;
	}

	VariantClear(&properties.filter[filteredProperty]);
	VariantCopy(&properties.filter[filteredProperty], &newValue);
	OptimizeFilter(filteredProperty);
	return S_OK;
}

STDMETHODIMP ListBoxItems::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIndex:
		case fpItemData:
		case fpSelected:
		case fpText:
			*pValue = properties.filterType[filteredProperty];
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListBoxItems::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if(newValue < 0 || newValue > 2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	switch(filteredProperty) {
		case fpIndex:
		case fpItemData:
		case fpSelected:
		case fpText:
			properties.filterType[filteredProperty] = newValue;
			if(newValue != ftDeactivated) {
				properties.usingFilters = TRUE;
			} else {
				properties.usingFilters = FALSE;
				for(int i = 0; i < NUMBEROFFILTERS_LST; ++i) {
					if(properties.filterType[i] != ftDeactivated) {
						properties.usingFilters = TRUE;
						break;
					}
				}
			}
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListBoxItems::get_Item(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/, IListBoxItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, IListBoxItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	// retrieve the item's index
	int itemIndex = -1;
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerLBox) {
				itemIndex = properties.pOwnerLBox->IDToItemIndex(itemIdentifier);
			}
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	if(itemIndex > -1) {
		if(IsPartOfCollection(itemIndex)) {
			ClassFactory::InitListItem(itemIndex, properties.pOwnerLBox, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
		}
		return S_OK;
	} else {
		// item not found
		if(itemIdentifierType == iitIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}
}

STDMETHODIMP ListBoxItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ListBoxItems::Add(BSTR itemText, LONG insertAt/* = -1*/, LONG itemData/* = 0*/, IListBoxItem** ppAddedItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedItem, IListBoxItem*);
	if(!ppAddedItem) {
		return E_POINTER;
	}

	if(insertAt < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));
	BOOL expectsString;
	DWORD style = CWindow(hWndLBox).GetStyle();
	if(style & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE)) {
		expectsString = ((style & LBS_HASSTRINGS) == LBS_HASSTRINGS);
	} else {
		expectsString = TRUE;
	}

	COLE2T converter(itemText);
	LPTSTR pBuffer = converter;

	int itemIndex = -1;
	if(insertAt == -1) {
		// we could use LB_INSERTSTRING as well, but this wouldn't cause a sorted list to sort
		if(expectsString) {
			properties.pOwnerLBox->BufferItemData(itemData);
			itemIndex = static_cast<int>(SendMessage(hWndLBox, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(pBuffer)));
		} else {
			itemIndex = static_cast<int>(SendMessage(hWndLBox, LB_ADDSTRING, 0, itemData));
		}
	} else {
		if(expectsString) {
			properties.pOwnerLBox->BufferItemData(itemData);
			itemIndex = static_cast<int>(SendMessage(hWndLBox, LB_INSERTSTRING, insertAt, reinterpret_cast<LPARAM>(pBuffer)));
		} else {
			itemIndex = static_cast<int>(SendMessage(hWndLBox, LB_INSERTSTRING, insertAt, itemData));
		}
	}
	if(itemIndex > LB_ERR) {
		ClassFactory::InitListItem(itemIndex, properties.pOwnerLBox, IID_IListBoxItem, reinterpret_cast<LPUNKNOWN*>(ppAddedItem));
		return S_OK;
	}
	*ppAddedItem = NULL;
	return E_FAIL;
}

STDMETHODIMP ListBoxItems::Contains(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// retrieve the item's index
	int itemIndex = -1;
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerLBox) {
				itemIndex = properties.pOwnerLBox->IDToItemIndex(itemIdentifier);
			}
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(itemIndex));
	return S_OK;
}

STDMETHODIMP ListBoxItems::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	if(!properties.usingFilters) {
		*pValue = static_cast<LONG>(SendMessage(hWndLBox, LB_GETCOUNT, 0, 0));
	}

	if(properties.filterType[fpSelected] != ftDeactivated && !properties.comparisonFunction[fpSelected]) {
		int numberOfActiveFilters = 0;
		for(int i = 0; i < NUMBEROFFILTERS_LST; ++i) {
			if(properties.filterType[i] != ftDeactivated) {
				++numberOfActiveFilters;
			}
		}
		if(numberOfActiveFilters == 1) {
			// we can use LB_GETSELCOUNT
			if(properties.filterType[fpSelected] == ftIncluding) {
				if(CWindow(hWndLBox).GetStyle() & (LBS_MULTIPLESEL | LBS_EXTENDEDSEL)) {
					*pValue = static_cast<LONG>(SendMessage(hWndLBox, LB_GETSELCOUNT, 0, 0));
				} else {
					int anchorItem = static_cast<int>(SendMessage(hWndLBox, LB_GETANCHORINDEX, 0, 0));
					*pValue = (SendMessage(hWndLBox, LB_GETSEL, anchorItem, 0) ? 1 : 0);
				}
				return S_OK;
			} else {
				if(CWindow(hWndLBox).GetStyle() & (LBS_MULTIPLESEL | LBS_EXTENDEDSEL)) {
					*pValue = static_cast<LONG>(SendMessage(hWndLBox, LB_GETCOUNT, 0, 0)) - static_cast<LONG>(SendMessage(hWndLBox, LB_GETSELCOUNT, 0, 0));
				} else {
					int anchorItem = static_cast<int>(SendMessage(hWndLBox, LB_GETANCHORINDEX, 0, 0));
					*pValue = static_cast<LONG>(SendMessage(hWndLBox, LB_GETCOUNT, 0, 0)) - (SendMessage(hWndLBox, LB_GETSEL, anchorItem, 0) ? 1 : 0);
				}
				return S_OK;
			}
		}
	}

	// count the items "by hand"
	*pValue = 0;
	int numberOfItems = static_cast<int>(SendMessage(hWndLBox, LB_GETCOUNT, 0, 0));
	int itemIndex = GetFirstItemToProcess(numberOfItems, hWndLBox);
	while(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex, hWndLBox)) {
			++(*pValue);
		}
		itemIndex = GetNextItemToProcess(itemIndex, numberOfItems, hWndLBox);
	}
	return S_OK;
}

STDMETHODIMP ListBoxItems::Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/)
{
	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	// retrieve the item's index
	int itemIndex = -1;
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerLBox) {
				itemIndex = properties.pOwnerLBox->IDToItemIndex(itemIdentifier);
			}
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	if(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex)) {
			if(SendMessage(hWndLBox, LB_DELETESTRING, itemIndex, 0) != LB_ERR) {
				return S_OK;
			}
		}
	} else {
		// item not found
		if(itemIdentifierType == iitIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListBoxItems::RemoveAll(void)
{
	HWND hWndLBox = properties.GetLBoxHWnd();
	ATLASSERT(IsWindow(hWndLBox));

	if(!properties.usingFilters) {
		SendMessage(hWndLBox, LB_RESETCONTENT, 0, 0);
		return S_OK;
	}

	// find the items to remove manually
	#ifdef USE_STL
		std::list<int> itemsToRemove;
	#else
		CAtlList<int> itemsToRemove;
	#endif
	int numberOfItems = static_cast<int>(SendMessage(hWndLBox, LB_GETCOUNT, 0, 0));
	int itemIndex = GetFirstItemToProcess(numberOfItems, hWndLBox);
	while(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex, hWndLBox)) {
			#ifdef USE_STL
				itemsToRemove.push_back(itemIndex);
			#else
				itemsToRemove.AddTail(itemIndex);
			#endif
		}
		itemIndex = GetNextItemToProcess(itemIndex, numberOfItems, hWndLBox);
	}
	return RemoveItems(itemsToRemove, hWndLBox);
}