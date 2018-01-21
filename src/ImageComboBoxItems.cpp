// ImageComboBoxItems.cpp: Manages a collection of ImageComboBoxItem objects

#include "stdafx.h"
#include "ImageComboBoxItems.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ImageComboBoxItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IImageComboBoxItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ImageComboBoxItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ImageComboBoxItems>* pICBoxItemsObj = NULL;
	CComObject<ImageComboBoxItems>::CreateInstance(&pICBoxItemsObj);
	pICBoxItemsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pICBoxItemsObj->properties);

	pICBoxItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pICBoxItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ImageComboBoxItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	// check each item in the combo box
	int numberOfItems = static_cast<int>(SendMessage(hWndICBox, CB_GETCOUNT, 0, 0));
	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		do {
			if(properties.lastEnumeratedItem >= 0) {
				if(properties.lastEnumeratedItem < numberOfItems) {
					properties.lastEnumeratedItem = GetNextItemToProcess(properties.lastEnumeratedItem, numberOfItems, hWndICBox);
				}
			} else {
				properties.lastEnumeratedItem = GetFirstItemToProcess(numberOfItems, hWndICBox);
			}
			if(properties.lastEnumeratedItem >= numberOfItems) {
				properties.lastEnumeratedItem = -1;
			}
		} while((properties.lastEnumeratedItem != -1) && (!IsPartOfCollection(properties.lastEnumeratedItem, hWndICBox)));

		if(properties.lastEnumeratedItem != -1) {
			ClassFactory::InitImageComboItem(properties.lastEnumeratedItem, properties.pOwnerICBox, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
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

STDMETHODIMP ImageComboBoxItems::Reset(void)
{
	properties.lastEnumeratedItem = -1;
	return S_OK;
}

STDMETHODIMP ImageComboBoxItems::Skip(ULONG numberOfItemsToSkip)
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


BOOL ImageComboBoxItems::IsValidBooleanFilter(VARIANT& filter)
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

BOOL ImageComboBoxItems::IsValidIntegerFilter(VARIANT& filter, int minValue)
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

BOOL ImageComboBoxItems::IsValidIntegerFilter(VARIANT& filter)
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

BOOL ImageComboBoxItems::IsValidStringFilter(VARIANT& filter)
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

int ImageComboBoxItems::GetFirstItemToProcess(int numberOfItems, HWND /*hWndICBox*/)
{
	//ATLASSERT(IsWindow(hWndICBox));

	if(numberOfItems == 0) {
		return -1;
	}
	return 0;
}

int ImageComboBoxItems::GetNextItemToProcess(int previousItem, int numberOfItems, HWND /*hWndICBox*/)
{
	//ATLASSERT(IsWindow(hWndICBox));

	if(previousItem < numberOfItems - 1) {
		return previousItem + 1;
	}
	return -1;
}

BOOL ImageComboBoxItems::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
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

BOOL ImageComboBoxItems::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
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

BOOL ImageComboBoxItems::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
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

BOOL ImageComboBoxItems::IsPartOfCollection(int itemIndex, HWND hWndICBox/* = NULL*/)
{
	if(!hWndICBox) {
		hWndICBox = properties.GetICBoxHWnd();
	}
	ATLASSERT(IsWindow(hWndICBox));

	if(!IsValidComboBoxItemIndex(itemIndex, hWndICBox)) {
		return FALSE;
	}

	BOOL isPart = FALSE;
	// we declare this one here to avoid compiler warnings
	COMBOBOXEXITEM item = {0};

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
		VARIANT_BOOL b = BOOL2VARIANTBOOL(itemIndex == static_cast<int>(SendMessage(hWndICBox, CB_GETCURSEL, 0, 0)));
		if(IsBooleanInSafeArray(properties.filter[fpSelected].parray, b, properties.comparisonFunction[fpSelected])) {
			if(properties.filterType[fpSelected] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpSelected] == ftIncluding) {
				goto Exit;
			}
		}
	}
	if(properties.filterType[fpText] != ftDeactivated) {
		CComBSTR text = L"";
		int textLength = static_cast<int>(SendMessage(hWndICBox, CB_GETLBTEXTLEN, itemIndex, 0));
		if(textLength != CB_ERR) {
			LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
			if(pBuffer) {
				ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
				SendMessage(hWndICBox, CB_GETLBTEXT, itemIndex, reinterpret_cast<LPARAM>(pBuffer));
				text = pBuffer;
				HeapFree(GetProcessHeap(), 0, pBuffer);
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

	item.iItem = itemIndex;
	BOOL mustRetrieveItemData = FALSE;
	if(properties.filterType[fpIconIndex] != ftDeactivated) {
		item.mask |= CBEIF_IMAGE;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpSelectedIconIndex] != ftDeactivated) {
		item.mask |= CBEIF_SELECTEDIMAGE;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpOverlayIndex] != ftDeactivated) {
		item.mask |= CBEIF_OVERLAY;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpIndent] != ftDeactivated) {
		item.mask |= CBEIF_INDENT;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpItemData] != ftDeactivated) {
		item.mask |= CBEIF_LPARAM;
		mustRetrieveItemData = TRUE;
	}

	if(mustRetrieveItemData) {
		if(!SendMessage(hWndICBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			goto Exit;
		}

		// apply the filters

		if(properties.filterType[fpIconIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpIconIndex].parray, item.iImage, properties.comparisonFunction[fpIconIndex])) {
				if(properties.filterType[fpIconIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIconIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpSelectedIconIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpSelectedIconIndex].parray, item.iSelectedImage, properties.comparisonFunction[fpSelectedIconIndex])) {
				if(properties.filterType[fpSelectedIconIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpSelectedIconIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpOverlayIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpOverlayIndex].parray, item.iOverlay, properties.comparisonFunction[fpOverlayIndex])) {
				if(properties.filterType[fpOverlayIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpOverlayIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpIndent] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpIndent].parray, item.iIndent, properties.comparisonFunction[fpIndent])) {
				if(properties.filterType[fpIndent] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIndent] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpItemData] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpItemData].parray, static_cast<int>(item.lParam), properties.comparisonFunction[fpItemData])) {
				if(properties.filterType[fpItemData] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpItemData] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	return isPart;
}

void ImageComboBoxItems::OptimizeFilter(FilteredPropertyConstants filteredProperty)
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
	HRESULT ImageComboBoxItems::RemoveItems(std::list<int>& itemsToRemove, HWND hWndICBox)
#else
	HRESULT ImageComboBoxItems::RemoveItems(CAtlList<int>& itemsToRemove, HWND hWndICBox)
#endif
{
	ATLASSERT(IsWindow(hWndICBox));

	CWindowEx2(hWndICBox).InternalSetRedraw(FALSE);
	// sort in reverse order
	#ifdef USE_STL
		itemsToRemove.sort(std::greater<int>());
		for(std::list<int>::const_iterator iter = itemsToRemove.begin(); iter != itemsToRemove.end(); ++iter) {
			SendMessage(hWndICBox, CBEM_DELETEITEM, *iter, 0);
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
			SendMessage(hWndICBox, CBEM_DELETEITEM, itemsToRemove.GetAt(itemsToRemove.FindIndex(i)), 0);
		}
	#endif
	CWindowEx2(hWndICBox).InternalSetRedraw(TRUE);

	return S_OK;
}


ImageComboBoxItems::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS_IMGCMB; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerICBox) {
		pOwnerICBox->Release();
	}
}

void ImageComboBoxItems::Properties::CopyTo(ImageComboBoxItems::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerICBox = this->pOwnerICBox;
		if(pTarget->pOwnerICBox) {
			pTarget->pOwnerICBox->AddRef();
		}
		pTarget->lastEnumeratedItem = this->lastEnumeratedItem;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS_IMGCMB; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND ImageComboBoxItems::Properties::GetICBoxHWnd(void)
{
	ATLASSUME(pOwnerICBox);

	OLE_HANDLE handle = NULL;
	pOwnerICBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ImageComboBoxItems::SetOwner(ImageComboBox* pOwner)
{
	if(properties.pOwnerICBox) {
		properties.pOwnerICBox->Release();
	}
	properties.pOwnerICBox = pOwner;
	if(properties.pOwnerICBox) {
		properties.pOwnerICBox->AddRef();
	}
}


STDMETHODIMP ImageComboBoxItems::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP ImageComboBoxItems::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP ImageComboBoxItems::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndent:
		case fpIndex:
		case fpItemData:
		case fpOverlayIndex:
		case fpSelected:
		case fpText:
		case fpSelectedIconIndex:
			*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ImageComboBoxItems::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
{
	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndent:
		case fpIndex:
		case fpItemData:
		case fpOverlayIndex:
		case fpSelected:
		case fpText:
		case fpSelectedIconIndex:
			properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ImageComboBoxItems::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndent:
		case fpIndex:
		case fpItemData:
		case fpOverlayIndex:
		case fpSelected:
		case fpText:
		case fpSelectedIconIndex:
			VariantClear(pValue);
			VariantCopy(pValue, &properties.filter[filteredProperty]);
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ImageComboBoxItems::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	// check 'newValue'
	switch(filteredProperty) {
		case fpSelected:
			if(!IsValidBooleanFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpIconIndex:
		case fpSelectedIconIndex:
			if(!IsValidIntegerFilter(newValue, -1)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpIndent:
		case fpIndex:
		case fpOverlayIndex:
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

STDMETHODIMP ImageComboBoxItems::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndent:
		case fpIndex:
		case fpItemData:
		case fpOverlayIndex:
		case fpSelected:
		case fpText:
		case fpSelectedIconIndex:
			*pValue = properties.filterType[filteredProperty];
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ImageComboBoxItems::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if(newValue < 0 || newValue > 2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndent:
		case fpIndex:
		case fpItemData:
		case fpOverlayIndex:
		case fpSelected:
		case fpText:
		case fpSelectedIconIndex:
			properties.filterType[filteredProperty] = newValue;
			if(newValue != ftDeactivated) {
				properties.usingFilters = TRUE;
			} else {
				properties.usingFilters = FALSE;
				for(int i = 0; i < NUMBEROFFILTERS_IMGCMB; ++i) {
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

STDMETHODIMP ImageComboBoxItems::get_Item(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/, IImageComboBoxItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, IImageComboBoxItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	// retrieve the item's index
	int itemIndex = -2;
	switch(itemIdentifierType) {
		case iitID:
			itemIndex = properties.pOwnerICBox->IDToItemIndex(itemIdentifier);
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	if(itemIndex > -1) {
		if(IsPartOfCollection(itemIndex)) {
			ClassFactory::InitImageComboItem(itemIndex, properties.pOwnerICBox, IID_IImageComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
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

STDMETHODIMP ImageComboBoxItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ImageComboBoxItems::Add(BSTR itemText, LONG insertAt/* = -1*/, LONG iconIndex/* = -2*/, LONG selectedIconIndex/* = -2*/, LONG overlayIndex/* = 0*/, LONG itemIndentation/* = 0*/, LONG itemData/* = 0*/, IImageComboBoxItem** ppAddedItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedItem, IImageComboBoxItem*);
	if(!ppAddedItem) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	if(insertAt < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	COMBOBOXEXITEM insertionData = {0};
	insertionData.iItem = insertAt;
	insertionData.mask = CBEIF_INDENT | CBEIF_LPARAM | CBEIF_TEXT;
	if(iconIndex != -2) {
		insertionData.iImage = iconIndex;
		insertionData.mask |= CBEIF_IMAGE;
	}
	if(selectedIconIndex != -2) {
		insertionData.iSelectedImage = selectedIconIndex;
		insertionData.mask |= CBEIF_SELECTEDIMAGE;
	} else if(iconIndex != -2) {
		insertionData.iSelectedImage = insertionData.iImage;
		insertionData.mask |= CBEIF_SELECTEDIMAGE;
	}
	if(overlayIndex != 0) {
		insertionData.iOverlay |= overlayIndex;
		insertionData.mask |= CBEIF_OVERLAY;
	}
	insertionData.iIndent = itemIndentation;
	insertionData.lParam = itemData;

	LPTSTR pItemText;
	#ifndef UNICODE
		COLE2T converter(itemText);
	#endif
	if(itemText) {
		#ifdef UNICODE
			pItemText = OLE2W(itemText);
		#else
			pItemText = converter;
		#endif
	} else {
		pItemText = LPSTR_TEXTCALLBACK;
	}
	insertionData.pszText = pItemText;

	// finally insert the tab
	*ppAddedItem = NULL;
	int itemIndex = SendMessage(hWndICBox, CBEM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));
	if(itemIndex != -1) {
		ClassFactory::InitImageComboItem(itemIndex, properties.pOwnerICBox, IID_IImageComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppAddedItem), FALSE);
		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP ImageComboBoxItems::Contains(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// retrieve the item's index
	int itemIndex = -2;
	switch(itemIdentifierType) {
		case iitID:
			itemIndex = properties.pOwnerICBox->IDToItemIndex(itemIdentifier);
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(itemIndex));
	return S_OK;
}

STDMETHODIMP ImageComboBoxItems::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	if(!properties.usingFilters) {
		*pValue = static_cast<LONG>(SendMessage(hWndICBox, CB_GETCOUNT, 0, 0));
	}

	// count the items "by hand"
	*pValue = 0;
	int numberOfItems = static_cast<int>(SendMessage(hWndICBox, CB_GETCOUNT, 0, 0));
	int itemIndex = GetFirstItemToProcess(numberOfItems, hWndICBox);
	while(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex, hWndICBox)) {
			++(*pValue);
		}
		itemIndex = GetNextItemToProcess(itemIndex, numberOfItems, hWndICBox);
	}
	return S_OK;
}

STDMETHODIMP ImageComboBoxItems::Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/)
{
	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	// retrieve the item's index
	int itemIndex = -2;
	switch(itemIdentifierType) {
		case iitID:
			itemIndex = properties.pOwnerICBox->IDToItemIndex(itemIdentifier);
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	if(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex)) {
			if(SendMessage(hWndICBox, CBEM_DELETEITEM, itemIndex, 0) != -1) {
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

STDMETHODIMP ImageComboBoxItems::RemoveAll(void)
{
	HWND hWndICBox = properties.GetICBoxHWnd();
	ATLASSERT(IsWindow(hWndICBox));

	if(!properties.usingFilters) {
		SendMessage(hWndICBox, CB_RESETCONTENT, 0, 0);
		return S_OK;
	}

	// find the items to remove manually
	#ifdef USE_STL
		std::list<int> itemsToRemove;
	#else
		CAtlList<int> itemsToRemove;
	#endif
	int numberOfItems = static_cast<int>(SendMessage(hWndICBox, CB_GETCOUNT, 0, 0));
	int itemIndex = GetFirstItemToProcess(numberOfItems, hWndICBox);
	while(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex, hWndICBox)) {
			#ifdef USE_STL
				itemsToRemove.push_back(itemIndex);
			#else
				itemsToRemove.AddTail(itemIndex);
			#endif
		}
		itemIndex = GetNextItemToProcess(itemIndex, numberOfItems, hWndICBox);
	}
	return RemoveItems(itemsToRemove, hWndICBox);
}