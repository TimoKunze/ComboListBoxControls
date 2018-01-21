// DriveComboBoxItems.cpp: Manages a collection of DriveComboBoxItem objects

#include "stdafx.h"
#include "DriveComboBoxItems.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP DriveComboBoxItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IDriveComboBoxItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP DriveComboBoxItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<DriveComboBoxItems>* pDCBoxItemsObj = NULL;
	CComObject<DriveComboBoxItems>::CreateInstance(&pDCBoxItemsObj);
	pDCBoxItemsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pDCBoxItemsObj->properties);

	pDCBoxItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pDCBoxItemsObj->Release();
	return S_OK;
}

STDMETHODIMP DriveComboBoxItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	// check each item in the combo box
	int numberOfItems = static_cast<int>(SendMessage(hWndDCBox, CB_GETCOUNT, 0, 0));
	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		do {
			if(properties.lastEnumeratedItem >= 0) {
				if(properties.lastEnumeratedItem < numberOfItems) {
					properties.lastEnumeratedItem = GetNextItemToProcess(properties.lastEnumeratedItem, numberOfItems, hWndDCBox);
				}
			} else {
				properties.lastEnumeratedItem = GetFirstItemToProcess(numberOfItems, hWndDCBox);
			}
			if(properties.lastEnumeratedItem >= numberOfItems) {
				properties.lastEnumeratedItem = -1;
			}
		} while((properties.lastEnumeratedItem != -1) && (!IsPartOfCollection(properties.lastEnumeratedItem, hWndDCBox)));

		if(properties.lastEnumeratedItem != -1) {
			ClassFactory::InitDriveComboItem(properties.lastEnumeratedItem, properties.pOwnerDCBox, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
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

STDMETHODIMP DriveComboBoxItems::Reset(void)
{
	properties.lastEnumeratedItem = -1;
	return S_OK;
}

STDMETHODIMP DriveComboBoxItems::Skip(ULONG numberOfItemsToSkip)
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


BOOL DriveComboBoxItems::IsValidBooleanFilter(VARIANT& filter)
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

BOOL DriveComboBoxItems::IsValidIntegerFilter(VARIANT& filter, int minValue)
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

BOOL DriveComboBoxItems::IsValidIntegerFilter(VARIANT& filter)
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

BOOL DriveComboBoxItems::IsValidStringFilter(VARIANT& filter)
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

int DriveComboBoxItems::GetFirstItemToProcess(int numberOfItems, HWND /*hWndDCBox*/)
{
	//ATLASSERT(IsWindow(hWndDCBox));

	if(numberOfItems == 0) {
		return -1;
	}
	return 0;
}

int DriveComboBoxItems::GetNextItemToProcess(int previousItem, int numberOfItems, HWND /*hWndDCBox*/)
{
	//ATLASSERT(IsWindow(hWndDCBox));

	if(previousItem < numberOfItems - 1) {
		return previousItem + 1;
	}
	return -1;
}

BOOL DriveComboBoxItems::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
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

BOOL DriveComboBoxItems::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
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

BOOL DriveComboBoxItems::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
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

BOOL DriveComboBoxItems::IsPartOfCollection(int itemIndex, HWND hWndDCBox/* = NULL*/)
{
	if(!hWndDCBox) {
		hWndDCBox = properties.GetDCBoxHWnd();
	}
	ATLASSERT(IsWindow(hWndDCBox));

	if(!IsValidComboBoxItemIndex(itemIndex, hWndDCBox)) {
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
		VARIANT_BOOL b = BOOL2VARIANTBOOL(itemIndex == static_cast<int>(SendMessage(hWndDCBox, CB_GETCURSEL, 0, 0)));
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
		int textLength = static_cast<int>(SendMessage(hWndDCBox, CB_GETLBTEXTLEN, itemIndex, 0));
		if(textLength != CB_ERR) {
			LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
			if(pBuffer) {
				ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
				SendMessage(hWndDCBox, CB_GETLBTEXT, itemIndex, reinterpret_cast<LPARAM>(pBuffer));
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
	if(properties.filterType[fpPath] != ftDeactivated) {
		CComBSTR path = L"";
		CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(itemIndex, properties.pOwnerDCBox, FALSE);
		if(pDCBItem) {
			pDCBItem->get_Path(&path);
		}

		if(IsStringInSafeArray(properties.filter[fpPath].parray, path, properties.comparisonFunction[fpPath])) {
			if(properties.filterType[fpPath] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpPath] == ftIncluding) {
				goto Exit;
			}
		}
	}
	if(properties.filterType[fpDriveType] != ftDeactivated) {
		DriveTypeConstants driveType = static_cast<DriveTypeConstants>(0);
		CComPtr<IDriveComboBoxItem> pDCBItem = ClassFactory::InitDriveComboItem(itemIndex, properties.pOwnerDCBox, FALSE);
		if(pDCBItem) {
			pDCBItem->get_DriveType(&driveType);
		}

		if(IsIntegerInSafeArray(properties.filter[fpDriveType].parray, driveType, properties.comparisonFunction[fpDriveType])) {
			if(properties.filterType[fpDriveType] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpDriveType] == ftIncluding) {
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
		if(!SendMessage(hWndDCBox, CBEM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
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

void DriveComboBoxItems::OptimizeFilter(FilteredPropertyConstants filteredProperty)
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
	HRESULT DriveComboBoxItems::RemoveItems(std::list<int>& itemsToRemove, HWND hWndDCBox)
#else
	HRESULT DriveComboBoxItems::RemoveItems(CAtlList<int>& itemsToRemove, HWND hWndDCBox)
#endif
{
	ATLASSERT(IsWindow(hWndDCBox));

	CWindowEx2(hWndDCBox).InternalSetRedraw(FALSE);
	// sort in reverse order
	#ifdef USE_STL
		itemsToRemove.sort(std::greater<int>());
		for(std::list<int>::const_iterator iter = itemsToRemove.begin(); iter != itemsToRemove.end(); ++iter) {
			SendMessage(hWndDCBox, CBEM_DELETEITEM, *iter, 0);
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
			SendMessage(hWndDCBox, CBEM_DELETEITEM, itemsToRemove.GetAt(itemsToRemove.FindIndex(i)), 0);
		}
	#endif
	CWindowEx2(hWndDCBox).InternalSetRedraw(TRUE);

	return S_OK;
}


DriveComboBoxItems::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS_DRVCMB; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerDCBox) {
		pOwnerDCBox->Release();
	}
}

void DriveComboBoxItems::Properties::CopyTo(DriveComboBoxItems::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerDCBox = this->pOwnerDCBox;
		if(pTarget->pOwnerDCBox) {
			pTarget->pOwnerDCBox->AddRef();
		}
		pTarget->lastEnumeratedItem = this->lastEnumeratedItem;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS_DRVCMB; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND DriveComboBoxItems::Properties::GetDCBoxHWnd(void)
{
	ATLASSUME(pOwnerDCBox);

	OLE_HANDLE handle = NULL;
	pOwnerDCBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void DriveComboBoxItems::SetOwner(DriveComboBox* pOwner)
{
	if(properties.pOwnerDCBox) {
		properties.pOwnerDCBox->Release();
	}
	properties.pOwnerDCBox = pOwner;
	if(properties.pOwnerDCBox) {
		properties.pOwnerDCBox->AddRef();
	}
}


STDMETHODIMP DriveComboBoxItems::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP DriveComboBoxItems::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP DriveComboBoxItems::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
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
		case fpDriveType:
		case fpPath:
			*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP DriveComboBoxItems::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
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
		case fpDriveType:
		case fpPath:
			properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP DriveComboBoxItems::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
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
		case fpDriveType:
		case fpPath:
			VariantClear(pValue);
			VariantCopy(pValue, &properties.filter[filteredProperty]);
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP DriveComboBoxItems::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
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
		case fpDriveType:
			if(!IsValidIntegerFilter(newValue, 1)) {
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
		case fpPath:
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

STDMETHODIMP DriveComboBoxItems::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
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
		case fpDriveType:
		case fpPath:
			*pValue = properties.filterType[filteredProperty];
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP DriveComboBoxItems::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
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
		case fpDriveType:
		case fpPath:
			properties.filterType[filteredProperty] = newValue;
			if(newValue != ftDeactivated) {
				properties.usingFilters = TRUE;
			} else {
				properties.usingFilters = FALSE;
				for(int i = 0; i < NUMBEROFFILTERS_DRVCMB; ++i) {
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

STDMETHODIMP DriveComboBoxItems::get_Item(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/, IDriveComboBoxItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, IDriveComboBoxItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	// retrieve the item's index
	int itemIndex = -2;
	switch(itemIdentifierType) {
		case iitID:
			itemIndex = properties.pOwnerDCBox->IDToItemIndex(itemIdentifier);
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	if(itemIndex > -1) {
		if(IsPartOfCollection(itemIndex)) {
			ClassFactory::InitDriveComboItem(itemIndex, properties.pOwnerDCBox, IID_IDriveComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
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

STDMETHODIMP DriveComboBoxItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP DriveComboBoxItems::Contains(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// retrieve the item's index
	int itemIndex = -2;
	switch(itemIdentifierType) {
		case iitID:
			itemIndex = properties.pOwnerDCBox->IDToItemIndex(itemIdentifier);
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(itemIndex));
	return S_OK;
}

STDMETHODIMP DriveComboBoxItems::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	if(!properties.usingFilters) {
		*pValue = static_cast<LONG>(SendMessage(hWndDCBox, CB_GETCOUNT, 0, 0));
	}

	// count the items "by hand"
	*pValue = 0;
	int numberOfItems = static_cast<int>(SendMessage(hWndDCBox, CB_GETCOUNT, 0, 0));
	int itemIndex = GetFirstItemToProcess(numberOfItems, hWndDCBox);
	while(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex, hWndDCBox)) {
			++(*pValue);
		}
		itemIndex = GetNextItemToProcess(itemIndex, numberOfItems, hWndDCBox);
	}
	return S_OK;
}

STDMETHODIMP DriveComboBoxItems::Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/)
{
	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	// retrieve the item's index
	int itemIndex = -2;
	switch(itemIdentifierType) {
		case iitID:
			itemIndex = properties.pOwnerDCBox->IDToItemIndex(itemIdentifier);
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}

	if(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex)) {
			if(SendMessage(hWndDCBox, CBEM_DELETEITEM, itemIndex, 0) != -1) {
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

STDMETHODIMP DriveComboBoxItems::RemoveAll(void)
{
	HWND hWndDCBox = properties.GetDCBoxHWnd();
	ATLASSERT(IsWindow(hWndDCBox));

	if(!properties.usingFilters) {
		SendMessage(hWndDCBox, CB_RESETCONTENT, 0, 0);
		return S_OK;
	}

	// find the items to remove manually
	#ifdef USE_STL
		std::list<int> itemsToRemove;
	#else
		CAtlList<int> itemsToRemove;
	#endif
	int numberOfItems = static_cast<int>(SendMessage(hWndDCBox, CB_GETCOUNT, 0, 0));
	int itemIndex = GetFirstItemToProcess(numberOfItems, hWndDCBox);
	while(itemIndex != -1) {
		if(IsPartOfCollection(itemIndex, hWndDCBox)) {
			#ifdef USE_STL
				itemsToRemove.push_back(itemIndex);
			#else
				itemsToRemove.AddTail(itemIndex);
			#endif
		}
		itemIndex = GetNextItemToProcess(itemIndex, numberOfItems, hWndDCBox);
	}
	return RemoveItems(itemsToRemove, hWndDCBox);
}