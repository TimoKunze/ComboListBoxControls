// ImageComboBoxItemContainer.cpp: Manages a collection of ImageComboBoxItem objects

#include "stdafx.h"
#include "ImageComboBoxItemContainer.h"
#include "ClassFactory.h"


DWORD ImageComboBoxItemContainer::nextID = 0;


ImageComboBoxItemContainer::ImageComboBoxItemContainer()
{
	containerID = ++nextID;
}

ImageComboBoxItemContainer::~ImageComboBoxItemContainer()
{
	properties.pOwnerICBox->DeregisterItemContainer(containerID);
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP ImageComboBoxItemContainer::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IImageComboBoxItemContainer, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ImageComboBoxItemContainer::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ImageComboBoxItemContainer>* pICBItemContainerObj = NULL;
	CComObject<ImageComboBoxItemContainer>::CreateInstance(&pICBItemContainerObj);
	pICBItemContainerObj->AddRef();

	// clone all settings
	pICBItemContainerObj->properties = properties;
	if(pICBItemContainerObj->properties.pOwnerICBox) {
		pICBItemContainerObj->properties.pOwnerICBox->AddRef();
	}
	#ifdef USE_STL
		pICBItemContainerObj->properties.items.resize(properties.items.size());
		std::copy(properties.items.begin(), properties.items.end(), pICBItemContainerObj->properties.items.begin());
	#else
		//pICBItemContainerObj->properties.items.Copy(properties.items);
		pICBItemContainerObj->items.Copy(items);
	#endif

	pICBItemContainerObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pICBItemContainerObj->Release();

	if(*ppEnumerator) {
		properties.pOwnerICBox->RegisterItemContainer(static_cast<IItemContainer*>(pICBItemContainerObj));
	}
	return S_OK;
}

STDMETHODIMP ImageComboBoxItemContainer::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		#ifdef USE_STL
			if(properties.nextEnumeratedItem >= static_cast<int>(properties.items.size())) {
				properties.nextEnumeratedItem = 0;
				// there's nothing more to iterate
				break;
			}
			LONG itemID = properties.items[properties.nextEnumeratedItem];
		#else
			//if(properties.nextEnumeratedItem >= static_cast<int>(properties.items.GetCount())) {
			if(properties.nextEnumeratedItem >= static_cast<int>(items.GetCount())) {
				properties.nextEnumeratedItem = 0;
				// there's nothing more to iterate
				break;
			}
			//LONG itemID = properties.items[properties.nextEnumeratedItem];
			LONG itemID = items[properties.nextEnumeratedItem];
		#endif
		int itemIndex = properties.pOwnerICBox->IDToItemIndex(itemID);
		if(itemIndex == -1) {
			CComPtr<IImageComboBoxItem> p;
			properties.pOwnerICBox->get_SelectionFieldItem(&p);
			if(p) {
				p->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&pItems[i].pdispVal));
			}
		} else if(itemIndex != -2) {
			ClassFactory::InitImageComboItem(itemIndex, properties.pOwnerICBox, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
		}
		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedItem;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ImageComboBoxItemContainer::Reset(void)
{
	properties.nextEnumeratedItem = 0;
	return S_OK;
}

STDMETHODIMP ImageComboBoxItemContainer::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedItem += numberOfItemsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IItemContainer
void ImageComboBoxItemContainer::RemovedItem(LONG itemIdentifier)
{
	if(itemIdentifier == -1) {
		RemoveAll();
	} else {
		Remove(itemIdentifier);
	}
}

void ImageComboBoxItemContainer::ReplacedItemID(LONG oldItemID, LONG newItemID)
{
	#ifdef USE_STL
		for(size_t i = 0; i < properties.items.size(); ++i) {
			if(properties.items[i] == oldItemID) {
				properties.items[i] = newItemID;
				break;
			}
		}
	#else
		//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
		for(size_t i = 0; i < items.GetCount(); ++i) {
			//if(properties.items[i] == oldItemID) {
			if(items[i] == oldItemID) {
				//properties.items[i] = newItemID;
				items[i] = newItemID;
				break;
			}
		}
	#endif
}

DWORD ImageComboBoxItemContainer::GetID(void)
{
	return containerID;
}
// implementation of IItemContainer
//////////////////////////////////////////////////////////////////////


ImageComboBoxItemContainer::Properties::~Properties()
{
	if(pOwnerICBox) {
		pOwnerICBox->Release();
	}
}

HWND ImageComboBoxItemContainer::Properties::GetICBoxHWnd(void)
{
	ATLASSUME(pOwnerICBox);

	OLE_HANDLE handle = NULL;
	pOwnerICBox->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ImageComboBoxItemContainer::SetOwner(ImageComboBox* pOwner)
{
	if(properties.pOwnerICBox) {
		properties.pOwnerICBox->Release();
	}
	properties.pOwnerICBox = pOwner;
	if(properties.pOwnerICBox) {
		properties.pOwnerICBox->AddRef();
	}
}


STDMETHODIMP ImageComboBoxItemContainer::get_Item(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitID*/, IImageComboBoxItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, IImageComboBoxItem*);
	if(!ppItem) {
		return E_POINTER;
	}
	if(itemIdentifierType != iitID) {
		return E_INVALIDARG;
	}

	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(properties.items.begin(), properties.items.end(), itemIdentifier);
		if(iter != properties.items.end()) {
			int itemIndex = properties.pOwnerICBox->IDToItemIndex(itemIdentifier);
			if(itemIndex == -1) {
				return properties.pOwnerICBox->get_SelectionFieldItem(ppItem);
			} else if(itemIndex != -2) {
				ClassFactory::InitImageComboItem(itemIndex, properties.pOwnerICBox, IID_IImageComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
				return S_OK;
			}
		}
	#else
		//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
		for(size_t i = 0; i < items.GetCount(); ++i) {
			//if(properties.items[i] == itemIdentifier) {
			if(items[i] == itemIdentifier) {
				int itemIndex = properties.pOwnerICBox->IDToItemIndex(itemIdentifier);
				if(itemIndex == -1) {
					return properties.pOwnerICBox->get_SelectionFieldItem(ppItem);
				} else if(itemIndex != -2) {
					ClassFactory::InitImageComboItem(itemIndex, properties.pOwnerICBox, IID_IImageComboBoxItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
					return S_OK;
				}
				break;
			}
		}
	#endif

	return E_INVALIDARG;
}

STDMETHODIMP ImageComboBoxItemContainer::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ImageComboBoxItemContainer::Add(VARIANT items)
{
	HRESULT hr = E_FAIL;
	LONG id = -1;
	switch(items.vt) {
		case VT_DISPATCH:
			if(items.pdispVal) {
				CComQIPtr<IImageComboBoxItem, &IID_IImageComboBoxItem> pICBItem = items.pdispVal;
				if(pICBItem) {
					// add a single ImageComboBoxItem object
					hr = pICBItem->get_ID(&id);
				} else {
					CComQIPtr<IImageComboBoxItems, &IID_IImageComboBoxItems> pICBItems(items.pdispVal);
					if(pICBItems) {
						// add a ImageComboBoxItems collection
						CComQIPtr<IEnumVARIANT, &IID_IEnumVARIANT> pEnumerator(pICBItems);
						if(pEnumerator) {
							hr = S_OK;
							VARIANT item;
							VariantInit(&item);
							ULONG dummy = 0;
							while(pEnumerator->Next(1, &item, &dummy) == S_OK) {
								if((item.vt == VT_DISPATCH) && item.pdispVal) {
									pICBItem = item.pdispVal;
									if(pICBItem) {
										id = -1;
										pICBItem->get_ID(&id);
										#ifdef USE_STL
											std::vector<LONG>::iterator iter = std::find(properties.items.begin(), properties.items.end(), id);
											if(iter == properties.items.end()) {
												properties.items.push_back(id);
											}
										#else
											BOOL alreadyThere = FALSE;
											//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
											for(size_t i = 0; i < this->items.GetCount(); ++i) {
												//if(properties.items[i] == id) {
												if(this->items[i] == id) {
													alreadyThere = TRUE;
													break;
												}
											}
											if(!alreadyThere) {
												//properties.tabs.Add(id);
												this->items.Add(id);
											}
										#endif
									}
								}
								VariantClear(&item);
							}
							return S_OK;
						}
					}
				}
			}
			break;
		default:
			VARIANT v;
			VariantInit(&v);
			hr = VariantChangeType(&v, &items, 0, VT_UI4);
			id = v.ulVal;
			break;
	}
	if(FAILED(hr)) {
		// invalid arg - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(properties.items.begin(), properties.items.end(), id);
		if(iter == properties.items.end()) {
			properties.items.push_back(id);
		}
	#else
		BOOL alreadyThere = FALSE;
		//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
		for(size_t i = 0; i < this->items.GetCount(); ++i) {
			//if(properties.items[i] == id) {
			if(this->items[i] == id) {
				alreadyThere = TRUE;
				break;
			}
		}
		if(!alreadyThere) {
			//properties.items.Add(id);
			this->items.Add(id);
		}
	#endif
	return S_OK;
}

STDMETHODIMP ImageComboBoxItemContainer::Clone(IImageComboBoxItemContainer** ppClone)
{
	ATLASSERT_POINTER(ppClone, IImageComboBoxItemContainer*);
	if(!ppClone) {
		return E_POINTER;
	}

	*ppClone = NULL;
	CComObject<ImageComboBoxItemContainer>* pICBItemContainerObj = NULL;
	CComObject<ImageComboBoxItemContainer>::CreateInstance(&pICBItemContainerObj);
	pICBItemContainerObj->AddRef();

	// clone all settings
	pICBItemContainerObj->properties = properties;
	if(pICBItemContainerObj->properties.pOwnerICBox) {
		pICBItemContainerObj->properties.pOwnerICBox->AddRef();
	}
	#ifdef USE_STL
		pICBItemContainerObj->properties.items.resize(properties.items.size());
		std::copy(properties.items.begin(), properties.items.end(), pICBItemContainerObj->properties.items.begin());
	#else
		//pICBItemContainerObj->properties.items.Copy(properties.items);
		pICBItemContainerObj->items.Copy(items);
	#endif

	pICBItemContainerObj->QueryInterface(__uuidof(IImageComboBoxItemContainer), reinterpret_cast<LPVOID*>(ppClone));
	pICBItemContainerObj->Release();

	if(*ppClone) {
		properties.pOwnerICBox->RegisterItemContainer(static_cast<IItemContainer*>(pICBItemContainerObj));
	}
	return S_OK;
}

STDMETHODIMP ImageComboBoxItemContainer::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef USE_STL
		*pValue = static_cast<LONG>(properties.items.size());
	#else
		//*pValue = static_cast<LONG>(properties.items.GetCount());
		*pValue = static_cast<LONG>(items.GetCount());
	#endif
	return S_OK;
}

STDMETHODIMP ImageComboBoxItemContainer::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	*phImageList = NULL;
	#ifdef USE_STL
		switch(properties.items.size()) {
	#else
		//switch(properties.items.GetCount()) {
		switch(items.GetCount()) {
	#endif
		case 0:
			return S_OK;
			break;
		case 1: {
			ATLASSUME(properties.pOwnerICBox);
			//int itemIndex = properties.pOwnerICBox->IDToItemIndex(properties.items[0]);
			#ifdef USE_STL
				int itemIndex = properties.pOwnerICBox->IDToItemIndex(properties.items[0]);
			#else
				int itemIndex = properties.pOwnerICBox->IDToItemIndex(items[0]);
			#endif
			if(itemIndex != -2) {
				POINT upperLeftPoint = {0};
				*phImageList = HandleToLong(properties.pOwnerICBox->CreateLegacyDragImage(itemIndex, &upperLeftPoint, NULL));
				if(*phImageList) {
					if(pXUpperLeft) {
						*pXUpperLeft = upperLeftPoint.x;
					}
					if(pYUpperLeft) {
						*pYUpperLeft = upperLeftPoint.y;
					}
					return S_OK;
				}
			}
			break;
		}
		default: {
			// create a large drag image out of small drag images
			ATLASSUME(properties.pOwnerICBox);

			BOOL use32BPPImage = RunTimeHelper::IsCommCtrl6();

			// calculate the bitmap's required size and collect each item's image list
			#ifdef USE_STL
				std::vector<HIMAGELIST> imageLists;
				std::vector<RECT> itemBoundingRects;
			#else
				CAtlArray<HIMAGELIST> imageLists;
				CAtlArray<RECT> itemBoundingRects;
			#endif
			POINT upperLeftPoint = {0};
			WTL::CRect boundingRect;
			#ifdef USE_STL
				for(std::vector<LONG>::iterator iter = properties.items.begin(); iter != properties.items.end(); ++iter) {
					int itemIndex = properties.pOwnerICBox->IDToItemIndex(*iter);
			#else
				//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
				for(size_t i = 0; i < items.GetCount(); ++i) {
					//int itemIndex = properties.pOwnerICBox->IDToItemIndex(properties.items[i]);
					int itemIndex = properties.pOwnerICBox->IDToItemIndex(items[i]);
			#endif
				if(itemIndex != -2) {
					// NOTE: Windows skips items outside the client area to improve performance. We don't.
					POINT pt = {0};
					RECT itemBoundingRect = {0};
					HIMAGELIST hImageList = properties.pOwnerICBox->CreateLegacyDragImage(itemIndex, &pt, &itemBoundingRect);
					boundingRect.UnionRect(&boundingRect, &itemBoundingRect);

					#ifdef USE_STL
						if(imageLists.size() == 0) {
					#else
						if(imageLists.GetCount() == 0) {
					#endif
						upperLeftPoint = pt;
					} else {
						upperLeftPoint.x = min(upperLeftPoint.x, pt.x);
						upperLeftPoint.y = min(upperLeftPoint.y, pt.y);
					}
					#ifdef USE_STL
						imageLists.push_back(hImageList);
						itemBoundingRects.push_back(itemBoundingRect);
					#else
						imageLists.Add(hImageList);
						itemBoundingRects.Add(itemBoundingRect);
					#endif
				}
			}
			WTL::CRect dragImageRect(0, 0, boundingRect.Width(), boundingRect.Height());

			// setup the DCs we'll draw into
			HDC hCompatibleDC = GetDC(HWND_DESKTOP);
			CDC memoryDC;
			memoryDC.CreateCompatibleDC(hCompatibleDC);
			CDC maskMemoryDC;
			maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

			// create the bitmap and its mask
			CBitmap dragImage;
			LPRGBQUAD pDragImageBits = NULL;
			if(use32BPPImage) {
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = dragImageRect.Width();
				bitmapInfo.bmiHeader.biHeight = -dragImageRect.Height();
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;

				dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
			} else {
				dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageRect.Width(), dragImageRect.Height());
			}
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
			memoryDC.FillRect(&dragImageRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
			CBitmap dragImageMask;
			dragImageMask.CreateBitmap(dragImageRect.Width(), dragImageRect.Height(), 1, 1, NULL);
			HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);
			maskMemoryDC.FillRect(&dragImageRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

			// draw each single drag image into our bitmap
			BOOL rightToLeft = FALSE;
			HWND hWndImageComboBox = properties.GetICBoxHWnd();
			if(IsWindow(hWndImageComboBox)) {
				rightToLeft = ((CWindow(hWndImageComboBox).GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
			}
			#ifdef USE_STL
				for(size_t i = 0; i < imageLists.size(); ++i) {
			#else
				for(size_t i = 0; i < imageLists.GetCount(); ++i) {
			#endif
				if(rightToLeft) {
					ImageList_Draw(imageLists[i], 0, memoryDC, boundingRect.right - itemBoundingRects[i].right, itemBoundingRects[i].top - boundingRect.top, ILD_NORMAL);
					ImageList_Draw(imageLists[i], 0, maskMemoryDC, boundingRect.right - itemBoundingRects[i].right, itemBoundingRects[i].top - boundingRect.top, ILD_MASK);
				} else {
					ImageList_Draw(imageLists[i], 0, memoryDC, itemBoundingRects[i].left - boundingRect.left, itemBoundingRects[i].top - boundingRect.top, ILD_NORMAL);
					ImageList_Draw(imageLists[i], 0, maskMemoryDC, itemBoundingRects[i].left - boundingRect.left, itemBoundingRects[i].top - boundingRect.top, ILD_MASK);
				}

				ImageList_Destroy(imageLists[i]);
			}

			// clean up
			#ifdef USE_STL
				imageLists.clear();
				itemBoundingRects.clear();
			#else
				imageLists.RemoveAll();
				itemBoundingRects.RemoveAll();
			#endif
			memoryDC.SelectBitmap(hPreviousBitmap);
			maskMemoryDC.SelectBitmap(hPreviousBitmapMask);
			ReleaseDC(HWND_DESKTOP, hCompatibleDC);

			// create the imagelist
			HIMAGELIST hImageList = ImageList_Create(dragImageRect.Width(), dragImageRect.Height(), (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
			ImageList_SetBkColor(hImageList, CLR_NONE);
			ImageList_Add(hImageList, dragImage, dragImageMask);
			*phImageList = HandleToLong(hImageList);

			if(*phImageList) {
				if(pXUpperLeft) {
					*pXUpperLeft = upperLeftPoint.x;
				}
				if(pYUpperLeft) {
					*pYUpperLeft = upperLeftPoint.y;
				}
				return S_OK;
			}
			break;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItemContainer::Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitID*/, VARIANT_BOOL removePhysically/* = VARIANT_FALSE*/)
{
	if(itemIdentifierType != iitID) {
		return E_INVALIDARG;
	}

	#ifdef USE_STL
		for(size_t i = 0; i < properties.items.size(); ++i) {
			if(properties.items[i] == itemIdentifier) {
				if(removePhysically == VARIANT_FALSE) {
					properties.items.erase(std::find(properties.items.begin(), properties.items.end(), itemIdentifier));
					if(i < (size_t) properties.nextEnumeratedItem) {
						--properties.nextEnumeratedItem;
					}
				} else {
					HWND hWndImageComboBox = properties.GetICBoxHWnd();
					if(IsWindow(hWndImageComboBox)) {
						int itemIndex = properties.pOwnerICBox->IDToItemIndex(itemIdentifier);
						if(itemIndex >= 0) {
							SendMessage(hWndImageComboBox, CBEM_DELETEITEM, itemIndex, 0);
						}
					}

					// our owner will notify us about the deletion, so we don't need to remove the item explicitly
				}

				return S_OK;
			}
		}
	#else
		//for(size_t i = 0; i < properties.items.GetCount(); ++i) {
		for(size_t i = 0; i < items.GetCount(); ++i) {
			//if(properties.items[i] == itemIdentifier) {
			if(items[i] == itemIdentifier) {
				if(removePhysically == VARIANT_FALSE) {
					//properties.items.RemoveAt(i);
					items.RemoveAt(i);
					if(i < (size_t) properties.nextEnumeratedItem) {
						--properties.nextEnumeratedItem;
					}
				} else {
					HWND hWndImageComboBox = properties.GetICBoxHWnd();
					if(IsWindow(hWndImageComboBox)) {
						int itemIndex = properties.pOwnerICBox->IDToItemIndex(itemIdentifier);
						if(itemIndex >= 0) {
							SendMessage(hWndImageComboBox, CBEM_DELETEITEM, itemIndex, 0);
						}
					}

					// our owner will notify us about the deletion, so we don't need to remove the item explicitly
				}

				return S_OK;
			}
		}
	#endif
	return E_FAIL;
}

STDMETHODIMP ImageComboBoxItemContainer::RemoveAll(VARIANT_BOOL removePhysically/* = VARIANT_FALSE*/)
{
	if(removePhysically != VARIANT_FALSE) {
		#ifdef USE_STL
			properties.items.clear();
		#else
			//properties.items.RemoveAll();
			items.RemoveAll();
		#endif
	} else {
		HWND hWndImageComboBox = properties.GetICBoxHWnd();
		if(IsWindow(hWndImageComboBox)) {
			#ifdef USE_STL
				while(properties.items.size() > 0) {
					int itemIndex = properties.pOwnerICBox->IDToItemIndex(*properties.items.begin());
					if(itemIndex >= 0) {
						SendMessage(hWndImageComboBox, CBEM_DELETEITEM, itemIndex, 0);
					}
				}
			#else
				//while(properties.items.GetCount() > 0) {
				while(items.GetCount() > 0) {
					//int itemIndex = properties.pOwnerICBox->IDToItemIndex(properties.items[0]);
					int itemIndex = properties.pOwnerICBox->IDToItemIndex(items[0]);
					if(itemIndex >= 0) {
						SendMessage(hWndImageComboBox, CBEM_DELETEITEM, itemIndex, 0);
					}
				}
			#endif
		}
	}
	Reset();
	return S_OK;
}