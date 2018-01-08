//////////////////////////////////////////////////////////////////////
/// \class Proxy_IImageComboBoxItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IImageComboBoxItemsEvents interface</em>
///
/// \if UNICODE
///   \sa ImageComboBoxItems, CBLCtlsLibU::_IImageComboBoxItemsEvents
/// \else
///   \sa ImageComboBoxItems, CBLCtlsLibA::_IImageComboBoxItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IImageComboBoxItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IImageComboBoxItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IImageComboBoxItemsEvents