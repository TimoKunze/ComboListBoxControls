//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListBoxItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListBoxItemsEvents interface</em>
///
/// \if UNICODE
///   \sa ListBoxItems, CBLCtlsLibU::_IListBoxItemsEvents
/// \else
///   \sa ListBoxItems, CBLCtlsLibA::_IListBoxItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListBoxItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListBoxItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListBoxItemsEvents