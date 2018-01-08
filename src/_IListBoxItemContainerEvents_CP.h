//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListBoxItemContainerEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListBoxItemContainerEvents interface</em>
///
/// \if UNICODE
///   \sa ListBoxItemContainer, CBLCtlsLibU::_IListBoxItemContainerEvents
/// \else
///   \sa ListBoxItemContainer, CBLCtlsLibA::_IListBoxItemContainerEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListBoxItemContainerEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListBoxItemContainerEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListBoxItemContainerEvents