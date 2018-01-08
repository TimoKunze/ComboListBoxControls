//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListBoxItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListBoxItemEvents interface</em>
///
/// \if UNICODE
///   \sa ListBoxItem, CBLCtlsLibU::_IListBoxItemEvents
/// \else
///   \sa ListBoxItem, CBLCtlsLibA::_IListBoxItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListBoxItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListBoxItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListBoxItemEvents