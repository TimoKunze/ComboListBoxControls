//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualListBoxItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualListBoxItemEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualListBoxItem, CBLCtlsLibU::_IVirtualListBoxItemEvents
/// \else
///   \sa VirtualListBoxItem, CBLCtlsLibA::_IVirtualListBoxItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualListBoxItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualListBoxItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualListBoxItemEvents