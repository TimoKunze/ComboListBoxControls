//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualComboBoxItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualComboBoxItemEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualComboBoxItem, CBLCtlsLibU::_IVirtualComboBoxItemEvents
/// \else
///   \sa VirtualComboBoxItem, CBLCtlsLibA::_IVirtualComboBoxItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualComboBoxItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualComboBoxItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualComboBoxItemEvents