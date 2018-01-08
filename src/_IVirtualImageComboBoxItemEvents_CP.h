//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualImageComboBoxItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualImageComboBoxItemEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualImageComboBoxItem, CBLCtlsLibU::_IVirtualImageComboBoxItemEvents
/// \else
///   \sa VirtualImageComboBoxItem, CBLCtlsLibA::_IVirtualImageComboBoxItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualImageComboBoxItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualImageComboBoxItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualImageComboBoxItemEvents