//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualDriveComboBoxItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualDriveComboBoxItemEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualDriveComboBoxItem, CBLCtlsLibU::_IVirtualDriveComboBoxItemEvents
/// \else
///   \sa VirtualDriveComboBoxItem, CBLCtlsLibA::_IVirtualDriveComboBoxItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualDriveComboBoxItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualDriveComboBoxItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualDriveComboBoxItemEvents