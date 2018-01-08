//////////////////////////////////////////////////////////////////////
/// \class Proxy_IDriveComboBoxItemContainerEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IDriveComboBoxItemContainerEvents interface</em>
///
/// \if UNICODE
///   \sa DriveComboBoxItemContainer, CBLCtlsLibU::_IDriveComboBoxItemContainerEvents
/// \else
///   \sa DriveComboBoxItemContainer, CBLCtlsLibA::_IDriveComboBoxItemContainerEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IDriveComboBoxItemContainerEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IDriveComboBoxItemContainerEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IDriveComboBoxItemContainerEvents