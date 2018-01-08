//////////////////////////////////////////////////////////////////////
/// \class Proxy_IDriveComboBoxItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IDriveComboBoxItemsEvents interface</em>
///
/// \if UNICODE
///   \sa DriveComboBoxItems, CBLCtlsLibU::_IDriveComboBoxItemsEvents
/// \else
///   \sa DriveComboBoxItems, CBLCtlsLibA::_IDriveComboBoxItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IDriveComboBoxItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IDriveComboBoxItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IDriveComboBoxItemsEvents