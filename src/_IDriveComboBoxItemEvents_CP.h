//////////////////////////////////////////////////////////////////////
/// \class Proxy_IDriveComboBoxItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IDriveComboBoxItemEvents interface</em>
///
/// \if UNICODE
///   \sa DriveComboBoxItem, CBLCtlsLibU::_IDriveComboBoxItemEvents
/// \else
///   \sa DriveComboBoxItem, CBLCtlsLibA::_IDriveComboBoxItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IDriveComboBoxItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IDriveComboBoxItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IDriveComboBoxItemEvents