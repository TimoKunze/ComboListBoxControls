//////////////////////////////////////////////////////////////////////
/// \class Proxy_IImageComboBoxItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IImageComboBoxItemEvents interface</em>
///
/// \if UNICODE
///   \sa ImageComboBoxItem, CBLCtlsLibU::_IImageComboBoxItemEvents
/// \else
///   \sa ImageComboBoxItem, CBLCtlsLibA::_IImageComboBoxItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IImageComboBoxItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IImageComboBoxItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IImageComboBoxItemEvents