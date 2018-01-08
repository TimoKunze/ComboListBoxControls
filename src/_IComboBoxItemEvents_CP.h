//////////////////////////////////////////////////////////////////////
/// \class Proxy_IComboBoxItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IComboBoxItemEvents interface</em>
///
/// \if UNICODE
///   \sa ComboBoxItem, CBLCtlsLibU::_IComboBoxItemEvents
/// \else
///   \sa ComboBoxItem, CBLCtlsLibA::_IComboBoxItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IComboBoxItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IComboBoxItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IComboBoxItemEvents