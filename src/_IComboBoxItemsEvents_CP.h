//////////////////////////////////////////////////////////////////////
/// \class Proxy_IComboBoxItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IComboBoxItemsEvents interface</em>
///
/// \if UNICODE
///   \sa ComboBoxItems, CBLCtlsLibU::_IComboBoxItemsEvents
/// \else
///   \sa ComboBoxItems, CBLCtlsLibA::_IComboBoxItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IComboBoxItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IComboBoxItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IComboBoxItemsEvents