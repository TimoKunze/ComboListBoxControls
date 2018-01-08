//////////////////////////////////////////////////////////////////////
/// \class Proxy_IImageComboBoxItemContainerEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IImageComboBoxItemContainerEvents interface</em>
///
/// \if UNICODE
///   \sa ImageComboBoxItemContainer, CBLCtlsLibU::_IImageComboBoxItemContainerEvents
/// \else
///   \sa ImageComboBoxItemContainer, CBLCtlsLibA::_IImageComboBoxItemContainerEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IImageComboBoxItemContainerEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IImageComboBoxItemContainerEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IImageComboBoxItemContainerEvents