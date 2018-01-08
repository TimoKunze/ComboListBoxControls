//////////////////////////////////////////////////////////////////////
/// \class IItemContainer
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between a \c ListBoxItemContainer object and its creator object</em>
///
/// This interface allows a \c ListBox object to inform a \c ListBoxItemContainer object about item
/// deletions.
///
/// \sa ListBox::RegisterItemContainer, ListBoxItemContainer
//////////////////////////////////////////////////////////////////////


#pragma once


class IItemContainer
{
public:
	/// \brief <em>An item was removed and the item container should check its content</em>
	///
	/// \param[in] itemID The unique ID of the removed item. If 0, all items were removed.
	///
	/// \sa ListBox::RemoveItemFromItemContainers
	virtual void RemovedItem(LONG itemID) = 0;
	/// \brief <em>An item's unique ID has been changed and the item container should check its content</em>
	///
	/// \param[in] oldItemID The old unique ID of the removed item.
	/// \param[in] newItemID The new unique ID of the removed item.
	///
	/// \sa ListBox::UpdateItemIDInItemContainers
	virtual void ReplacedItemID(LONG oldItemID, LONG newItemID) = 0;
	/// \brief <em>Retrieves the container's ID used to identify it</em>
	///
	/// \return The container's ID.
	///
	/// \sa ListBox::DeregisterItemContainer, containerID
	virtual DWORD GetID(void) = 0;

	/// \brief <em>Holds the container's ID</em>
	DWORD containerID;
};     // IItemContainer