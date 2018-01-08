//////////////////////////////////////////////////////////////////////
/// \class IMouseHookHandler
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between \c MouseHookProc and the object that installed the mouse hook</em>
///
/// This interface allows \c MouseHookProc to delegate the hooked mouse message to the object that
/// installed the mouse hook.
///
/// \sa ImageComboBox::OnRButtonDown, DriveComboBox::OnRButtonDown, MouseHookProc
//////////////////////////////////////////////////////////////////////


#pragma once


class IMouseHookHandler
{
public:
	/// \brief <em>Processes a hooked mouse message</em>
	///
	/// This method is called by the callback function that we defined when we installed a mouse hook to trap
	/// \c WM_RBUTTONUP messages for the \c ImageComboBox and \c DriveComboBox controls.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam The identifier of the mouse message.
	/// \param[in] lParam Points to a \c MOUSEHOOKSTRUCT structure that contains more information.
	///
	/// \return The value returned by \c CallNextHookEx.
	///
	/// \sa ImageComboBox::HandleMessage, DriveComboBox::HandleMessage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644968.aspx">MOUSEHOOKSTRUCT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644988.aspx">MouseProc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644974.aspx">CallNextHookEx</a>
	virtual LRESULT HandleMessage(int code, WPARAM wParam, LPARAM lParam) = 0;
};     // IMouseHookHandler