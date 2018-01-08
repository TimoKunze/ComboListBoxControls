//////////////////////////////////////////////////////////////////////
/// \class APIWrapper
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Provides wrappers for API functions not available on all supported systems</em>
///
/// This class wraps calls to parts of the Win32 API that may be missing on the executing system.
//////////////////////////////////////////////////////////////////////


#pragma once

#include "helpers.h"


#ifndef DOXYGEN_SHOULD_SKIP_THIS
	typedef HRESULT WINAPI SHGetImageListFn(int iImageList, REFIID riid, __out LPVOID* ppvObj);
	typedef HRESULT WINAPI SHGetKnownFolderIDListFn(__in REFKNOWNFOLDERID rfid, __in DWORD dwFlags, __in_opt HANDLE hToken, __deref_out PIDLIST_ABSOLUTE* ppidl);
#endif

class APIWrapper
{
private:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	APIWrapper(void);

public:
	/// \brief <em>Checks whether the executing system supports the \c SHGetImageList function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHGetImageList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762185.aspx">SHGetImageList</a>
	static BOOL IsSupported_SHGetImageList(void);
	/// \brief <em>Checks whether the executing system supports the \c SHGetKnownFolderIDList function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHGetKnownFolderIDList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762187.aspx">SHGetKnownFolderIDList</a>
	static BOOL IsSupported_SHGetKnownFolderIDList(void);

	/// \brief <em>Calls the \c SHGetImageList function if it is available on the executing system</em>
	///
	/// \param[in] imageList The \c iImageList parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in,out] ppInterfaceImpl The \c ppvObj parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHGetImageList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762185.aspx">SHGetImageList</a>
	static HRESULT SHGetImageList(int imageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHGetKnownFolderIDList function if it is available on the executing system</em>
	///
	/// \param[in] folderID The \c rfid parameter of the wrapped function.
	/// \param[in] flags The \c dwFlags parameter of the wrapped function.
	/// \param[in] hToken The \c hToken parameter of the wrapped function.
	/// \param[in,out] ppIDL The \c ppidl parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHGetKnownFolderIDList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762187.aspx">SHGetKnownFolderIDList</a>
	static HRESULT SHGetKnownFolderIDList(__in REFKNOWNFOLDERID folderID, __in DWORD flags, __in_opt HANDLE hToken, __deref_out PIDLIST_ABSOLUTE* ppIDL, __deref_out_opt HRESULT* pReturnValue);

protected:
	/// \brief <em>Stores whether support for the \c SHGetImageList API function has been checked</em>
	static BOOL checkedSupport_SHGetImageList;
	/// \brief <em>Stores whether support for the \c SHGetKnownFolderIDList API function has been checked</em>
	static BOOL checkedSupport_SHGetKnownFolderIDList;
	/// \brief <em>Caches the pointer to the \c SHGetImageList API function</em>
	static SHGetImageListFn* pfnSHGetImageList;
	/// \brief <em>Caches the pointer to the \c SHGetKnownFolderIDList API function</em>
	static SHGetKnownFolderIDListFn* pfnSHGetKnownFolderIDList;
};