// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:FCCB83BF-E483-4317-9FF2-A460758238B5> version("1.5") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_CMBFONT, CMainDlg, &__uuidof(_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_LSTFONT, CMainDlg, &__uuidof(_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_CMBCOLOR, CMainDlg, &__uuidof(_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_LSTCOLOR, CMainDlg, &__uuidof(_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> cmbFontWnd;
	CContainedWindowT<CAxWindow> lstFontWnd;
	CContainedWindowT<CAxWindow> cmbColorWnd;
	CContainedWindowT<CAxWindow> lstColorWnd;

	HFONT hDefaultFont;

	typedef struct ENUMFONTPARAM
	{
		LOGFONT lf;
		BOOL fillList;
		CComPtr<IComboBoxItems> pComboItemsToAddTo;
		CComPtr<IListBoxItems> pListItemsToAddTo;
	} ENUMFONTPARAM, *LPENUMFONTPARAM;
	static int CALLBACK EnumFontFamExProc(const LPENUMLOGFONTEX lpElfe, const NEWTEXTMETRICEX* /*lpntme*/, DWORD /*FontType*/, LPARAM lParam);

	CMainDlg() :
	    cmbFontWnd(this, 1),
	    lstFontWnd(this, 2),
	    cmbColorWnd(this, 3),
	    lstColorWnd(this, 4)
	{
		hDefaultFont = NULL;
	}

	struct Controls
	{
		CComPtr<IComboBox> cmbFont;
		CComPtr<IListBox> lstFont;
		CComPtr<IComboBox> cmbColor;
		CComPtr<IListBox> lstColor;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		ALT_MSG_MAP(4)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_CMBFONT, __uuidof(_IComboBoxEvents), 11, FreeItemDataCmbFont)
		SINK_ENTRY_EX(IDC_CMBFONT, __uuidof(_IComboBoxEvents), 30, MeasureItemCmbFont)
		SINK_ENTRY_EX(IDC_CMBFONT, __uuidof(_IComboBoxEvents), 42, OwnerDrawItemCmbFont)
		SINK_ENTRY_EX(IDC_CMBFONT, __uuidof(_IComboBoxEvents), 45, RecreatedControlWindowCmbFont)

		SINK_ENTRY_EX(IDC_LSTFONT, __uuidof(_IListBoxEvents), 9, FreeItemDataLstFont)
		SINK_ENTRY_EX(IDC_LSTFONT, __uuidof(_IListBoxEvents), 15, ItemGetInfoTipTextLstFont)
		SINK_ENTRY_EX(IDC_LSTFONT, __uuidof(_IListBoxEvents), 23, MeasureItemLstFont)
		SINK_ENTRY_EX(IDC_LSTFONT, __uuidof(_IListBoxEvents), 43, OwnerDrawItemLstFont)
		SINK_ENTRY_EX(IDC_LSTFONT, __uuidof(_IListBoxEvents), 48, RecreatedControlWindowLstFont)

		SINK_ENTRY_EX(IDC_CMBCOLOR, __uuidof(_IComboBoxEvents), 42, OwnerDrawItemCmbColor)
		SINK_ENTRY_EX(IDC_CMBCOLOR, __uuidof(_IComboBoxEvents), 45, RecreatedControlWindowCmbColor)

		SINK_ENTRY_EX(IDC_LSTCOLOR, __uuidof(_IListBoxEvents), 43, OwnerDrawItemLstColor)
		SINK_ENTRY_EX(IDC_LSTCOLOR, __uuidof(_IListBoxEvents), 48, RecreatedControlWindowLstColor)
	END_SINK_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertComboItems_Font(void);
	void InsertListItems_Font(void);
	void InsertComboItems_Color(void);
	void InsertListItems_Color(void);

	void __stdcall FreeItemDataCmbFont(LPDISPATCH comboItem);
	void __stdcall MeasureItemCmbFont(LPDISPATCH comboItem, OLE_YSIZE_PIXELS* itemHeight);
	void __stdcall OwnerDrawItemCmbFont(LPDISPATCH comboItem, OwnerDrawActionConstants /*requiredAction*/, OwnerDrawItemStateConstants itemState, long hDC, RECTANGLE* drawingRectangle);
	void __stdcall RecreatedControlWindowCmbFont(long /*hWnd*/);

	void __stdcall FreeItemDataLstFont(LPDISPATCH listItem);
	void __stdcall ItemGetInfoTipTextLstFont(LPDISPATCH listItem, long /*maxInfoTipLength*/, BSTR* infoTipText, VARIANT_BOOL* /*abortToolTip*/);
	void __stdcall MeasureItemLstFont(LPDISPATCH listItem, OLE_YSIZE_PIXELS* itemHeight);
	void __stdcall OwnerDrawItemLstFont(LPDISPATCH listItem, OwnerDrawActionConstants /*requiredAction*/, OwnerDrawItemStateConstants itemState, long hDC, RECTANGLE* drawingRectangle);
	void __stdcall RecreatedControlWindowLstFont(long /*hWnd*/);

	void __stdcall OwnerDrawItemCmbColor(LPDISPATCH comboItem, OwnerDrawActionConstants requiredAction, OwnerDrawItemStateConstants itemState, long hDC, RECTANGLE* drawingRectangle);
	void __stdcall RecreatedControlWindowCmbColor(long /*hWnd*/);

	void __stdcall OwnerDrawItemLstColor(LPDISPATCH listItem, OwnerDrawActionConstants requiredAction, OwnerDrawItemStateConstants itemState, long hDC, RECTANGLE* drawingRectangle);
	void __stdcall RecreatedControlWindowLstColor(long /*hWnd*/);
};
