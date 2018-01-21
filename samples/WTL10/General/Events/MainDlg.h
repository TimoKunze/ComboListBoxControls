// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#include <initguid.h>

#import <libid:FCCB83BF-E483-4317-9FF2-A460758238B5> version("1.5") raw_dispinterfaces
#import <libid:EE7B09EE-4DEB-47aa-8B0F-FA832AF08A0F> version("1.5") raw_dispinterfaces

DEFINE_GUID(LIBID_CBLCtlsLibU, 0xFCCB83BF, 0xE483, 0x4317, 0x9F, 0xF2, 0xA4, 0x60, 0x75, 0x82, 0x38, 0xB5);
DEFINE_GUID(LIBID_CBLCtlsLibA, 0xEE7B09EE, 0x4DEB, 0x47AA, 0x8B, 0x0F, 0xFA, 0x83, 0x2A, 0xF0, 0x8A, 0x0F);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_LSTU, CMainDlg, &__uuidof(CBLCtlsLibU::_IListBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_CMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_DRVCMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_IMGCMBU, CMainDlg, &__uuidof(CBLCtlsLibU::_IImageComboBoxEvents), &LIBID_CBLCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_LSTA, CMainDlg, &__uuidof(CBLCtlsLibA::_IListBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>,
    public IDispEventImpl<IDC_CMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>,
    public IDispEventImpl<IDC_DRVCMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>,
    public IDispEventImpl<IDC_IMGCMBA, CMainDlg, &__uuidof(CBLCtlsLibA::_IImageComboBoxEvents), &LIBID_CBLCtlsLibA, 1, 5>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> lstUWnd;
	CContainedWindowT<CAxWindow> cmbUWnd;
	CContainedWindowT<CAxWindow> drvcmbUWnd;
	CContainedWindowT<CAxWindow> imgcmbUWnd;
	CContainedWindowT<CAxWindow> lstAWnd;
	CContainedWindowT<CAxWindow> cmbAWnd;
	CContainedWindowT<CAxWindow> drvcmbAWnd;
	CContainedWindowT<CAxWindow> imgcmbAWnd;

	CMainDlg() :
	    lstUWnd(this, 1),
	    cmbUWnd(this, 2),
	    drvcmbUWnd(this, 3),
	    imgcmbUWnd(this, 4),
	    lstAWnd(this, 5),
	    cmbAWnd(this, 6),
	    drvcmbAWnd(this, 7),
	    imgcmbAWnd(this, 8)
	{
	}

	int focusedControl;

	struct Controls
	{
		CEdit logEdit;
		CButton aboutButton;
		CImageList imageList;
		CComPtr<CBLCtlsLibU::IListBox> lstU;
		CComPtr<CBLCtlsLibU::IComboBox> cmbU;
		CComPtr<CBLCtlsLibU::IDriveComboBox> drvcmbU;
		CComPtr<CBLCtlsLibU::IImageComboBox> imgcmbU;
		CComPtr<CBLCtlsLibA::IListBox> lstA;
		CComPtr<CBLCtlsLibA::IComboBox> cmbA;
		CComPtr<CBLCtlsLibA::IDriveComboBox> drvcmbA;
		CComPtr<CBLCtlsLibA::IImageComboBox> imgcmbA;
	} controls;

	BOOL IsComctl32Version600OrNewer(void);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		COMMAND_CODE_HANDLER(LBN_SETFOCUS, OnSetFocusLstU)

		ALT_MSG_MAP(2)
		COMMAND_CODE_HANDLER(CBN_SETFOCUS, OnSetFocusCmbU)

		ALT_MSG_MAP(3)
		COMMAND_CODE_HANDLER(CBN_SETFOCUS, OnSetFocusDrvCmbU)

		ALT_MSG_MAP(4)
		COMMAND_CODE_HANDLER(CBN_SETFOCUS, OnSetFocusImgCmbU)

		ALT_MSG_MAP(5)
		COMMAND_CODE_HANDLER(LBN_SETFOCUS, OnSetFocusLstA)

		ALT_MSG_MAP(6)
		COMMAND_CODE_HANDLER(CBN_SETFOCUS, OnSetFocusCmbA)

		ALT_MSG_MAP(7)
		COMMAND_CODE_HANDLER(CBN_SETFOCUS, OnSetFocusDrvCmbA)

		ALT_MSG_MAP(8)
		COMMAND_CODE_HANDLER(CBN_SETFOCUS, OnSetFocusImgCmbA)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 0, CaretChangedLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 1, AbortedDragLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 2, ClickLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 3, CompareItemsLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 4, ContextMenuLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 5, DblClickLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 6, DestroyedControlWindowLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 7, DragMouseMoveLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 8, DropLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 9, FreeItemDataLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 10, InsertedItemLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 11, InsertingItemLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 12, ItemBeginDragLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 13, ItemBeginRDragLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 14, ItemGetDisplayInfoLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 15, ItemGetInfoTipTextLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 16, ItemMouseEnterLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 17, ItemMouseLeaveLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 18, KeyDownLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 19, KeyPressLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 20, KeyUpLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 21, MClickLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 22, MDblClickLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 23, MeasureItemLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 24, MouseDownLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 25, MouseEnterLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 26, MouseHoverLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 27, MouseLeaveLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 28, MouseMoveLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 29, MouseUpLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 30, OLECompleteDragLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 31, OLEDragDropLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 32, OLEDragEnterLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 33, OLEDragEnterPotentialTargetLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 34, OLEDragLeaveLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 35, OLEDragLeavePotentialTargetLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 36, OLEDragMouseMoveLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 37, OLEGiveFeedbackLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 38, OLEQueryContinueDragLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 39, OLEReceivedNewDataLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 40, OLESetDataLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 41, OLEStartDragLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 42, OutOfMemoryLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 43, OwnerDrawItemLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 44, ProcessCharacterInputLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 45, ProcessKeyStrokeLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 46, RClickLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 47, RDblClickLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 48, RecreatedControlWindowLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 49, RemovedItemLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 50, RemovingItemLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 51, ResizedControlWindowLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 52, SelectionChangedLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 53, MouseWheelLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 54, XClickLstU)
		SINK_ENTRY_EX(IDC_LSTU, __uuidof(CBLCtlsLibU::_IListBoxEvents), 55, XDblClickLstU)

		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 0, SelectionChangedCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 1, BeforeDrawTextCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 2, ClickCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 3, CompareItemsCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 4, ContextMenuCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 5, CreatedEditControlWindowCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 6, CreatedListBoxControlWindowCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 7, DblClickCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 8, DestroyedControlWindowCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 9, DestroyedEditControlWindowCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 10, DestroyedListBoxControlWindowCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 11, FreeItemDataCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 12, InsertedItemCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 13, InsertingItemCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 14, ItemMouseEnterCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 15, ItemMouseLeaveCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 16, KeyDownCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 17, KeyPressCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 18, KeyUpCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 19, ListCloseUpCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 20, ListDropDownCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 21, ListMouseDownCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 22, ListMouseMoveCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 23, ListMouseUpCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 24, ListOLEDragDropCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 25, ListOLEDragEnterCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 26, ListOLEDragLeaveCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 27, ListOLEDragMouseMoveCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 28, MClickCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 29, MDblClickCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 30, MeasureItemCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 31, MouseDownCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 32, MouseEnterCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 33, MouseHoverCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 34, MouseLeaveCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 35, MouseMoveCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 36, MouseUpCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 37, OLEDragDropCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 38, OLEDragEnterCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 39, OLEDragLeaveCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 40, OLEDragMouseMoveCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 41, OutOfMemoryCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 42, OwnerDrawItemCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 43, RClickCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 44, RDblClickCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 45, RecreatedControlWindowCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 46, RemovedItemCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 47, RemovingItemCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 48, ResizedControlWindowCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 49, SelectionCanceledCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 50, SelectionChangingCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 51, TextChangedCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 52, TruncatedTextCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 53, WritingDirectionChangedCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 54, ListMouseWheelCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 55, MouseWheelCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 56, XClickCmbU)
		SINK_ENTRY_EX(IDC_CMBU, __uuidof(CBLCtlsLibU::_IComboBoxEvents), 57, XDblClickCmbU)

		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 0, SelectedDriveChangedDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 1, BeginSelectionChangeDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 2, ClickDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 3, ContextMenuDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 4, CreatedComboBoxControlWindowDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 5, CreatedListBoxControlWindowDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 6, DblClickDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 7, DestroyedComboBoxControlWindowDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 8, DestroyedControlWindowDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 9, DestroyedListBoxControlWindowDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 10, FreeItemDataDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 11, InsertedItemDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 12, InsertingItemDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 13, ItemBeginDragDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 14, ItemBeginRDragDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 15, ItemGetDisplayInfoDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 16, ItemMouseEnterDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 17, ItemMouseLeaveDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 18, KeyDownDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 19, KeyPressDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 20, KeyUpDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 21, ListClickDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 22, ListCloseUpDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 23, ListDropDownDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 24, ListMouseDownDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 25, ListMouseMoveDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 26, ListMouseUpDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 27, ListOLEDragDropDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 28, ListOLEDragEnterDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 29, ListOLEDragLeaveDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 30, ListOLEDragMouseMoveDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 31, MClickDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 32, MDblClickDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 33, MouseDownDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 34, MouseEnterDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 35, MouseHoverDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 36, MouseLeaveDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 37, MouseMoveDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 38, MouseUpDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 39, OLECompleteDragDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 40, OLEDragDropDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 41, OLEDragEnterDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 42, OLEDragEnterPotentialTargetDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 43, OLEDragLeaveDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 44, OLEDragLeavePotentialTargetDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 45, OLEDragMouseMoveDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 46, OLEGiveFeedbackDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 47, OLEQueryContinueDragDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 48, OLEReceivedNewDataDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 49, OLESetDataDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 50, OLEStartDragDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 51, OutOfMemoryDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 52, RClickDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 53, RDblClickDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 54, RecreatedControlWindowDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 55, RemovedItemDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 56, RemovingItemDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 57, ResizedControlWindowDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 58, SelectionCanceledDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 59, SelectionChangedDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 60, SelectionChangingDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 61, ListMouseWheelDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 62, MouseWheelDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 63, XClickDrvCmbU)
		SINK_ENTRY_EX(IDC_DRVCMBU, __uuidof(CBLCtlsLibU::_IDriveComboBoxEvents), 64, XDblClickDrvCmbU)

		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 0, SelectionChangedImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 1, BeforeDrawTextImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 2, BeginSelectionChangeImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 3, ClickImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 4, ContextMenuImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 5, CreatedComboBoxControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 6, CreatedEditControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 7, CreatedListBoxControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 8, DblClickImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 9, DestroyedComboBoxControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 10, DestroyedControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 11, DestroyedEditControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 12, DestroyedListBoxControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 13, FreeItemDataImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 14, InsertedItemImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 15, InsertingItemImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 16, ItemBeginDragImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 17, ItemBeginRDragImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 18, ItemGetDisplayInfoImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 19, ItemMouseEnterImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 20, ItemMouseLeaveImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 21, KeyDownImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 22, KeyPressImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 23, KeyUpImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 24, ListClickImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 25, ListCloseUpImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 26, ListDropDownImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 27, ListMouseDownImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 28, ListMouseMoveImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 29, ListMouseUpImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 30, ListOLEDragDropImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 31, ListOLEDragEnterImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 32, ListOLEDragLeaveImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 33, ListOLEDragMouseMoveImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 34, MClickImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 35, MDblClickImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 36, MouseDownImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 37, MouseEnterImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 38, MouseHoverImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 39, MouseLeaveImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 40, MouseMoveImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 41, MouseUpImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 42, OLECompleteDragImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 43, OLEDragDropImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 44, OLEDragEnterImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 45, OLEDragEnterPotentialTargetImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 46, OLEDragLeaveImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 47, OLEDragLeavePotentialTargetImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 48, OLEDragMouseMoveImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 49, OLEGiveFeedbackImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 50, OLEQueryContinueDragImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 51, OLEReceivedNewDataImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 52, OLESetDataImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 53, OLEStartDragImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 54, OutOfMemoryImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 55, RClickImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 56, RDblClickImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 57, RecreatedControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 58, RemovedItemImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 59, RemovingItemImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 60, ResizedControlWindowImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 61, SelectionCanceledImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 62, SelectionChangingImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 63, TextChangedImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 64, TruncatedTextImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 65, WritingDirectionChangedImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 66, ListMouseWheelImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 67, MouseWheelImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 68, XClickImgCmbU)
		SINK_ENTRY_EX(IDC_IMGCMBU, __uuidof(CBLCtlsLibU::_IImageComboBoxEvents), 69, XDblClickImgCmbU)

		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 0, CaretChangedLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 1, AbortedDragLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 2, ClickLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 3, CompareItemsLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 4, ContextMenuLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 5, DblClickLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 6, DestroyedControlWindowLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 7, DragMouseMoveLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 8, DropLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 9, FreeItemDataLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 10, InsertedItemLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 11, InsertingItemLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 12, ItemBeginDragLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 13, ItemBeginRDragLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 14, ItemGetDisplayInfoLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 15, ItemGetInfoTipTextLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 16, ItemMouseEnterLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 17, ItemMouseLeaveLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 18, KeyDownLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 19, KeyPressLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 20, KeyUpLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 21, MClickLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 22, MDblClickLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 23, MeasureItemLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 24, MouseDownLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 25, MouseEnterLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 26, MouseHoverLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 27, MouseLeaveLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 28, MouseMoveLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 29, MouseUpLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 30, OLECompleteDragLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 31, OLEDragDropLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 32, OLEDragEnterLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 33, OLEDragEnterPotentialTargetLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 34, OLEDragLeaveLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 35, OLEDragLeavePotentialTargetLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 36, OLEDragMouseMoveLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 37, OLEGiveFeedbackLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 38, OLEQueryContinueDragLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 39, OLEReceivedNewDataLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 40, OLESetDataLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 41, OLEStartDragLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 42, OutOfMemoryLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 43, OwnerDrawItemLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 44, ProcessCharacterInputLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 45, ProcessKeyStrokeLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 46, RClickLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 47, RDblClickLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 48, RecreatedControlWindowLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 49, RemovedItemLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 50, RemovingItemLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 51, ResizedControlWindowLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 52, SelectionChangedLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 53, MouseWheelLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 54, XClickLstA)
		SINK_ENTRY_EX(IDC_LSTA, __uuidof(CBLCtlsLibA::_IListBoxEvents), 55, XDblClickLstA)

		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 0, SelectionChangedCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 1, BeforeDrawTextCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 2, ClickCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 3, CompareItemsCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 4, ContextMenuCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 5, CreatedEditControlWindowCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 6, CreatedListBoxControlWindowCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 7, DblClickCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 8, DestroyedControlWindowCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 9, DestroyedEditControlWindowCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 10, DestroyedListBoxControlWindowCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 11, FreeItemDataCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 12, InsertedItemCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 13, InsertingItemCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 14, ItemMouseEnterCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 15, ItemMouseLeaveCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 16, KeyDownCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 17, KeyPressCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 18, KeyUpCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 19, ListCloseUpCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 20, ListDropDownCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 21, ListMouseDownCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 22, ListMouseMoveCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 23, ListMouseUpCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 24, ListOLEDragDropCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 25, ListOLEDragEnterCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 26, ListOLEDragLeaveCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 27, ListOLEDragMouseMoveCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 28, MClickCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 29, MDblClickCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 30, MeasureItemCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 31, MouseDownCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 32, MouseEnterCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 33, MouseHoverCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 34, MouseLeaveCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 35, MouseMoveCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 36, MouseUpCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 37, OLEDragDropCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 38, OLEDragEnterCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 39, OLEDragLeaveCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 40, OLEDragMouseMoveCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 41, OutOfMemoryCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 42, OwnerDrawItemCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 43, RClickCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 44, RDblClickCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 45, RecreatedControlWindowCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 46, RemovedItemCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 47, RemovingItemCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 48, ResizedControlWindowCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 49, SelectionCanceledCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 50, SelectionChangingCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 51, TextChangedCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 52, TruncatedTextCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 53, WritingDirectionChangedCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 54, ListMouseWheelCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 55, MouseWheelCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 56, XClickCmbA)
		SINK_ENTRY_EX(IDC_CMBA, __uuidof(CBLCtlsLibA::_IComboBoxEvents), 57, XDblClickCmbA)

		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 0, SelectedDriveChangedDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 1, BeginSelectionChangeDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 2, ClickDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 3, ContextMenuDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 4, CreatedComboBoxControlWindowDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 5, CreatedListBoxControlWindowDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 6, DblClickDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 7, DestroyedComboBoxControlWindowDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 8, DestroyedControlWindowDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 9, DestroyedListBoxControlWindowDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 10, FreeItemDataDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 11, InsertedItemDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 12, InsertingItemDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 13, ItemBeginDragDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 14, ItemBeginRDragDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 15, ItemGetDisplayInfoDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 16, ItemMouseEnterDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 17, ItemMouseLeaveDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 18, KeyDownDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 19, KeyPressDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 20, KeyUpDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 21, ListClickDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 22, ListCloseUpDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 23, ListDropDownDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 24, ListMouseDownDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 25, ListMouseMoveDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 26, ListMouseUpDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 27, ListOLEDragDropDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 28, ListOLEDragEnterDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 29, ListOLEDragLeaveDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 30, ListOLEDragMouseMoveDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 31, MClickDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 32, MDblClickDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 33, MouseDownDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 34, MouseEnterDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 35, MouseHoverDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 36, MouseLeaveDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 37, MouseMoveDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 38, MouseUpDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 39, OLECompleteDragDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 40, OLEDragDropDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 41, OLEDragEnterDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 42, OLEDragEnterPotentialTargetDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 43, OLEDragLeaveDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 44, OLEDragLeavePotentialTargetDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 45, OLEDragMouseMoveDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 46, OLEGiveFeedbackDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 47, OLEQueryContinueDragDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 48, OLEReceivedNewDataDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 49, OLESetDataDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 50, OLEStartDragDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 51, OutOfMemoryDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 52, RClickDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 53, RDblClickDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 54, RecreatedControlWindowDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 55, RemovedItemDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 56, RemovingItemDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 57, ResizedControlWindowDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 58, SelectionCanceledDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 59, SelectionChangedDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 60, SelectionChangingDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 61, ListMouseWheelDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 62, MouseWheelDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 63, XClickDrvCmbA)
		SINK_ENTRY_EX(IDC_DRVCMBA, __uuidof(CBLCtlsLibA::_IDriveComboBoxEvents), 64, XDblClickDrvCmbA)

		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 0, SelectionChangedImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 1, BeforeDrawTextImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 2, BeginSelectionChangeImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 3, ClickImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 4, ContextMenuImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 5, CreatedComboBoxControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 6, CreatedEditControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 7, CreatedListBoxControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 8, DblClickImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 9, DestroyedComboBoxControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 10, DestroyedControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 11, DestroyedEditControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 12, DestroyedListBoxControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 13, FreeItemDataImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 14, InsertedItemImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 15, InsertingItemImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 16, ItemBeginDragImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 17, ItemBeginRDragImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 18, ItemGetDisplayInfoImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 19, ItemMouseEnterImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 20, ItemMouseLeaveImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 21, KeyDownImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 22, KeyPressImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 23, KeyUpImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 24, ListClickImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 25, ListCloseUpImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 26, ListDropDownImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 27, ListMouseDownImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 28, ListMouseMoveImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 29, ListMouseUpImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 30, ListOLEDragDropImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 31, ListOLEDragEnterImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 32, ListOLEDragLeaveImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 33, ListOLEDragMouseMoveImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 34, MClickImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 35, MDblClickImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 36, MouseDownImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 37, MouseEnterImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 38, MouseHoverImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 39, MouseLeaveImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 40, MouseMoveImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 41, MouseUpImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 42, OLECompleteDragImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 43, OLEDragDropImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 44, OLEDragEnterImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 45, OLEDragEnterPotentialTargetImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 46, OLEDragLeaveImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 47, OLEDragLeavePotentialTargetImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 48, OLEDragMouseMoveImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 49, OLEGiveFeedbackImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 50, OLEQueryContinueDragImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 51, OLEReceivedNewDataImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 52, OLESetDataImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 53, OLEStartDragImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 54, OutOfMemoryImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 55, RClickImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 56, RDblClickImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 57, RecreatedControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 58, RemovedItemImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 59, RemovingItemImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 60, ResizedControlWindowImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 61, SelectionCanceledImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 62, SelectionChangingImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 63, TextChangedImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 64, TruncatedTextImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 65, WritingDirectionChangedImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 66, ListMouseWheelImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 67, MouseWheelImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 68, XClickImgCmbA)
		SINK_ENTRY_EX(IDC_IMGCMBA, __uuidof(CBLCtlsLibA::_IImageComboBoxEvents), 69, XDblClickImgCmbA)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusLstU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusCmbU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusDrvCmbU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusImgCmbU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusLstA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusCmbA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusDrvCmbA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusImgCmbA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);
	void InsertComboItemsA(void);
	void InsertImageComboItemsA(void);
	void InsertListItemsA(void);
	void InsertComboItemsU(void);
	void InsertImageComboItemsU(void);
	void InsertListItemsU(void);

	void __stdcall CaretChangedLstU(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem);
	void __stdcall AbortedDragLstU(void);
	void __stdcall ClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CompareItemsLstU(LPDISPATCH firstItem, LPDISPATCH secondItem, LONG locale, CBLCtlsLibU::CompareResultConstants* result);
	void __stdcall ContextMenuLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DblClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DestroyedControlWindowLstU(LONG hWnd);
	void __stdcall DragMouseMoveLstU(LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall DropLstU(LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall FreeItemDataLstU(LPDISPATCH listItem);
	void __stdcall InsertedItemLstU(LPDISPATCH listItem);
	void __stdcall InsertingItemLstU(LPDISPATCH listItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemBeginDragLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemBeginRDragLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemGetDisplayInfoLstU(LPDISPATCH listItem, CBLCtlsLibU::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* overlayIndex);
	void __stdcall ItemGetInfoTipTextLstU(LPDISPATCH listItem, LONG maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip);
	void __stdcall ItemMouseEnterLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemMouseLeaveLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall KeyDownLstU(SHORT* keyCode, SHORT shift);
	void __stdcall KeyPressLstU(SHORT* keyAscii);
	void __stdcall KeyUpLstU(SHORT* keyCode, SHORT shift);
	void __stdcall MClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MDblClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MeasureItemLstU(LPDISPATCH listItem, OLE_YSIZE_PIXELS* itemHeight);
	void __stdcall MouseDownLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseUpLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLECompleteDragLstU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropLstU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterLstU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetLstU(LONG hWndPotentialTarget);
	void __stdcall OLEDragLeaveLstU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetLstU(void);
	void __stdcall OLEDragMouseMoveLstU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackLstU(CBLCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragLstU(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataLstU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLESetDataLstU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLEStartDragLstU(LPDISPATCH data);
	void __stdcall OutOfMemoryLstU(void);
	void __stdcall OwnerDrawItemLstU(LPDISPATCH listItem, CBLCtlsLibU::OwnerDrawActionConstants requiredAction, CBLCtlsLibU::OwnerDrawItemStateConstants itemState, LONG hDC, CBLCtlsLibU::RECTANGLE* drawingRectangle);
	void __stdcall ProcessCharacterInputLstU(SHORT keyAscii, SHORT shift, LONG* itemToSelect);
	void __stdcall ProcessKeyStrokeLstU(SHORT keyCode, SHORT shift, LONG* itemToSelect);
	void __stdcall RClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RDblClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowLstU(LONG hWnd);
	void __stdcall RemovedItemLstU(LPDISPATCH listItem);
	void __stdcall RemovingItemLstU(LPDISPATCH listItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowLstU(void);
	void __stdcall SelectionChangedLstU(void);
	void __stdcall XClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall XDblClickLstU(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);

	void __stdcall SelectionChangedCmbU(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem);
	void __stdcall BeforeDrawTextCmbU(void);
	void __stdcall ClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CompareItemsCmbU(LPDISPATCH firstItem, LPDISPATCH secondItem, LONG locale, CBLCtlsLibU::CompareResultConstants* result);
	void __stdcall ContextMenuCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedEditControlWindowCmbU(LONG hWndEdit);
	void __stdcall CreatedListBoxControlWindowCmbU(LONG hWndListBox);
	void __stdcall DblClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DestroyedControlWindowCmbU(LONG hWnd);
	void __stdcall DestroyedEditControlWindowCmbU(LONG hWndEdit);
	void __stdcall DestroyedListBoxControlWindowCmbU(LONG hWndListBox);
	void __stdcall FreeItemDataCmbU(LPDISPATCH comboItem);
	void __stdcall InsertedItemCmbU(LPDISPATCH comboItem);
	void __stdcall InsertingItemCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemMouseEnterCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemMouseLeaveCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall KeyDownCmbU(SHORT* keyCode, SHORT shift);
	void __stdcall KeyPressCmbU(SHORT* keyAscii);
	void __stdcall KeyUpCmbU(SHORT* keyCode, SHORT shift);
	void __stdcall ListCloseUpCmbU(void);
	void __stdcall ListDropDownCmbU(void);
	void __stdcall ListMouseDownCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseMoveCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseUpCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseWheelCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragDropCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragEnterCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall ListOLEDragLeaveCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp);
	void __stdcall ListOLEDragMouseMoveCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall MClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MDblClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MeasureItemCmbU(LPDISPATCH comboItem, OLE_YSIZE_PIXELS* itemHeight);
	void __stdcall MouseDownCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseUpCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragDropCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEDragLeaveCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragMouseMoveCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OutOfMemoryCmbU(void);
	void __stdcall OwnerDrawItemCmbU(LPDISPATCH comboItem, CBLCtlsLibU::OwnerDrawActionConstants requiredAction, CBLCtlsLibU::OwnerDrawItemStateConstants itemState, LONG hDC, CBLCtlsLibU::RECTANGLE* drawingRectangle);
	void __stdcall RClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RDblClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowCmbU(LONG hWnd);
	void __stdcall RemovedItemCmbU(LPDISPATCH comboItem);
	void __stdcall RemovingItemCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowCmbU(void);
	void __stdcall SelectionCanceledCmbU(void);
	void __stdcall SelectionChangingCmbU(void);
	void __stdcall TextChangedCmbU(void);
	void __stdcall TruncatedTextCmbU(void);
	void __stdcall WritingDirectionChangedCmbU(CBLCtlsLibU::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall XDblClickCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);

	void __stdcall SelectedDriveChangedDrvCmbU(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem);
	void __stdcall BeginSelectionChangeDrvCmbU(void);
	void __stdcall ClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ContextMenuDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedComboBoxControlWindowDrvCmbU(LONG hWndComboBox);
	void __stdcall CreatedListBoxControlWindowDrvCmbU(LONG hWndListBox);
	void __stdcall DblClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DestroyedComboBoxControlWindowDrvCmbU(LONG hWndComboBox);
	void __stdcall DestroyedControlWindowDrvCmbU(LONG hWnd);
	void __stdcall DestroyedListBoxControlWindowDrvCmbU(LONG hWndListBox);
	void __stdcall FreeItemDataDrvCmbU(LPDISPATCH comboItem);
	void __stdcall InsertedItemDrvCmbU(LPDISPATCH comboItem);
	void __stdcall InsertingItemDrvCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemBeginDragDrvCmbU(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemBeginRDragDrvCmbU(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemGetDisplayInfoDrvCmbU(LPDISPATCH comboItem, CBLCtlsLibU::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* selectedIconIndex, LONG* overlayIndex, LONG* indent, LONG maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain);
	void __stdcall ItemMouseEnterDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemMouseLeaveDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall KeyDownDrvCmbU(SHORT* keyCode, SHORT shift);
	void __stdcall KeyPressDrvCmbU(SHORT* keyAscii);
	void __stdcall KeyUpDrvCmbU(SHORT* keyCode, SHORT shift);
	void __stdcall ListClickDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListCloseUpDrvCmbU(void);
	void __stdcall ListDropDownDrvCmbU(void);
	void __stdcall ListMouseDownDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseMoveDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseUpDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseWheelDrvCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragDropDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragEnterDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall ListOLEDragLeaveDrvCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp);
	void __stdcall ListOLEDragMouseMoveDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall MClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MDblClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseDownDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseUpDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLECompleteDragDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEDragEnterPotentialTargetDrvCmbU(LONG hWndPotentialTarget);
	void __stdcall OLEDragLeaveDrvCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetDrvCmbU(void);
	void __stdcall OLEDragMouseMoveDrvCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEGiveFeedbackDrvCmbU(CBLCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragDrvCmbU(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataDrvCmbU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLESetDataDrvCmbU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLEStartDragDrvCmbU(LPDISPATCH data);
	void __stdcall OutOfMemoryDrvCmbU(void);
	void __stdcall RClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RDblClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowDrvCmbU(LONG hWnd);
	void __stdcall RemovedItemDrvCmbU(LPDISPATCH comboItem);
	void __stdcall RemovingItemDrvCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowDrvCmbU(void);
	void __stdcall SelectionCanceledDrvCmbU(void);
	void __stdcall SelectionChangedDrvCmbU(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem);
	void __stdcall SelectionChangingDrvCmbU(LPDISPATCH newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, CBLCtlsLibU::SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* cancelChange);
	void __stdcall XClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall XDblClickDrvCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);

	void __stdcall SelectionChangedImgCmbU(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem);
	void __stdcall BeforeDrawTextImgCmbU(void);
	void __stdcall BeginSelectionChangeImgCmbU(void);
	void __stdcall ClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ContextMenuImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedComboBoxControlWindowImgCmbU(LONG hWndComboBox);
	void __stdcall CreatedEditControlWindowImgCmbU(LONG hWndEdit);
	void __stdcall CreatedListBoxControlWindowImgCmbU(LONG hWndListBox);
	void __stdcall DblClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DestroyedComboBoxControlWindowImgCmbU(LONG hWndComboBox);
	void __stdcall DestroyedControlWindowImgCmbU(LONG hWnd);
	void __stdcall DestroyedEditControlWindowImgCmbU(LONG hWndEdit);
	void __stdcall DestroyedListBoxControlWindowImgCmbU(LONG hWndListBox);
	void __stdcall FreeItemDataImgCmbU(LPDISPATCH comboItem);
	void __stdcall InsertedItemImgCmbU(LPDISPATCH comboItem);
	void __stdcall InsertingItemImgCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemBeginDragImgCmbU(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemBeginRDragImgCmbU(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemGetDisplayInfoImgCmbU(LPDISPATCH comboItem, CBLCtlsLibU::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* selectedIconIndex, LONG* overlayIndex, LONG* indent, LONG maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain);
	void __stdcall ItemMouseEnterImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ItemMouseLeaveImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall KeyDownImgCmbU(SHORT* keyCode, SHORT shift);
	void __stdcall KeyPressImgCmbU(SHORT* keyAscii);
	void __stdcall KeyUpImgCmbU(SHORT* keyCode, SHORT shift);
	void __stdcall ListClickImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListCloseUpImgCmbU(void);
	void __stdcall ListDropDownImgCmbU(void);
	void __stdcall ListMouseDownImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseMoveImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseUpImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListMouseWheelImgCmbU(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragDropImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragEnterImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall ListOLEDragLeaveImgCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp);
	void __stdcall ListOLEDragMouseMoveImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibU::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall MClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MDblClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseDownImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseUpImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLECompleteDragImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEDragEnterPotentialTargetImgCmbU(LONG hWndPotentialTarget);
	void __stdcall OLEDragLeaveImgCmbU(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetImgCmbU(void);
	void __stdcall OLEDragMouseMoveImgCmbU(LPDISPATCH data, CBLCtlsLibU::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEGiveFeedbackImgCmbU(CBLCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragImgCmbU(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataImgCmbU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLESetDataImgCmbU(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLEStartDragImgCmbU(LPDISPATCH data);
	void __stdcall OutOfMemoryImgCmbU(void);
	void __stdcall RClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RDblClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowImgCmbU(LONG hWnd);
	void __stdcall RemovedItemImgCmbU(LPDISPATCH comboItem);
	void __stdcall RemovingItemImgCmbU(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowImgCmbU(void);
	void __stdcall SelectionCanceledImgCmbU(void);
	void __stdcall SelectionChangingImgCmbU(LPDISPATCH newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, CBLCtlsLibU::SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* cancelChange);
	void __stdcall TextChangedImgCmbU(void);
	void __stdcall TruncatedTextImgCmbU(void);
	void __stdcall WritingDirectionChangedImgCmbU(CBLCtlsLibU::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall XDblClickImgCmbU(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibU::HitTestConstants hitTestDetails);

	void __stdcall CaretChangedLstA(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem);
	void __stdcall AbortedDragLstA(void);
	void __stdcall ClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CompareItemsLstA(LPDISPATCH firstItem, LPDISPATCH secondItem, LONG locale, CBLCtlsLibA::CompareResultConstants* result);
	void __stdcall ContextMenuLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DblClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DestroyedControlWindowLstA(LONG hWnd);
	void __stdcall DragMouseMoveLstA(LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall DropLstA(LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall FreeItemDataLstA(LPDISPATCH listItem);
	void __stdcall InsertedItemLstA(LPDISPATCH listItem);
	void __stdcall InsertingItemLstA(LPDISPATCH listItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemBeginDragLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemBeginRDragLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemGetDisplayInfoLstA(LPDISPATCH listItem, CBLCtlsLibA::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* overlayIndex);
	void __stdcall ItemGetInfoTipTextLstA(LPDISPATCH listItem, LONG maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip);
	void __stdcall ItemMouseEnterLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemMouseLeaveLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall KeyDownLstA(SHORT* keyCode, SHORT shift);
	void __stdcall KeyPressLstA(SHORT* keyAscii);
	void __stdcall KeyUpLstA(SHORT* keyCode, SHORT shift);
	void __stdcall MClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MDblClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MeasureItemLstA(LPDISPATCH listItem, OLE_YSIZE_PIXELS* itemHeight);
	void __stdcall MouseDownLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseUpLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLECompleteDragLstA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropLstA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterLstA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetLstA(LONG hWndPotentialTarget);
	void __stdcall OLEDragLeaveLstA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetLstA(void);
	void __stdcall OLEDragMouseMoveLstA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackLstA(CBLCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragLstA(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataLstA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLESetDataLstA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLEStartDragLstA(LPDISPATCH data);
	void __stdcall OutOfMemoryLstA(void);
	void __stdcall OwnerDrawItemLstA(LPDISPATCH listItem, CBLCtlsLibA::OwnerDrawActionConstants requiredAction, CBLCtlsLibA::OwnerDrawItemStateConstants itemState, LONG hDC, CBLCtlsLibA::RECTANGLE* drawingRectangle);
	void __stdcall ProcessCharacterInputLstA(SHORT keyAscii, SHORT shift, LONG* itemToSelect);
	void __stdcall ProcessKeyStrokeLstA(SHORT keyCode, SHORT shift, LONG* itemToSelect);
	void __stdcall RClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RDblClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowLstA(LONG hWnd);
	void __stdcall RemovedItemLstA(LPDISPATCH listItem);
	void __stdcall RemovingItemLstA(LPDISPATCH listItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowLstA(void);
	void __stdcall SelectionChangedLstA(void);
	void __stdcall XClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall XDblClickLstA(LPDISPATCH listItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);

	void __stdcall SelectionChangedCmbA(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem);
	void __stdcall BeforeDrawTextCmbA(void);
	void __stdcall ClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CompareItemsCmbA(LPDISPATCH firstItem, LPDISPATCH secondItem, LONG locale, CBLCtlsLibA::CompareResultConstants* result);
	void __stdcall ContextMenuCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedEditControlWindowCmbA(LONG hWndEdit);
	void __stdcall CreatedListBoxControlWindowCmbA(LONG hWndListBox);
	void __stdcall DblClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DestroyedControlWindowCmbA(LONG hWnd);
	void __stdcall DestroyedEditControlWindowCmbA(LONG hWndEdit);
	void __stdcall DestroyedListBoxControlWindowCmbA(LONG hWndListBox);
	void __stdcall FreeItemDataCmbA(LPDISPATCH comboItem);
	void __stdcall InsertedItemCmbA(LPDISPATCH comboItem);
	void __stdcall InsertingItemCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemMouseEnterCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemMouseLeaveCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall KeyDownCmbA(SHORT* keyCode, SHORT shift);
	void __stdcall KeyPressCmbA(SHORT* keyAscii);
	void __stdcall KeyUpCmbA(SHORT* keyCode, SHORT shift);
	void __stdcall ListCloseUpCmbA(void);
	void __stdcall ListDropDownCmbA(void);
	void __stdcall ListMouseDownCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseMoveCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseUpCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseWheelCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragDropCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragEnterCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall ListOLEDragLeaveCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp);
	void __stdcall ListOLEDragMouseMoveCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall MClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MDblClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MeasureItemCmbA(LPDISPATCH comboItem, OLE_YSIZE_PIXELS* itemHeight);
	void __stdcall MouseDownCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseUpCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragDropCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEDragLeaveCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragMouseMoveCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OutOfMemoryCmbA(void);
	void __stdcall OwnerDrawItemCmbA(LPDISPATCH comboItem, CBLCtlsLibA::OwnerDrawActionConstants requiredAction, CBLCtlsLibA::OwnerDrawItemStateConstants itemState, LONG hDC, CBLCtlsLibA::RECTANGLE* drawingRectangle);
	void __stdcall RClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RDblClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowCmbA(LONG hWnd);
	void __stdcall RemovedItemCmbA(LPDISPATCH comboItem);
	void __stdcall RemovingItemCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowCmbA(void);
	void __stdcall SelectionCanceledCmbA(void);
	void __stdcall SelectionChangingCmbA(void);
	void __stdcall TextChangedCmbA(void);
	void __stdcall TruncatedTextCmbA(void);
	void __stdcall WritingDirectionChangedCmbA(CBLCtlsLibA::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall XDblClickCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);

	void __stdcall SelectedDriveChangedDrvCmbA(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem);
	void __stdcall BeginSelectionChangeDrvCmbA(void);
	void __stdcall ClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ContextMenuDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedComboBoxControlWindowDrvCmbA(LONG hWndComboBox);
	void __stdcall CreatedListBoxControlWindowDrvCmbA(LONG hWndListBox);
	void __stdcall DblClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DestroyedComboBoxControlWindowDrvCmbA(LONG hWndComboBox);
	void __stdcall DestroyedControlWindowDrvCmbA(LONG hWnd);
	void __stdcall DestroyedListBoxControlWindowDrvCmbA(LONG hWndListBox);
	void __stdcall FreeItemDataDrvCmbA(LPDISPATCH comboItem);
	void __stdcall InsertedItemDrvCmbA(LPDISPATCH comboItem);
	void __stdcall InsertingItemDrvCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemBeginDragDrvCmbA(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemBeginRDragDrvCmbA(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemGetDisplayInfoDrvCmbA(LPDISPATCH comboItem, CBLCtlsLibA::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* selectedIconIndex, LONG* overlayIndex, LONG* indent, LONG maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain);
	void __stdcall ItemMouseEnterDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemMouseLeaveDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall KeyDownDrvCmbA(SHORT* keyCode, SHORT shift);
	void __stdcall KeyPressDrvCmbA(SHORT* keyAscii);
	void __stdcall KeyUpDrvCmbA(SHORT* keyCode, SHORT shift);
	void __stdcall ListClickDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListCloseUpDrvCmbA(void);
	void __stdcall ListDropDownDrvCmbA(void);
	void __stdcall ListMouseDownDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseMoveDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseUpDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseWheelDrvCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragDropDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragEnterDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall ListOLEDragLeaveDrvCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp);
	void __stdcall ListOLEDragMouseMoveDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall MClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MDblClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseDownDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseUpDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLECompleteDragDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEDragEnterPotentialTargetDrvCmbA(LONG hWndPotentialTarget);
	void __stdcall OLEDragLeaveDrvCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetDrvCmbA(void);
	void __stdcall OLEDragMouseMoveDrvCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEGiveFeedbackDrvCmbA(CBLCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragDrvCmbA(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataDrvCmbA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLESetDataDrvCmbA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLEStartDragDrvCmbA(LPDISPATCH data);
	void __stdcall OutOfMemoryDrvCmbA(void);
	void __stdcall RClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RDblClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowDrvCmbA(LONG hWnd);
	void __stdcall RemovedItemDrvCmbA(LPDISPATCH comboItem);
	void __stdcall RemovingItemDrvCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowDrvCmbA(void);
	void __stdcall SelectionCanceledDrvCmbA(void);
	void __stdcall SelectionChangedDrvCmbA(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem);
	void __stdcall SelectionChangingDrvCmbA(LPDISPATCH newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, CBLCtlsLibA::SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* cancelChange);
	void __stdcall XClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall XDblClickDrvCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);

	void __stdcall SelectionChangedImgCmbA(LPDISPATCH previousSelectedItem, LPDISPATCH newSelectedItem);
	void __stdcall BeforeDrawTextImgCmbA(void);
	void __stdcall BeginSelectionChangeImgCmbA(void);
	void __stdcall ClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ContextMenuImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedComboBoxControlWindowImgCmbA(LONG hWndComboBox);
	void __stdcall CreatedEditControlWindowImgCmbA(LONG hWndEdit);
	void __stdcall CreatedListBoxControlWindowImgCmbA(LONG hWndListBox);
	void __stdcall DblClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DestroyedComboBoxControlWindowImgCmbA(LONG hWndComboBox);
	void __stdcall DestroyedControlWindowImgCmbA(LONG hWnd);
	void __stdcall DestroyedEditControlWindowImgCmbA(LONG hWndEdit);
	void __stdcall DestroyedListBoxControlWindowImgCmbA(LONG hWndListBox);
	void __stdcall FreeItemDataImgCmbA(LPDISPATCH comboItem);
	void __stdcall InsertedItemImgCmbA(LPDISPATCH comboItem);
	void __stdcall InsertingItemImgCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall ItemBeginDragImgCmbA(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemBeginRDragImgCmbA(LPDISPATCH comboItem, BSTR selectionFieldText, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemGetDisplayInfoImgCmbA(LPDISPATCH comboItem, CBLCtlsLibA::RequestedInfoConstants requestedInfo, LONG* iconIndex, LONG* selectedIconIndex, LONG* overlayIndex, LONG* indent, LONG maxItemTextLength, BSTR* itemText, VARIANT_BOOL* dontAskAgain);
	void __stdcall ItemMouseEnterImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ItemMouseLeaveImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall KeyDownImgCmbA(SHORT* keyCode, SHORT shift);
	void __stdcall KeyPressImgCmbA(SHORT* keyAscii);
	void __stdcall KeyUpImgCmbA(SHORT* keyCode, SHORT shift);
	void __stdcall ListClickImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListCloseUpImgCmbA(void);
	void __stdcall ListDropDownImgCmbA(void);
	void __stdcall ListMouseDownImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseMoveImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseUpImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListMouseWheelImgCmbA(LPDISPATCH comboItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragDropImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ListOLEDragEnterImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall ListOLEDragLeaveImgCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp);
	void __stdcall ListOLEDragMouseMoveImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG yToItemTop, CBLCtlsLibA::HitTestConstants hitTestDetails, LONG* autoHScrollVelocity, LONG* autoVScrollVelocity);
	void __stdcall MClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MDblClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseDownImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseUpImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::ScrollAxisConstants scrollAxis, SHORT wheelDelta, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLECompleteDragImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEDragEnterPotentialTargetImgCmbA(LONG hWndPotentialTarget);
	void __stdcall OLEDragLeaveImgCmbA(LPDISPATCH data, LPDISPATCH dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetImgCmbA(void);
	void __stdcall OLEDragMouseMoveImgCmbA(LPDISPATCH data, CBLCtlsLibA::OLEDropEffectConstants* effect, LPDISPATCH* dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEGiveFeedbackImgCmbA(CBLCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragImgCmbA(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, CBLCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataImgCmbA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLESetDataImgCmbA(LPDISPATCH data, LONG formatID, LONG index, LONG dataOrViewAspect);
	void __stdcall OLEStartDragImgCmbA(LPDISPATCH data);
	void __stdcall OutOfMemoryImgCmbA(void);
	void __stdcall RClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RDblClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowImgCmbA(LONG hWnd);
	void __stdcall RemovedItemImgCmbA(LPDISPATCH comboItem);
	void __stdcall RemovingItemImgCmbA(LPDISPATCH comboItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall ResizedControlWindowImgCmbA(void);
	void __stdcall SelectionCanceledImgCmbA(void);
	void __stdcall SelectionChangingImgCmbA(LPDISPATCH newSelectedItem, BSTR selectionFieldText, VARIANT_BOOL selectionFieldHasBeenEdited, CBLCtlsLibA::SelectionChangeReasonConstants selectionChangeReason, VARIANT_BOOL* cancelChange);
	void __stdcall TextChangedImgCmbA(void);
	void __stdcall TruncatedTextImgCmbA(void);
	void __stdcall WritingDirectionChangedImgCmbA(CBLCtlsLibA::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall XDblClickImgCmbA(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CBLCtlsLibA::HitTestConstants hitTestDetails);
};
