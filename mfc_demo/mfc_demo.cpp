// mfc_demo.cpp : 定义应用程序的入口点。
//

#include "framework.h"
#include "mfc_demo.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// Excel COM 接口
#import "C:\Program Files\Microsoft Office\Office15\EXCEL.EXE" no_namespace rename("DialogBox", "ExcelDialogBox") rename("RGB", "ExcelRGB") rename("CopyFile", "ExcelCopyFile") rename("ReplaceText", "ExcelReplaceText")

// CExcelDemoDlg

IMPLEMENT_DYNAMIC(CExcelDemoDlg, CDialogEx)

CExcelDemoDlg::CExcelDemoDlg(CWnd* pParent /*=nullptr*/)
    : CDialogEx(IDD_MFCDEMO_DIALOG, pParent)
{
}

CExcelDemoDlg::~CExcelDemoDlg()
{
}

void CExcelDemoDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_LIST1, m_listCtrl);
}

BEGIN_MESSAGE_MAP(CExcelDemoDlg, CDialogEx)
    ON_BN_CLICKED(IDC_SAVE_BUTTON, &CExcelDemoDlg::OnBnClickedSaveButton)
    ON_WM_CREATE()
END_MESSAGE_MAP()

// CExcelDemoDlg 消息处理程序

int CExcelDemoDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
    if (CDialogEx::OnCreate(lpCreateStruct) == -1)
        return -1;

    // TODO:  在此添加您专用的创建代码
    return 0;
}

BOOL CExcelDemoDlg::OnInitDialog()
{
    CDialogEx::OnInitDialog();

    // 设置列表控件的扩展样式
    m_listCtrl.SetExtendedStyle(m_listCtrl.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);

    // 添加列
    m_listCtrl.InsertColumn(0, _T("列1"), LVCFMT_LEFT, 100);
    m_listCtrl.InsertColumn(1, _T("列2"), LVCFMT_LEFT, 100);
    m_listCtrl.InsertColumn(2, _T("列3"), LVCFMT_LEFT, 100);
    m_listCtrl.InsertColumn(3, _T("列4"), LVCFMT_LEFT, 100);
    m_listCtrl.InsertColumn(4, _T("列5"), LVCFMT_LEFT, 100);

    // 添加示例行
    for (int i = 0; i < 10; i++)
    {
        CString strRow;
        strRow.Format(_T("行%d"), i + 1);
        int nItem = m_listCtrl.InsertItem(i, strRow);
        for (int j = 1; j < 5; j++)
        {
            CString strCell;
            strCell.Format(_T("单元格%d-%d"), i + 1, j + 1);
            m_listCtrl.SetItemText(nItem, j, strCell);
        }
    }

    return TRUE;
}

void CExcelDemoDlg::OnBnClickedSaveButton()
{
    // 显示文件保存对话框
    CFileDialog dlg(FALSE, _T("xlsx"), _T("Excel文档"), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("Excel文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*||"), this);
    if (dlg.DoModal() == IDOK)
    {
        CString strFilePath = dlg.GetPathName();
        
        try
        {
            // 初始化COM
            CoInitialize(NULL);
            
            // 创建Excel应用程序
            _ApplicationPtr pExcelApp;
            if (pExcelApp.CreateInstance("Excel.Application") != S_OK)
            {
                AfxMessageBox(_T("无法创建Excel应用程序"));
                return;
            }
            
            // 显示Excel窗口
            pExcelApp->Visible = FALSE;
            
            // 创建工作簿
            _WorkbookPtr pWorkbook = pExcelApp->Workbooks->Add();
            _WorksheetPtr pWorksheet = pWorkbook->ActiveSheet;
            
            // 获取列表控件的列数和行数
            int nCols = m_listCtrl.GetHeaderCtrl()->GetItemCount();
            int nRows = m_listCtrl.GetItemCount();
            
            // 写入列标题
            for (int i = 0; i < nCols; i++)
            {
                LVCOLUMN lvColumn;
                TCHAR szText[256];
                lvColumn.mask = LVCF_TEXT;
                lvColumn.pszText = szText;
                lvColumn.cchTextMax = 256;
                m_listCtrl.GetColumn(i, &lvColumn);
                pWorksheet->Cells->Item[1][i + 1] = szText;
            }
            
            // 写入数据
            for (int i = 0; i < nRows; i++)
            {
                for (int j = 0; j < nCols; j++)
                {
                    CString strText = m_listCtrl.GetItemText(i, j);
                    pWorksheet->Cells->Item[i + 2][j + 1] = (LPCTSTR)strText;
                }
            }
            
            // 保存文件
            pWorkbook->SaveAs(_bstr_t((LPCTSTR)strFilePath), vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 0, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);
            
            // 关闭工作簿
            pWorkbook->Close(FALSE);
            
            // 退出Excel
            pExcelApp->Quit();
            
            // 释放COM对象
            pWorksheet = NULL;
            pWorkbook = NULL;
            pExcelApp = NULL;
            
            // 显示保存成功消息
            AfxMessageBox(_T("Excel文件保存成功！"));
        }
        catch (_com_error& e)
        {
            CString strError;
            strError.Format(_T("Excel操作错误: %s"), e.ErrorMessage());
            AfxMessageBox(strError);
        }
        catch (...)
        {
            AfxMessageBox(_T("未知错误"));
        }
        
        // 释放COM
        CoUninitialize();
    }
}

// 应用程序类
class CMyApp : public CWinApp
{
public:
    virtual BOOL InitInstance();
};

CMyApp theApp;

BOOL CMyApp::InitInstance()
{
    CWinApp::InitInstance();

    CExcelDemoDlg dlg;
    m_pMainWnd = &dlg;
    dlg.DoModal();

    return FALSE;
}
