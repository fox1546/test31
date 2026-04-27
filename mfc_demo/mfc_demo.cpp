// mfc_demo.cpp : 定义应用程序的入口点。
//

#include "framework.h"
#include "mfc_demo.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// Excel COM 接口
// 暂时注释掉Excel COM接口，使用CSV格式保存
// #import "C:\Program Files\Microsoft Office\Office15\EXCEL.EXE" no_namespace rename("DialogBox", "ExcelDialogBox") rename("RGB", "ExcelRGB") rename("CopyFile", "ExcelCopyFile") rename("ReplaceText", "ExcelReplaceText")

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
    CFileDialog dlg(FALSE, _T("csv"), _T("Excel文档"), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("Excel文件 (*.csv)|*.csv|所有文件 (*.*)|*.*||"), this);
    if (dlg.DoModal() == IDOK)
    {
        CString strFilePath = dlg.GetPathName();
        
        try
        {
            // 打开文件进行写入
            CStdioFile file;
            if (!file.Open(strFilePath, CFile::modeCreate | CFile::modeWrite | CFile::typeText))
            {
                AfxMessageBox(_T("无法打开文件进行写入"));
                return;
            }
            
            // 获取列表控件的列数和行数
            int nCols = m_listCtrl.GetHeaderCtrl()->GetItemCount();
            int nRows = m_listCtrl.GetItemCount();
            
            // 写入列标题
            CString strLine;
            for (int i = 0; i < nCols; i++)
            {
                LVCOLUMN lvColumn;
                TCHAR szText[256];
                lvColumn.mask = LVCF_TEXT;
                lvColumn.pszText = szText;
                lvColumn.cchTextMax = 256;
                m_listCtrl.GetColumn(i, &lvColumn);
                if (i > 0)
                    strLine += _T(",");
                strLine += _T("\"") + CString(szText) + _T("\"");
            }
            file.WriteString(strLine + _T("\n"));
            
            // 写入数据
            for (int i = 0; i < nRows; i++)
            {
                strLine.Empty();
                for (int j = 0; j < nCols; j++)
                {
                    CString strText = m_listCtrl.GetItemText(i, j);
                    if (j > 0)
                        strLine += _T(",");
                    strLine += _T("\"") + strText + _T("\"");
                }
                file.WriteString(strLine + _T("\n"));
            }
            
            // 关闭文件
            file.Close();
            
            // 显示保存成功消息
            AfxMessageBox(_T("Excel文件保存成功！"));
        }
        catch (CException* e)
        {
            TCHAR szError[1024];
            e->GetErrorMessage(szError, 1024);
            AfxMessageBox(CString(_T("文件操作错误: ")) + szError);
            e->Delete();
        }
        catch (...)
        {
            AfxMessageBox(_T("未知错误"));
        }
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
