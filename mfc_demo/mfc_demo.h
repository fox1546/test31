#pragma once

#include "resource.h"

// CExcelDemoDlg 对话框
class CExcelDemoDlg : public CDialogEx
{
    DECLARE_DYNAMIC(CExcelDemoDlg)

public:
    CExcelDemoDlg(CWnd* pParent = nullptr);   // 标准构造函数
    virtual ~CExcelDemoDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
    enum { IDD = IDD_MFCDEMO_DIALOG };
#endif

protected:
    virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持
    virtual BOOL OnInitDialog();

    DECLARE_MESSAGE_MAP()
public:
    afx_msg void OnBnClickedSaveButton();
    CListCtrl m_listCtrl;
    afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
};
