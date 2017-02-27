
// ReadXlsDlg.h : 头文件
//

#pragma once
//#import "msxml6.dll"
#include "afxwin.h"
#include "Excel.h"
#include <map>
#include <vector>
#include <string>
// #include <msxml6.h>
// #include <comutil.h>
#include <locale>
// 
// #pragma comment(lib, "comsuppwd.lib") 

using namespace std;
//using namespace MSXML2;

class CReadXlsDlgAutoProxy;


// CReadXlsDlg 对话框
class CReadXlsDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CReadXlsDlg);
	friend class CReadXlsDlgAutoProxy;

// 构造
public:
	CReadXlsDlg(CWnd* pParent = NULL);	// 标准构造函数
	virtual ~CReadXlsDlg();

// 对话框数据
	enum { IDD = IDD_READXLS_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	CReadXlsDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()

	char* __CString2Constchar(CString& strSrc);
	void __CreateXmlFile(CString& strPath);
	void __WriteItem(CStdioFile& obFile, size_t nID, CString strParentID);
public:
	afx_msg void OnBnClickedToxml();
	// 退出程序
	CButton m_Exit;
	CButton m_btnToXml;
	afx_msg void OnBnClickedExit();
	CButton m_btnOpenFile;
	CButton m_btnSavePath;
	afx_msg void OnBnClickedOpenfile();
	CEdit m_editInputPath;
	CEdit m_editOutputPath;
	afx_msg void OnBnClickedSavepath();

private:
	char* m_oldLocale;
	Excel m_obExcel;
	vector<CString> m_vecSheetName;
	map<CString, vector<vector<CString>>> m_mapSheetList;
};
