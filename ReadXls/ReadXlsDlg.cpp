
// ReadXlsDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ReadXls.h"
#include "ReadXlsDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CReadXlsDlg 对话框


IMPLEMENT_DYNAMIC(CReadXlsDlg, CDialogEx);

CReadXlsDlg::CReadXlsDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CReadXlsDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CReadXlsDlg::~CReadXlsDlg()
{
	// 如果该对话框有自动化代理，则
	//  将此代理指向该对话框的后向指针设置为 NULL，以便
	//  此代理知道该对话框已被删除。
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;

}

void CReadXlsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_Exit, m_Exit);
	DDX_Control(pDX, IDC_ToXml, m_btnToXml);
	DDX_Control(pDX, IDC_OpenFile, m_btnOpenFile);
	DDX_Control(pDX, IDC_SavePath, m_btnSavePath);
	DDX_Control(pDX, IDC_InputPath, m_editInputPath);
	DDX_Control(pDX, IDC_OutputPath, m_editOutputPath);
}

BEGIN_MESSAGE_MAP(CReadXlsDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_ToXml, &CReadXlsDlg::OnBnClickedToxml)
//	ON_EN_CHANGE(IDC_EDIT1, &CReadXlsDlg::OnEnChangeEdit1)
//ON_EN_CHANGE(IDC_EDIT1, &CReadXlsDlg::OnEnChangeEdit1)
ON_BN_CLICKED(IDC_Exit, &CReadXlsDlg::OnBnClickedExit)
ON_BN_CLICKED(IDC_OpenFile, &CReadXlsDlg::OnBnClickedOpenfile)
ON_BN_CLICKED(IDC_SavePath, &CReadXlsDlg::OnBnClickedSavepath)
ON_BN_CLICKED(IDC_ToXml2, &CReadXlsDlg::OnBnClickedToxml2)
END_MESSAGE_MAP()


// CReadXlsDlg 消息处理程序

BOOL CReadXlsDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	ShowWindow(SW_SHOW);


	// TODO:  在此添加额外的初始化代码
	if (!m_obExcel.initExcel())
	{
		return false;
	}
	m_oldLocale = _strdup(setlocale(LC_CTYPE, NULL));
	setlocale(LC_CTYPE, "chs");//设定
	m_strInputName = "";

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CReadXlsDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CReadXlsDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CReadXlsDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 当用户关闭 UI 时，如果控制器仍保持着它的某个
//  对象，则自动化服务器不应退出。  这些
//  消息处理程序确保如下情形: 如果代理仍在使用，
//  则将隐藏 UI；但是在关闭对话框时，
//  对话框仍然会保留在那里。

void CReadXlsDlg::OnClose()
{
	if (CanExit())
	{
		m_obExcel.release();
		//还原区域设定
		setlocale(LC_CTYPE, m_oldLocale);
		free(m_oldLocale);
		CDialogEx::OnClose();
	}
}

void CReadXlsDlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CReadXlsDlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CReadXlsDlg::CanExit()
{
	// 如果代理对象仍保留在那里，则自动化
	//  控制器仍会保持此应用程序。
	//  使对话框保留在那里，但将其 UI 隐藏起来。
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}



void CReadXlsDlg::OnBnClickedToxml()
{
	// TODO:  在此添加控件通知处理程序代码
	//使用Excel类
	CString strInputFile;
	CString strOutputPath;
	m_editInputPath.GetWindowTextW(strInputFile);
	m_editOutputPath.GetWindowTextW(strOutputPath);
	
	if (0 == strInputFile.GetLength())
	{
		MessageBox(_T("请先选择需要读取的文件"), _T("错误"), MB_OK);
		return;
	}
	if (0 == strOutputPath.GetLength())
	{
		MessageBox(_T("请先选择输出路径"), _T("错误"), MB_OK);
		return;
	}
	
	if (!m_obExcel.open(__CString2Constchar(strInputFile)))
	{
		MessageBox(_T("无法打开该文件"), _T("错误"), MB_OK);
		return;
	}
	int nSheetCount = m_obExcel.getSheetCount();

	//将整个excel读入m_mapSheetList中
	for (int s = 1; s <= nSheetCount; ++s)
	{
		CString strSheetName = m_obExcel.getSheetName(s);//获取sheet名
		m_vecSheetName.push_back(strSheetName);
		bool bLoad = m_obExcel.loadSheet(strSheetName);//装载sheet  
		int nRow = m_obExcel.getRowCount();//获取sheet中行数
		int nCol = m_obExcel.getColumnCount();//获取sheet中列数  


		CString cell;
		vector<vector<CString>> vecSheet;
		for (int i = 1; i <= nRow; ++i)
		{
			vector<CString> vecRow;
			for (int j = 1; j <= nCol; ++j)
			{
				cell = m_obExcel.getCellString(i, j);
				vecRow.push_back(cell);
			}
			vecSheet.push_back(vecRow);
		}
		m_mapSheetList[strSheetName] = vecSheet;
	}
	
	//将m_mapSheetList内容写入xml文件

	//第一种模式
	__CreateXmlFile(strOutputPath);

// 	//第二种模式
// 	__CreateXmlFile2(strOutputPath);
	
}

void CReadXlsDlg::OnBnClickedExit()
{
	// TODO:  在此添加控件通知处理程序代码
	SendMessage(WM_CLOSE, 0, 0);
}


void CReadXlsDlg::OnBnClickedOpenfile()
{
	// TODO:  在此添加控件通知处理程序代码
	// 设置过滤器   
	TCHAR szFilter[] = _T("Microsoft Excel 2007 工作表(*.xlsx)|*.xlsx|所有文件(*.*)|*.*||");
	// 构造打开文件对话框   
	CFileDialog fileDlg(TRUE, /*_T("xls")*/NULL, NULL, 0, szFilter, this);
	CString strFilePath;

	// 显示打开文件对话框   
	if (IDOK == fileDlg.DoModal())
	{
		// 如果点击了文件对话框上的“打开”按钮，则将选择的文件路径显示到编辑框里   
		strFilePath = fileDlg.GetPathName();
		SetDlgItemText(IDC_InputPath, strFilePath);

		m_strInputName = __GetFileName(strFilePath);
	}
}


void CReadXlsDlg::OnBnClickedSavepath()
{
	// TODO:  在此添加控件通知处理程序代码
	// 设置过滤器   
	TCHAR szFilter[] = _T("xml文件(*.xml)|*.xml|所有文件(*.*)|*.*||");
	//默认输出文件名
	CString strDefName;
	if (0 == m_strInputName.GetLength())
	{
		strDefName = "my";
	}
	else
	{
		strDefName = m_strInputName;
	}

	// 构造保存文件对话框   
	CFileDialog fileDlg(FALSE, _T("xml"), strDefName, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, this);
	CString strFilePath;

	// 显示保存文件对话框   
	if (IDOK == fileDlg.DoModal())
	{
		// 如果点击了文件对话框上的“保存”按钮，则将选择的文件路径显示到编辑框里   
		strFilePath = fileDlg.GetPathName();
		SetDlgItemText(IDC_OutputPath, strFilePath);
	}
}

char* CReadXlsDlg::__CString2Constchar(CString& strSrc)
{
	char* pszRet;
	const size_t strsize = (strSrc.GetLength() + 1) * 2;
	pszRet = new char[strsize];
	size_t sz = 0;
	wcstombs_s(&sz, pszRet, strsize, strSrc, _TRUNCATE);
	int n = atoi((const char*)pszRet);
	return pszRet;
}

void CReadXlsDlg::__CreateXmlFile(CString& strPath)
{
	//http://www.jizhuomi.com/software/497.html
	CStdioFile obFile;
	if (obFile.Open(strPath, CStdioFile::modeCreate | CStdioFile::modeWrite))
	{
		CString strHead = _T("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
		obFile.WriteString(strHead);

		CString strRoot = _T("<") + m_vecSheetName[0] + _T("List>\n");
		CString strFRoot = _T("</") + m_vecSheetName[0] + _T("List>\n");
		obFile.WriteString(strRoot);
		vector<vector<CString>> vecMark = m_mapSheetList[m_vecSheetName[0]];
		//size_t nMarkLen = vecMark.size();
		for (size_t i = 2; i < vecMark.size(); ++i)
		{
			CString strMark = _T("\t<") + m_vecSheetName[0];
			for (size_t j = 0; j < vecMark[i].size(); ++j)
			{
				strMark += _T(" ") + vecMark[0][j] + _T("=\"") + vecMark[i][j] + _T("\"");
			}
			strMark += _T(">\n");
			obFile.WriteString(strMark);
			for (size_t childID = 1; childID < m_vecSheetName.size(); ++childID)
			{
				__WriteItem(obFile, childID, vecMark[i][0]);
			}
			obFile.WriteString(_T("\t</") + m_vecSheetName[0] + _T(">\n"));
		}

		obFile.WriteString(strFRoot);

		obFile.Close();
		MessageBox(_T("生成成功"), _T("成功"), MB_OK);
	}
	else
	{
		MessageBox(_T("生成失败"), _T("失败"), MB_OK);
	}
}

void CReadXlsDlg::__WriteItem(CStdioFile& obFile, size_t nID, CString strParentID)
{
	CString strChildHead = _T("\t\t");
	strChildHead += _T("<") + m_vecSheetName[nID] + _T("List>\n");
	obFile.WriteString(strChildHead);
	vector<vector<CString>> vecChild = m_mapSheetList[m_vecSheetName[nID]];
	for (size_t i = 2; i < vecChild.size(); ++i)
	{
		if (strParentID != vecChild[i][0])
		{
			continue;
		}
		CString strChild = _T("\t\t\t<") + m_vecSheetName[nID];
		for (size_t j = 0; j < vecChild[i].size(); ++j)
		{
			strChild += _T(" ") + vecChild[0][j] + _T("=\"") + vecChild[i][j] + _T("\"");
		}
		strChild += _T("/>\n");
		obFile.WriteString(strChild);
	}

	obFile.WriteString(_T("\t\t</") + m_vecSheetName[nID] + _T("List>\n"));
}

void CReadXlsDlg::__CreateXmlFile2(CString& strPath)
{
	CString strFirstTag = __GetFileName(strPath);
	CStdioFile obFile;
	if (obFile.Open(strPath, CStdioFile::modeCreate | CStdioFile::modeWrite))
	{
		CString strHead = _T("<? xml version=\"1.0\" encoding=\"UTF-8\" ?>\n");
		obFile.WriteString(strHead);
		CString strRoot = _T("<") + strFirstTag + _T(">\n");
		CString strRootEnd = _T("</") + strFirstTag + _T(">\n");
		obFile.WriteString(strRoot);

		for (size_t si = 0; si < m_vecSheetName.size(); ++si)
		{
			vector<vector<CString>> vecContentList = m_mapSheetList[m_vecSheetName[si]];
			CString strSheetRoot = _T("\t<") + m_vecSheetName[si] + _T(">\n");
			CString strSheetRootEnd = _T("\t</") + m_vecSheetName[si] + _T(">\n");
			if (vecContentList.size() > 3)
			{
				obFile.WriteString(strSheetRoot);
			}
			for (size_t i = 2; i < vecContentList.size(); ++i)
			{
				CString strItem;
				if (vecContentList.size() > 3)
				{
					strItem = _T("\t\t<") + m_vecSheetName[si];
				}
				else
				{	
					strItem = _T("\t<") + m_vecSheetName[si];
				}
				for (size_t j = 0; j < vecContentList[i].size(); ++j)
				{
					strItem += _T(" ") + vecContentList[0][j] + _T("=\"") + vecContentList[i][j] + _T("\"");
				}
				strItem += _T("/>\n");
				obFile.WriteString(strItem);
			}
			if (vecContentList.size() > 3)
			{
				obFile.WriteString(strSheetRootEnd);
			}
			//obFile.WriteString(_T("\n"));
		}
		obFile.WriteString(strRootEnd);

		obFile.Close();
		MessageBox(_T("生成成功"), _T("成功"), MB_OK);
	}
	else
	{
		MessageBox(_T("生成失败"), _T("失败"), MB_OK);
	}
}

CString CReadXlsDlg::__GetFileName(CString strPath)
{
	vector<CString> vecItem;
	int nCurPos = 0;
	while (1)
	{
		CString resToken = strPath.Tokenize(_T("\\"), nCurPos);
		if (resToken.IsEmpty()) break;
		vecItem.push_back(resToken);
	}
	CString strFileName = vecItem[vecItem.size() - 1];
	nCurPos = 0;
	strFileName = strFileName.Tokenize(_T("."), nCurPos);
	return strFileName;
}

void CReadXlsDlg::OnBnClickedToxml2()
{
	// TODO:  在此添加控件通知处理程序代码
	CString strInputFile;
	CString strOutputPath;
	m_editInputPath.GetWindowTextW(strInputFile);
	m_editOutputPath.GetWindowTextW(strOutputPath);

	if (0 == strInputFile.GetLength())
	{
		MessageBox(_T("请先选择需要读取的文件"), _T("错误"), MB_OK);
		return;
	}
	if (0 == strOutputPath.GetLength())
	{
		MessageBox(_T("请先选择输出路径"), _T("错误"), MB_OK);
		return;
	}

	if (!m_obExcel.open(__CString2Constchar(strInputFile)))
	{
		MessageBox(_T("无法打开该文件"), _T("错误"), MB_OK);
		return;
	}
	int nSheetCount = m_obExcel.getSheetCount();

	//将整个excel读入m_mapSheetList中
	for (int s = 1; s <= nSheetCount; ++s)
	{
		CString strSheetName = m_obExcel.getSheetName(s);//获取sheet名
		m_vecSheetName.push_back(strSheetName);
		bool bLoad = m_obExcel.loadSheet(strSheetName);//装载sheet  
		int nRow = m_obExcel.getRowCount();//获取sheet中行数
		int nCol = m_obExcel.getColumnCount();//获取sheet中列数  


		CString cell;
		vector<vector<CString>> vecSheet;
		for (int i = 1; i <= nRow; ++i)
		{
			vector<CString> vecRow;
			for (int j = 1; j <= nCol; ++j)
			{
				cell = m_obExcel.getCellString(i, j);
				vecRow.push_back(cell);
			}
			vecSheet.push_back(vecRow);
		}
		m_mapSheetList[strSheetName] = vecSheet;
	}

	//将m_mapSheetList内容写入xml文件

// 	//第一种模式
// 	__CreateXmlFile(strOutputPath);

	//第二种模式
	__CreateXmlFile2(strOutputPath);
}
