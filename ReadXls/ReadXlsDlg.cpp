
// ReadXlsDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "ReadXls.h"
#include "ReadXlsDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CReadXlsDlg �Ի���


IMPLEMENT_DYNAMIC(CReadXlsDlg, CDialogEx);

CReadXlsDlg::CReadXlsDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CReadXlsDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CReadXlsDlg::~CReadXlsDlg()
{
	// ����öԻ������Զ���������
	//  ���˴���ָ��öԻ���ĺ���ָ������Ϊ NULL���Ա�
	//  �˴���֪���öԻ����ѱ�ɾ����
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


// CReadXlsDlg ��Ϣ�������

BOOL CReadXlsDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	ShowWindow(SW_SHOW);


	// TODO:  �ڴ���Ӷ���ĳ�ʼ������
	if (!m_obExcel.initExcel())
	{
		return false;
	}
	m_oldLocale = _strdup(setlocale(LC_CTYPE, NULL));
	setlocale(LC_CTYPE, "chs");//�趨
	m_strInputName = "";

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CReadXlsDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CReadXlsDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// ���û��ر� UI ʱ������������Ա���������ĳ��
//  �������Զ�����������Ӧ�˳���  ��Щ
//  ��Ϣ�������ȷ����������: �����������ʹ�ã�
//  ������ UI�������ڹرնԻ���ʱ��
//  �Ի�����Ȼ�ᱣ�������

void CReadXlsDlg::OnClose()
{
	if (CanExit())
	{
		m_obExcel.release();
		//��ԭ�����趨
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
	// �����������Ա�����������Զ���
	//  �������Իᱣ�ִ�Ӧ�ó���
	//  ʹ�Ի���������������� UI ����������
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}



void CReadXlsDlg::OnBnClickedToxml()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	//ʹ��Excel��
	CString strInputFile;
	CString strOutputPath;
	m_editInputPath.GetWindowTextW(strInputFile);
	m_editOutputPath.GetWindowTextW(strOutputPath);
	
	if (0 == strInputFile.GetLength())
	{
		MessageBox(_T("����ѡ����Ҫ��ȡ���ļ�"), _T("����"), MB_OK);
		return;
	}
	if (0 == strOutputPath.GetLength())
	{
		MessageBox(_T("����ѡ�����·��"), _T("����"), MB_OK);
		return;
	}
	
	if (!m_obExcel.open(__CString2Constchar(strInputFile)))
	{
		MessageBox(_T("�޷��򿪸��ļ�"), _T("����"), MB_OK);
		return;
	}
	int nSheetCount = m_obExcel.getSheetCount();

	//������excel����m_mapSheetList��
	for (int s = 1; s <= nSheetCount; ++s)
	{
		CString strSheetName = m_obExcel.getSheetName(s);//��ȡsheet��
		m_vecSheetName.push_back(strSheetName);
		bool bLoad = m_obExcel.loadSheet(strSheetName);//װ��sheet  
		int nRow = m_obExcel.getRowCount();//��ȡsheet������
		int nCol = m_obExcel.getColumnCount();//��ȡsheet������  


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
	
	//��m_mapSheetList����д��xml�ļ�

	//��һ��ģʽ
	__CreateXmlFile(strOutputPath);

// 	//�ڶ���ģʽ
// 	__CreateXmlFile2(strOutputPath);
	
}

void CReadXlsDlg::OnBnClickedExit()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	SendMessage(WM_CLOSE, 0, 0);
}


void CReadXlsDlg::OnBnClickedOpenfile()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	// ���ù�����   
	TCHAR szFilter[] = _T("Microsoft Excel 2007 ������(*.xlsx)|*.xlsx|�����ļ�(*.*)|*.*||");
	// ������ļ��Ի���   
	CFileDialog fileDlg(TRUE, /*_T("xls")*/NULL, NULL, 0, szFilter, this);
	CString strFilePath;

	// ��ʾ���ļ��Ի���   
	if (IDOK == fileDlg.DoModal())
	{
		// ���������ļ��Ի����ϵġ��򿪡���ť����ѡ����ļ�·����ʾ���༭����   
		strFilePath = fileDlg.GetPathName();
		SetDlgItemText(IDC_InputPath, strFilePath);

		m_strInputName = __GetFileName(strFilePath);
	}
}


void CReadXlsDlg::OnBnClickedSavepath()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	// ���ù�����   
	TCHAR szFilter[] = _T("xml�ļ�(*.xml)|*.xml|�����ļ�(*.*)|*.*||");
	//Ĭ������ļ���
	CString strDefName;
	if (0 == m_strInputName.GetLength())
	{
		strDefName = "my";
	}
	else
	{
		strDefName = m_strInputName;
	}

	// ���챣���ļ��Ի���   
	CFileDialog fileDlg(FALSE, _T("xml"), strDefName, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, this);
	CString strFilePath;

	// ��ʾ�����ļ��Ի���   
	if (IDOK == fileDlg.DoModal())
	{
		// ���������ļ��Ի����ϵġ����桱��ť����ѡ����ļ�·����ʾ���༭����   
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
		MessageBox(_T("���ɳɹ�"), _T("�ɹ�"), MB_OK);
	}
	else
	{
		MessageBox(_T("����ʧ��"), _T("ʧ��"), MB_OK);
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
		MessageBox(_T("���ɳɹ�"), _T("�ɹ�"), MB_OK);
	}
	else
	{
		MessageBox(_T("����ʧ��"), _T("ʧ��"), MB_OK);
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
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString strInputFile;
	CString strOutputPath;
	m_editInputPath.GetWindowTextW(strInputFile);
	m_editOutputPath.GetWindowTextW(strOutputPath);

	if (0 == strInputFile.GetLength())
	{
		MessageBox(_T("����ѡ����Ҫ��ȡ���ļ�"), _T("����"), MB_OK);
		return;
	}
	if (0 == strOutputPath.GetLength())
	{
		MessageBox(_T("����ѡ�����·��"), _T("����"), MB_OK);
		return;
	}

	if (!m_obExcel.open(__CString2Constchar(strInputFile)))
	{
		MessageBox(_T("�޷��򿪸��ļ�"), _T("����"), MB_OK);
		return;
	}
	int nSheetCount = m_obExcel.getSheetCount();

	//������excel����m_mapSheetList��
	for (int s = 1; s <= nSheetCount; ++s)
	{
		CString strSheetName = m_obExcel.getSheetName(s);//��ȡsheet��
		m_vecSheetName.push_back(strSheetName);
		bool bLoad = m_obExcel.loadSheet(strSheetName);//װ��sheet  
		int nRow = m_obExcel.getRowCount();//��ȡsheet������
		int nCol = m_obExcel.getColumnCount();//��ȡsheet������  


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

	//��m_mapSheetList����д��xml�ļ�

// 	//��һ��ģʽ
// 	__CreateXmlFile(strOutputPath);

	//�ڶ���ģʽ
	__CreateXmlFile2(strOutputPath);
}
