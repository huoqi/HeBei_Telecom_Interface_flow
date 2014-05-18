// XunJianDlg.cpp : implementation file
//
#include "stdafx.h"
#include "XunJian.h"
#include "XunJianDlg.h"
#include "excel.h"
#include "ExcelFile.h"
#include "ReadTxt.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

bool opened = false;

CString PROJ[/*6*/]={"     �ֵ�     ","�˿ں�","�˿�����","ʵʱ����(bps)","������","��ʷ����û���"};
CString	COMMAND[]={ "terminal length 0",                // 0
					"show ospf nei",                    // 1
					"show arp",                         // 2
					"show hard",                        // 3 
					"show port counter",                // 4
					"show ip route summary all"         // 5
};

CXunJianDlg::CXunJianDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CXunJianDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CXunJianDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	xj_ipListFileName = _T("ip_list.txt");
}

void CXunJianDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CXunJianDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
	DDX_Text(pDX, IDC_HOSTLISTFILENAME, xj_ipListFileName);
}

BEGIN_MESSAGE_MAP(CXunJianDlg, CDialog)
	//{{AFX_MSG_MAP(CXunJianDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDOK, &CXunJianDlg::OnBnClickedOk)
	ON_EN_CHANGE(IDC_HOSTLISTFILENAME, &CXunJianDlg::OnEnChangeHostlistfilename)
	ON_BN_CLICKED(IDC_OPENFILE, &CXunJianDlg::OnBnClickedOpenfile)
//	ON_EN_CHANGE(IDC_EDIT1, &CXunJianDlg::OnEnChangeEdit1)
	ON_LBN_SELCHANGE(IDC_HOSTLIST, &CXunJianDlg::OnLbnSelchangeHostlist)
	ON_BN_CLICKED(IDCANCEL, &CXunJianDlg::OnBnClickedCancel)
	ON_NOTIFY(NM_THEMECHANGED, IDC_SCROLLBAR1, &CXunJianDlg::OnNMThemeChangedScrollbar1)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CXunJianDlg message handlers

BOOL CXunJianDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CXunJianDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CXunJianDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CXunJianDlg::OnOK() 
{
	try{
	pList= (CListBox *)GetDlgItem(IDC_HOSTLIST);
	/*if( pList->GetTextLen(0)>15 || pList->GetTextLen(0)<7 )
	{
		MessageBox("�����б��ļ�δ���أ�������ѡ��");
		return;
	}*/

	_GUID clsid;
	IUnknown *pUnk;
	IDispatch *pDisp;
	LPDISPATCH lpDisp;

	_Application app;
	Workbooks xj_books;
	_Workbook xj_book;
	Worksheets xj_sheets;
	_Worksheet xj_sheet;
	Range range;
	Range unionRange;
	Range cols;

	Font font;
//	COleVariant background;

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	::CLSIDFromProgID(L"Excel.Application",&clsid); // from registry
	if(GetActiveObject(clsid, NULL,&pUnk) == S_OK)
	{
		VERIFY(pUnk->QueryInterface(IID_IDispatch,(void**) &pDisp) == S_OK);
		app.AttachDispatch(pDisp);
		pUnk->Release();
	} 
	else
	{  
		if(!app.CreateDispatch("Excel.Application"))
		{
			MessageBox("Excel program not found");     
			app.Quit();     
			return;
		}
	}

	xj_books=app.GetWorkbooks();
	xj_book=   xj_books.Add(covOptional);
	xj_sheets= xj_book.GetSheets();
	xj_sheet=  xj_sheets.GetItem(COleVariant((short)1));

	int i;
	Range item;
	range=xj_sheet.GetRange(COleVariant("A1"),COleVariant("F1"));
	for(i= 0; i < 6; i++)
	{
		item.AttachDispatch(range.GetItem(COleVariant((long)1),COleVariant((long)i+1)).pdispVal);
		item.SetValue2(COleVariant(PROJ[i]));
	}  //����һ��Ŀ¼

	//range=xj_sheet.GetRange(COleVariant("A1"),COleVariant("F1"));
	lpDisp=range.GetInterior();
	Interior   cellinterior;
	cellinterior.AttachDispatch(lpDisp);
	cellinterior.SetColor(COleVariant((long)0xc0c0c0));  //���ñ���ɫΪ��ɫ
	cellinterior.ReleaseDispatch();
	//range=xj_sheet.GetRange(COleVariant("A1"),COleVariant("F1"));
	range.SetHorizontalAlignment(COleVariant((long)-4108)); //ȫ������
	Borders bord;
	bord=range.GetBorders();
	bord.SetLineStyle(COleVariant((short)1));  //���ñ߿�
	//range=xj_sheet.GetRange(COleVariant("A1"),COleVariant("F1"));
	cols=range.GetEntireColumn();
	cols.AutoFit();  //�Զ�����

/**************************����ʼ�滭���************************************/

	long usedRowNum; //�м���
	CString handleFile;
	CString hostFileName,hostip;
	bool error = false;
	CString infos,info;
	ExcelFile excelFile;
	ReadTxt xj_txt;
	xj_HostCount=pList->GetCount();
	for(int n_host=0;n_host<xj_HostCount;n_host++)  //��ѭ����һ���ļ�һ��ѭ����
	{		
		pList->GetText(n_host,hostFileName);
		hostip = hostFileName;
		handleFile = hostFileName + _T(" ���ڴ���...");
		pList->DeleteString(n_host);
		pList->InsertString(n_host,handleFile);
		pList->SetCurSel(n_host);
		pList->UpdateWindow();

		hostFileName = xj_FilePath + hostFileName;
		hostFileName += _T(".txt");
		CStdioFile hostFile;
		if(!hostFile.Open(hostFileName,CFile::modeRead,0))
		{  //��¼�������ļ���
			handleFile.Replace("���ڴ���...","ʧ�ܣ�");
			error = true;
			pList->DeleteString(n_host);
			pList->InsertString(n_host,handleFile);
			pList->UpdateWindow();
			continue;
		}
		usedRowNum = excelFile.GetRowCount(xj_sheet);
		range.AttachDispatch(xj_sheet.GetCells());

		//info.Format( _T("%d"), n_host+1);
		info = xj_txt.ReadHostName(&hostFile,COMMAND[0],COMMAND[1]);  //��ȡ�ڵ�����
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(1)),COleVariant(info));

		int portCount = 0;   //�˿���Ŀ��������7/1
		CString nSend, nRecv;
		float n_Send,n_Recv;

		while(hostFile.ReadString(info))
			if(info.Find( COMMAND[4]) > -1) break;
		while( hostFile.ReadString(info) && info.Find( "[local]" ) == -1 )  //�˿ںź�����
		{
			if( info.Find( "/" ) == -1 || info.Find( "7/1" ) > -1 ) continue;

			info.Replace( "ethernet","");
			info = _T("'") + info;
			infos = info;

			while( hostFile.ReadString(info) )
				if( info.Find( "send bit rate" ) > -1 ) break;
			nSend = info.Mid( 60 );
			hostFile.ReadString(info);
			nRecv = info.Mid( 60 );
			nSend.Trim();
			nRecv.Trim();
			n_Send = (float)atof(nSend);
			n_Recv = (float)atof(nRecv);

			if( n_Send < 1000 && n_Recv < 1000 ) continue;
			portCount++;
			range.SetItem(COleVariant(usedRowNum+portCount),COleVariant(long(2)),COleVariant(infos.Trim()));
			range.SetItem(COleVariant(usedRowNum+portCount),COleVariant(long(4)),COleVariant((n_Send>n_Recv)?nSend:nRecv));
		}

		hostFile.SeekToBegin();
		infos = xj_txt.ReadLine(&hostFile,"ubscriber Address");  //��ʷ��������û���
		if( infos == _T("") ) info = _T("0");
		else
		{
			int token = 0;
			for(i = 0; i < 5 ; i++) info = infos.Tokenize(" ",token);
		}
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(6)),COleVariant(info.Trim()));  

		hostFile.Close();

		if(portCount > 1)
		{
			unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)1)).pdispVal);
			unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)portCount),COleVariant((long)1)));
			unionRange.Merge(COleVariant((long)0));  //�ڵ����Ƶ�Ԫ��ϲ�

			unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)6)).pdispVal);
			unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)portCount),COleVariant((long)1)));
			unionRange.Merge(COleVariant((long)0)); 			
			//��ʷ����û����ϲ�
		}

		unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)1)).pdispVal);
		unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)portCount),COleVariant((long)6)));
		unionRange.SetRowHeight(COleVariant(13.5));
		bord = unionRange.GetBorders();
		bord.SetLineStyle(COleVariant((short)1));  //���ñ߿�



		handleFile.Replace("���ڴ���...","�����");
		pList->DeleteString(n_host);
		pList->InsertString(n_host,handleFile);
		pList->UpdateWindow();
	}

	CTime time;
	time = time.GetCurrentTime();
	infos = time.Format("%Y%m%d%H%M%S");  //time.Format();
	info = _T("Ѳ�챨��") + infos + _T(".xlsx");
	info = xj_FilePath + info;
	info.Replace("\\\\","\\");

	xj_book.SaveAs(COleVariant(info),covOptional,covOptional,covOptional,covOptional,covOptional,0,covOptional,covOptional,covOptional,covOptional,covOptional);
	
	if(error == true )
	{
		MessageBox("Ѳ�챨������ɣ��ѱ��浽\r\n" + info + "\r\n���ļ��򿪴��󣬵��\"ȷ��\"���ز鿴","���ļ��򿪴���",MB_OK|MB_ICONWARNING);
		app.Quit();
	}
	else 
	{
		if(MessageBox("Ѳ�챨������ɣ��ѱ��浽\r\n" + info + "\r\n���\"ȷ��\"���ļ��鿴","���ɱ������",MB_OKCANCEL) == IDOK)
		{
			app.SetVisible(TRUE);
			app.SetUserControl(TRUE);
		}
		else app.Quit();
	}
}
catch (CFileException* e)
    {
        e->ReportError();
        e->Delete();
    }
	//CDialog::OnOK();
}

void CXunJianDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	OnOK();
}

void CXunJianDlg::OnEnChangeHostlistfilename()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ�������������
	// ���͸�֪ͨ��������д CDialog::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}

void CXunJianDlg::OnBnClickedOpenfile()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CStdioFile listfile;
	if( opened == true || listfile.Open(xj_ipListFileName,CFile::modeRead,0) == false)
	{
		//listfile.Close();
		m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
		CFileDialog FileDlg(true, _T("txt"),NULL,OFN_FILEMUSTEXIST|OFN_HIDEREADONLY, 
										 "�ı��ļ�(*.TXT)|*.TXT|All Files(*.*)|*.*||"); 
		if( FileDlg.DoModal() == IDOK )
		{ 
			xj_ipListFileName = FileDlg.GetFileName();
			listfile.Open(xj_ipListFileName,CFile::modeRead,0);
		}
		else return;
	}
	xj_ipListFileName = listfile.GetFileName();//FileDlg.GetFileName();
	xj_FilePath = listfile.GetFilePath();//FileDlg.GetPathName();
	xj_FilePath.Replace(xj_ipListFileName,"");
	xj_FilePath.Replace("\\","\\\\");
	pList= (CListBox *)GetDlgItem(IDC_HOSTLIST);
    pList->ResetContent();   
    CString str;
    while(listfile.ReadString(str))   
	{
		if(str.Find("\'") == -1)  //���в����� ,�� '���Ų����룬�൱��ע�͵�����
		{
			str.Replace(" ","");  //ȥ�����пո�
			pList->AddString(str);
		}
    }
	opened = true;
    listfile.Close(); 
	UpdateData(false);
}


void CXunJianDlg::OnLbnSelchangeHostlist()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
}

void CXunJianDlg::OnBnClickedCancel()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	OnCancel();
}

void CXunJianDlg::OnNMThemeChangedScrollbar1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// �ù���Ҫ��ʹ�� Windows XP ����߰汾��
	// ���� _WIN32_WINNT ���� >= 0x0501��
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	*pResult = 0;
}
