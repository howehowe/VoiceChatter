
// VoiceChatterDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "VoiceChatter.h"
#include "VoiceChatterDlg.h"


#include <InitGuid.h>
#include <uuids.h>

#include "debug_util.h"
#include "dshow_enum_helper.h"


using boost::intrusive_ptr;
using std::wstring;

#define DEFAULT_BUFFER_TIME ((float) 0.05)  /* 10 milliseconds*/

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

WinVersion GetWinVersion() {
    static bool checked_version = false;
    static WinVersion win_version = WINVERSION_PRE_2000;
    if (!checked_version) {
        OSVERSIONINFOEX version_info;
        version_info.dwOSVersionInfoSize = sizeof version_info;
        GetVersionEx(reinterpret_cast<OSVERSIONINFO*>(&version_info));
        if (version_info.dwMajorVersion == 5) {
            switch (version_info.dwMinorVersion) {
        case 0:
            win_version = WINVERSION_2000;
            break;
        case 1:
            win_version = WINVERSION_XP;
            break;
        case 2:
        default:
            win_version = WINVERSION_SERVER_2003;
            break;
            }
        } else if (version_info.dwMajorVersion == 6) {
            if (version_info.wProductType != VER_NT_WORKSTATION) {
                // 2008 is 6.0, and 2008 R2 is 6.1.
                win_version = WINVERSION_2008;
            } else {
                if (version_info.dwMinorVersion == 0) {
                    win_version = WINVERSION_VISTA;
                } else {
                    win_version = WINVERSION_WIN7;
                }
            }
        } else if (version_info.dwMajorVersion > 6) {
            win_version = WINVERSION_WIN7;
        }
        checked_version = true;
    }
    return win_version;
}

HRESULT GetPin( IBaseFilter * pFilter, PIN_DIRECTION dirrequired, int iNum, IPin **ppPin)
{
    CComPtr< IEnumPins > pEnum;
    *ppPin = NULL;

    if (!pFilter)
        return E_POINTER;

    HRESULT hr = pFilter->EnumPins(&pEnum);
    if(FAILED(hr)) 
        return hr;

    ULONG ulFound;
    IPin *pPin;
    hr = E_FAIL;

    while(S_OK == pEnum->Next(1, &pPin, &ulFound))
    {
        PIN_DIRECTION pindir = (PIN_DIRECTION)3;

        pPin->QueryDirection(&pindir);
        if(pindir == dirrequired)
        {
            if(iNum == 0)
            {
                *ppPin = pPin;  // Return the pin's interface
                hr = S_OK;      // Found requested pin, so clear error
                break;
            }
            iNum--;
        } 

        pPin->Release();
    } 

    return hr;
}

void UtilDeleteMediaType(AM_MEDIA_TYPE *pmt)
{
    // Allow NULL pointers for coding simplicity
    if (pmt == NULL) {
        return;
    }

    // Free media type's format data
    if (pmt->cbFormat != 0) 
    {
        CoTaskMemFree((PVOID)pmt->pbFormat);

        // Strictly unnecessary but tidier
        pmt->cbFormat = 0;
        pmt->pbFormat = NULL;
    }

    // Release interface
    if (pmt->pUnk != NULL) 
    {
        pmt->pUnk->Release();
        pmt->pUnk = NULL;
    }

    // Free media type
    CoTaskMemFree((PVOID)pmt);
}

// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
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

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CVoiceChatterDlg 对话框




CVoiceChatterDlg::CVoiceChatterDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CVoiceChatterDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CVoiceChatterDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_COMBO1, m_microphoneCombox);
    DDX_Control(pDX, IDC_COMBO2, m_microphonePinCombox);
}

BEGIN_MESSAGE_MAP(CVoiceChatterDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
    ON_BN_CLICKED(IDOK, &CVoiceChatterDlg::OnBnClickedOk)
    ON_BN_CLICKED(IDC_BUTTON1, &CVoiceChatterDlg::OnBnClickedButton1)
    ON_BN_CLICKED(IDC_BUTTON2, &CVoiceChatterDlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// CVoiceChatterDlg 消息处理程序

BOOL CVoiceChatterDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    m_captureFilter = -1;
    m_capturePin = -1;

    HRESULT r = initCapture();
    if (FAILED(r))
        PostQuitMessage(0);

    initGraph();


    if (!checkCapture())
    {
        m_microphoneCombox.ShowWindow(SW_SHOW);
        m_microphonePinCombox.ShowWindow(SW_SHOW);

        m_microphoneCombox.ResetContent();
        for (int i = 0; i < m_captureFilterVec.size(); i++)
        {
            m_microphoneCombox.AddString(m_captureFilterVec[i].Name.c_str());
        }
        m_microphoneCombox.SetCurSel(0);

        m_microphonePinCombox.ResetContent();

        for (int i = 0; i < m_captureFilterVec[0].PinVec.size(); i++)
        {
            m_microphonePinCombox.AddString(m_captureFilterVec[0].PinVec[i].c_str());
        }

        m_microphonePinCombox.SetCurSel(0);

        m_captureFilter = 0;
        m_capturePin = 0;
    }
    else
    {
        if (!buildMicrophoneGraph())
            PostQuitMessage(0);   
    }



	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CVoiceChatterDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CVoiceChatterDlg::OnPaint()
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
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CVoiceChatterDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

HRESULT CVoiceChatterDlg::initGraph()
{
    HRESULT r;
    r = CoCreateInstance(CLSID_FilterGraph, NULL, CLSCTX_INPROC_SERVER, 
        IID_IGraphBuilder, (void**)&m_graphBuilder);

    if (FAILED(r) || !m_graphBuilder)
        return E_NOINTERFACE;

    // Get useful interfaces from the graph builder
    r = m_graphBuilder->QueryInterface(IID_IMediaControl,
        (void **)&m_mediaControl);

    if (FAILED(r))
        return r;

    return r;
}

HRESULT CVoiceChatterDlg::initCapture()
{
    int mixerNum = 0;
    int microphoneNum = 0;
    intrusive_ptr<ICreateDevEnum> sysDevEnum = NULL;
    HRESULT r = CoCreateInstance(CLSID_SystemDeviceEnum, NULL, 
        CLSCTX_INPROC, IID_ICreateDevEnum, 
        (void **)&sysDevEnum);

    if FAILED(r)
        return r;

    intrusive_ptr<IEnumMoniker> pEnumCat = NULL;
    r = sysDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, (IEnumMoniker**)&pEnumCat, 0);

    if (SUCCEEDED(r))
    {
        // Enumerate all filters using the category enumerator
        r = enumFiltersAndMonikersToList(pEnumCat.get());

    }


    intrusive_ptr<IMoniker> pMoniker = NULL;
    intrusive_ptr<IBaseFilter> filter;
    for (int i = 0;i < m_captureFilterVec.size(); i++)
    {
        pMoniker = m_captureFilterVec[i].Moniker;
        // Use the moniker to create the specified audio capture device
        r = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (void**)&filter);   
        if (FAILED(r))
            return r;

        r = enumPinsOnFilter(filter.get(), PINDIR_INPUT, i);
        if (FAILED(r))
            return r;

        //filter->Release();
    }

    return S_OK;

}

HRESULT CVoiceChatterDlg::enumFiltersAndMonikersToList( IEnumMoniker *pEnumCat )
{
    HRESULT r=S_OK;
    IMoniker* pMoniker=0;
    ULONG cFetched=0;
    VARIANT varName={0};
    int nFilters=0;

    // If there are no filters of a requested type, show default string
    if (!pEnumCat)
    {
        return S_FALSE;
    }

    // Enumerate all items associated with the moniker
    while(pEnumCat->Next(1, &pMoniker, &cFetched) == S_OK)
    {
        intrusive_ptr<IPropertyBag> pPropBag;
        ASSERT(pMoniker);

        // Associate moniker with a file
        r = pMoniker->BindToStorage(0, 0, IID_IPropertyBag, 
            (void **)&pPropBag);
        ASSERT(SUCCEEDED(r));
        ASSERT(pPropBag);
        if (FAILED(r))
            continue;

        // Read filter name from property bag
        varName.vt = VT_BSTR;
        r = pPropBag->Read(L"FriendlyName", &varName, 0);
        if (FAILED(r))
            continue;

        // Get filter name (converting BSTR name to a CString)
        CString str(varName.bstrVal);
        SysFreeString(varName.bstrVal);
        nFilters++;

        TCaptureFilter filter = {str, pMoniker};
        m_captureFilterVec.push_back(filter);

        // Cleanup interfaces
        //SAFE_RELEASE(pPropBag);

        // Intentionally DO NOT release the pMoniker, since it is
        // stored in a listbox for later use
    }

    return r;
}

HRESULT CVoiceChatterDlg::enumPinsOnFilter(IBaseFilter *pFilter, PIN_DIRECTION PinDir, int index)
{
    HRESULT r;
    intrusive_ptr<IEnumPins> pEnum = NULL;
    intrusive_ptr<IPin> pPin = NULL;

    // Verify filter interface
    if (!pFilter)
        return E_NOINTERFACE;

    // Get pin enumerator
    r = pFilter->EnumPins((IEnumPins **)&pEnum);
    if (FAILED(r))
        return r;

    pEnum->Reset();

    // Enumerate all pins on this filter
    while((r = pEnum->Next(1, (IPin **)&pPin, 0)) == S_OK)
    {
        PIN_DIRECTION PinDirThis;

        r = pPin->QueryDirection(&PinDirThis);
        if (FAILED(r))
        {
            //pPin->Release();
            continue;
        }

        // Does the pin's direction match the requested direction?
        if (PinDir == PinDirThis)
        {
            PIN_INFO pininfo={0};

            // Direction matches, so add pin name to listbox
            r = pPin->QueryPinInfo(&pininfo);
            if (SUCCEEDED(r))
            {
                wstring str = pininfo.achName;
                m_captureFilterVec[index].PinVec.push_back(str);
            }

            // The pininfo structure contains a reference to an IBaseFilter,
            // so you must release its reference to prevent resource a leak.
            pininfo.pFilter->Release();
        }
        //pPin->Release();
    }
    //pEnum->Release();

    return r;
}
void CVoiceChatterDlg::OnBnClickedOk()
{
    // TODO: 在此添加控件通知处理程序代码
    m_mediaControl->Stop();
    CoUninitialize();
    OnOK();
}

bool CVoiceChatterDlg::buildMicrophoneGraph( int filterIndex, int pinIndex )
{
    HRESULT r;
    int index = -1;

    if (filterIndex != -1)
    {
        index = findFilter((IBaseFilter**)&m_pInputDevice, NULL, filterIndex);
    }
    else
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
            index = findFilter((IBaseFilter **)&m_pInputDevice, L"麦克风");
        else
            index = findFilter((IBaseFilter **)&m_pInputDevice);
    }


    if (index == -1)
        return false;

    if (pinIndex != -1)
    {
        activePin(m_pInputDevice.get(), index, NULL, pinIndex);
    }
    else
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
            activePin(m_pInputDevice.get(), index);
        else
        {
            activePin(m_pInputDevice.get(), index, L"麦克风");
        }
    }


    SetAudioProperties(m_pInputDevice.get());

    r = m_graphBuilder->AddFilter(m_pInputDevice.get(), L"Microphone Capture");
    if (FAILED(r))
        return false;

    intrusive_ptr<IBaseFilter> renderer;

    r = CoCreateInstance(CLSID_AudioRender, 
        NULL, CLSCTX_INPROC_SERVER,IID_IBaseFilter,
        reinterpret_cast<void**>(&renderer));

    if (FAILED(r))
        return false;

    intrusive_ptr<IPin> sourceOutPin;
    r = GetPinByDirection(m_pInputDevice.get(), 
        reinterpret_cast<IPin**>(&sourceOutPin), 
        PINDIR_OUTPUT);

    if (FAILED(r))
        return false;

    r = m_graphBuilder->AddFilter((IBaseFilter *)renderer.get(), L"DShow Render");
    if (FAILED(r))
        return false;

    intrusive_ptr<IPin> rendererInputPin;
    r = GetPinByDirection(renderer.get(), 
        reinterpret_cast<IPin**>(&rendererInputPin), 
        PINDIR_INPUT);

    if (FAILED(r))
        return false;


    r = m_graphBuilder->ConnectDirect(sourceOutPin.get(), rendererInputPin.get(), NULL);
    if (FAILED(r))
        return false;  

    return true;
}

int CVoiceChatterDlg::findFilter( IBaseFilter** filter, wchar_t* str , int index )
{
    int r = -1;

    if (index != -1)
    {
        IMoniker *pMoniker = m_captureFilterVec[index].Moniker;
        HRESULT hr = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (void**)filter);
        if (FAILED(hr))
            return -1;

        return index;
    }

    if (str == NULL)
    {
        IMoniker *pMoniker = m_captureFilterVec[0].Moniker;
        HRESULT hr = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (void**)filter);
        if (FAILED(hr))
            return -1;

        return 0;
    }

    for (int i = 0; i < m_captureFilterVec.size(); i++)
    {
        if (m_captureFilterVec[i].Name.find(str) != wstring::npos)
        {
            IMoniker *pMoniker = m_captureFilterVec[i].Moniker;
            HRESULT hr = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (void**)filter);
            if (FAILED(hr))
                return -1;

            return i;
        }

    }       

    return -1;
}

bool CVoiceChatterDlg::checkCapture()
{
    int mixCount = 0;
    int captureCount = 0;

    for (int i = 0; i < m_captureFilterVec.size(); i++)
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
        {
            if (m_captureFilterVec[i].Name.find(L"麦克风") != wstring::npos || m_captureFilterVec[i].Name.find(L"Mic") != wstring::npos)
            {
                captureCount++;
            }

            if (m_captureFilterVec[i].Name.find(L"混音") != wstring::npos || m_captureFilterVec[i].Name.find(L"Mix") != wstring::npos)
            {
                mixCount++;
            }
        }
        else
        {
            for (int j = 0; j < m_captureFilterVec[i].PinVec.size(); j++)
            {
                if (m_captureFilterVec[i].PinVec[j].find(L"麦克风") != wstring::npos || m_captureFilterVec[i].PinVec[j].find(L"Mic") != wstring::npos)
                {
                    captureCount++;
                }

                if (m_captureFilterVec[i].PinVec[j].find(L"混音") != wstring::npos || m_captureFilterVec[i].PinVec[j].find(L"Mix") != wstring::npos)
                {
                    mixCount++;
                }
            }
        }

    }

    if (captureCount != 1 || mixCount != 1)
        return false;

    return true;
}

HRESULT CVoiceChatterDlg::activePin( IBaseFilter* filter, int index, wchar_t* str /*= NULL*/, int pinIndex /*= -1*/ )
{
    HRESULT r = -1;
    int nActivePin = -1;

    if (pinIndex == -1)
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
        {
            nActivePin = 0;
        }
        else
        {
            for (int j = 0; j < m_captureFilterVec[index].PinVec.size(); j++)
            {
                if (m_captureFilterVec[index].PinVec[j].find(str) != wstring::npos)
                {
                    nActivePin = j;
                    break;
                }
            }
        }
    }
    else
    {
        nActivePin = pinIndex;
    }

    intrusive_ptr<IPin> pPin=0;
    intrusive_ptr<IAMAudioInputMixer> pPinMixer;

    // How many pins are in the input pin list?
    int nPins = m_captureFilterVec[index].PinVec.size();


    // Activate the selected input pin and deactivate all others
    for (int i=0; i<nPins; i++)
    {
        // Get this pin's interface
        r = GetPin(filter, PINDIR_INPUT, i, (IPin **)&pPin);
        if (SUCCEEDED(r))
        {
            r = pPin->QueryInterface(IID_IAMAudioInputMixer, (void **)&pPinMixer);
            if (SUCCEEDED(r))
            {
                // If this is our selected pin, enable it
                if (i == nActivePin)
                {
                    // Set any other audio properties on this pin
                    //r = SetInputPinProperties(pPinMixer);

                    // If there is only one input pin, this method
                    // might return E_NOTIMPL.
                    r = pPinMixer->put_Enable(TRUE);
                }
                // Otherwise, disable it
                else
                {
                    //r = pPinMixer->put_Enable(FALSE);
                }

               // pPinMixer->Release();
            }

            // Release pin interfaces
           // pPin->Release();
        }
    }

    return S_OK;
}

HRESULT CVoiceChatterDlg::SetAudioProperties( IBaseFilter* filter )
{
    HRESULT hr=0;
    IPin *pPin=0;
    IAMBufferNegotiation *pNeg=0;
    IAMStreamConfig *pCfg=0;
    int nFrequency=8000;
    int nChannels = 2;
    int nBytesPerSample = 2;
    // Determine audio properties
    //     int nChannels = IsDlgButtonChecked(IDC_RADIO_MONO) ? 1 : 2;
    //     int nBytesPerSample = IsDlgButtonChecked(IDC_RADIO_8) ? 1 : 2;

    //     // Determine requested frequency
    //     if (IsDlgButtonChecked(IDC_RADIO_11KHZ))      
    //         nFrequency = 11025;
    //     else if (IsDlgButtonChecked(IDC_RADIO_22KHZ)) 
    //         nFrequency = 22050;
    //     else 
    //         nFrequency = 44100;

    // Find number of bytes in one second
    long lBytesPerSecond = (long) (nBytesPerSample * nFrequency * nChannels);

    // Set to 50ms worth of data    
    long lBufferSize = (long) ((float) lBytesPerSecond * DEFAULT_BUFFER_TIME);

    for (int i=0; i<2; i++)
    {
        hr = GetPin(filter, PINDIR_OUTPUT, i, &pPin);
        if (SUCCEEDED(hr))
        {
            // Get buffer negotiation interface
            hr = pPin->QueryInterface(IID_IAMBufferNegotiation, (void **)&pNeg);
            if (FAILED(hr))
            {
                pPin->Release();
                return hr;
            }

            // Set the buffer size based on selected settings
            ALLOCATOR_PROPERTIES prop={0};
            prop.cbBuffer = lBufferSize;
            prop.cBuffers = 6;
            prop.cbAlign = nBytesPerSample * nChannels;
            hr = pNeg->SuggestAllocatorProperties(&prop);
            pNeg->Release();

            // Now set the actual format of the audio data
            hr = pPin->QueryInterface(IID_IAMStreamConfig, (void **)&pCfg);
            if (FAILED(hr))
            {
                pPin->Release();
                return hr;
            }            

            // Read current media type/format
            AM_MEDIA_TYPE *pmt={0};
            hr = pCfg->GetFormat(&pmt);

            if (SUCCEEDED(hr))
            {
                // Fill in values for the new format
                WAVEFORMATEX *pWF = (WAVEFORMATEX *) pmt->pbFormat;
                pWF->nChannels = (WORD) nChannels;
                pWF->nSamplesPerSec = nFrequency;
                pWF->nAvgBytesPerSec = lBytesPerSecond;
                pWF->wBitsPerSample = (WORD) (nBytesPerSample * 8);
                pWF->nBlockAlign = (WORD) (nBytesPerSample * nChannels);

                // Set the new formattype for the output pin
                hr = pCfg->SetFormat(pmt);
                UtilDeleteMediaType(pmt);
            }

            // Release interfaces
            pCfg->Release();
            pPin->Release();
        }
        else
            break;
    }
    // No more output pins on this filter
    return hr;
}
void CVoiceChatterDlg::OnBnClickedButton1()
{
    // TODO: 在此添加控件通知处理程序代码
    m_mediaControl->Run();
}

void CVoiceChatterDlg::OnBnClickedButton2()
{
    // TODO: 在此添加控件通知处理程序代码
    m_mediaControl->Pause();
}
