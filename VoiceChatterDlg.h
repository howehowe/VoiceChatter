
// VoiceChatterDlg.h : 头文件
//

#pragma once

#include <string>
#include <vector>
#include <strmif.h>
#include <DShow.h>
#include <MMSystem.h>
#include "afxwin.h"

#include "intrusive_ptr_helper.h"

struct IGraphBuilder;
struct IBaseFilter;
struct IMoniker;
struct IMediaControl;

enum WinVersion {
    WINVERSION_PRE_2000 = 0,  // Not supported
    WINVERSION_2000 = 1,
    WINVERSION_XP = 2,
    WINVERSION_SERVER_2003 = 3,
    WINVERSION_VISTA = 4,
    WINVERSION_2008 = 5,
    WINVERSION_WIN7 = 6,
};


// CVoiceChatterDlg 对话框
class CVoiceChatterDlg : public CDialog
{
// 构造
public:
	CVoiceChatterDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_VOICECHATTER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
    CComboBox m_microphoneCombox;
    CComboBox m_microphonePinCombox;

private:
    struct TCaptureFilter
    {
        std::wstring Name;
        IMoniker* Moniker;
        std::vector<std::wstring> PinVec; 
    };
    std::vector<TCaptureFilter> m_captureFilterVec;

    boost::intrusive_ptr<IGraphBuilder> m_graphBuilder;
    boost::intrusive_ptr<IMediaControl> m_mediaControl;
    boost::intrusive_ptr<IBaseFilter> m_pInputDevice;

    int m_captureFilter;
    int m_capturePin;

    HRESULT initGraph();
    HRESULT initCapture();
    HRESULT enumFiltersAndMonikersToList(IEnumMoniker *pEnumCat);
    HRESULT enumPinsOnFilter (IBaseFilter *pFilter, PIN_DIRECTION PinDir, int index);
    bool buildMicrophoneGraph(int filterIndex = -1, int pinIndex = -1);
    int findFilter(IBaseFilter** filter, wchar_t* str = NULL, int index = -1);
    bool checkCapture(); 
    HRESULT activePin(IBaseFilter* filter, int index, wchar_t* str = NULL, int pinIndex = -1);
    HRESULT SetAudioProperties(IBaseFilter* filter);

public:
    afx_msg void OnBnClickedOk();
    afx_msg void OnBnClickedButton1();
    afx_msg void OnBnClickedButton2();
};
