#include "pch.hpp"
#if false // 使ってない

/* extern "C" __declspec(dllexport) */ bool WINAPI AttachVoicePeak(HWND hVoicePeak);


namespace {
	const TCHAR* MSG_VP_CONNECT = TEXT("yarukizero-vp-connect");
	const TCHAR* FILENAME_HOOK2 = TEXT("yarukizero-vp-connect.hook2");
	const int VPC_MSG_CALLBACKWND = 1;
	const int VPC_MSG_ENABLEHOOK = 2;
	const int VPC_MSG_ENDSPEECH = 3;
	const int VPC_MSG_ENABLEHOOK2 = 4;
	const int VPC_HOOK_ENABLE = 1;
	const int VPC_HOOK_DISABLE = 0;

	// RegisterWindowMessageで登録したメッセージは外部プロセスに飛ばせない？

	const int VPCM_ENDSPEECH = (WM_USER + VPC_MSG_ENDSPEECH);

	/* なんかVOICEPEAK内で流れてくるメッセージ
	const int VP_MSG = 0x0118;
	const int VP_MSG_WPARAM = 0x0000FFFF;
	const int VP_MSG_LPARAM = 0x00000118;
	*/

	UINT gs_msgVpConnect = 0;
	bool gs_isHookAction = false;

	// DLLインジェクション処理
	WNDPROC pVpProc;

	LRESULT CALLBACK SubWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
		if (gs_msgVpConnect && (gs_msgVpConnect == msg)) {
			switch (wp) {
			case 0:
				return 1;
			case 1:
				gs_isHookAction = wp != 0;
				return 1;
			}
			return 0;
		}

		switch (msg) {
		case WM_MOUSELEAVE:
			if (gs_isHookAction) {
				return 0;
			}
			break;
		}

		return ::CallWindowProcW(pVpProc, hwnd, msg, wp, lp);
	}


	BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
		DWORD id1 = ::GetCurrentProcessId();
		DWORD id2;
		::GetWindowThreadProcessId(hwnd, &id2);
		if (id1 == id2) {
			TCHAR strTitle[64];
			::GetWindowText(hwnd, strTitle, 64);
			if (lstrcmpi(strTitle, TEXT("voicepeak")) == 0) {
				::AttachVoicePeak(hwnd);
				return FALSE;
			}
		}
		return TRUE;
	}
}

// DLLインジェクション処理
/* extern "C" __declspec(dllexport) */ bool WINAPI AttachVoicePeak(HWND hVoicePeak) {
	gs_msgVpConnect = ::RegisterWindowMessage(MSG_VP_CONNECT);
	pVpProc = reinterpret_cast<WNDPROC>(::GetWindowLongPtr(hVoicePeak, GWLP_WNDPROC));
	::SetWindowLongPtr(hVoicePeak, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(SubWndProc));

	return true;
}

void Attach() {
	//MessageBoxA(NULL, "test", "aaaa", 0);

	//::EnumWindows(EnumWindowsProc, NULL);
	::AttachVoicePeak(::FindWindow(NULL, TEXT("VOICEPEAK")));
}
#endif





