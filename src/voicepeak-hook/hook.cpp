#include "pch.h"

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

	HHOOK gs_hWndHook = NULL;
	HHOOK gs_hMsgHook = NULL;
	UINT gs_msgVpConnect = 0;
	bool gs_isHookAction = false;
	bool gs_isStartAction = false;
	bool gs_isStartSpeach = false;
	bool gs_isEndAction = false;
	HWND gs_hCallBackWnd;
	UINT_PTR gs_timer = NULL;
	ULONGLONG gs_startTime = 0;

	inline void click(HWND hwnd, int x, int y) {
		PostMessage(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(x, y));
		PostMessage(hwnd, WM_LBUTTONUP, 0, MAKELPARAM(x, y));
	}

#if false
	// DLLインジェクション処理
	WNDPROC pVpProc;

	LRESULT CALLBACK SubWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
		if(gs_msgVpConnect && (gs_msgVpConnect == msg)) {
			switch(wp) {
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
#endif
	void WaitTimerProc(HWND hWnd, UINT uMsg, UINT_PTR nIDEvent, DWORD dwTime) {
		::KillTimer(hWnd, ::gs_timer);
		::gs_timer = NULL;

		if (!::gs_isStartAction) {
			::gs_isStartAction = true;
			// フォーカスを削除してカーソルのWM_PAINTを抑制する
			::PostMessage(hWnd, WM_KILLFOCUS, 0, 0);
		}
	}

	void SpeechTimerProc(HWND hWnd, UINT uMsg, UINT_PTR nIDEvent, DWORD dwTime) {
		::KillTimer(hWnd, ::gs_timer);
		::gs_timer = NULL;

		if(!::gs_isStartSpeach) {
			::gs_isStartSpeach = true;
			RECT rc = { 0 };
			::GetClientRect(hWnd, &rc);
			auto w = rc.right - rc.left;
			::click(hWnd, w / 2 + 125, 20);
			::click(hWnd, w / 2 + 165, 20);
		} else if(!::gs_isEndAction) {
			::gs_isEndAction = true;

			if(::gs_hCallBackWnd) {
				::PostMessage(::gs_hCallBackWnd, ::VPCM_ENDSPEECH, 0, 0);
			}
		}
	}
}

#if false
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

extern "C" __declspec(dllexport) LRESULT CALLBACK HookWndProc(int code, WPARAM wParam, LPARAM lParam) {
	if(code == HC_ACTION) {
		auto pcwp = reinterpret_cast<CWPSTRUCT*>(lParam);
		if(!::gs_msgVpConnect) {
			::gs_msgVpConnect = ::RegisterWindowMessage(MSG_VP_CONNECT);
		}
		if (pcwp->message == ::gs_msgVpConnect) {
			switch (pcwp->wParam) {
			case VPC_MSG_ENABLEHOOK2:
				{
					auto hMapObj = CreateFileMapping(
						reinterpret_cast<HANDLE>(-1),
						0, PAGE_READONLY, 0, 2048,
						FILENAME_HOOK2);
					if(hMapObj && (GetLastError() == ERROR_ALREADY_EXISTS)) {
						auto speech = reinterpret_cast<PCWCHAR>(::MapViewOfFile(hMapObj, FILE_MAP_READ, 0, 0, 0));
						if(speech) {
							::gs_isHookAction = true;
							::gs_isStartAction = false;
							::gs_isStartSpeach = false;
							::gs_isEndAction = false;
							::gs_startTime = ::gs_isHookAction ? ::GetTickCount64() : 0;

							auto len = ::lstrlenW(speech);
							::click(pcwp->hwnd, 400, 140);
							for(auto i = 0; i < len; i++) {
								::SendMessage(pcwp->hwnd, WM_IME_CHAR, speech[i], 0);
							}
							::PostMessage(pcwp->hwnd, WM_KEYDOWN, VK_HOME, 0x000000001);
							::PostMessage(pcwp->hwnd, WM_KEYUP, VK_HOME, 0xC00000001);
							::gs_timer = ::SetTimer(pcwp->hwnd, 0, 500, ::WaitTimerProc);

							::UnmapViewOfFile(speech);
						}
						::CloseHandle(hMapObj);
					}
				}
				break;
			}
			pcwp->message = WM_NULL;
		}
	}
	return CallNextHookEx(::gs_hWndHook, code, wParam, lParam);
}

extern "C" __declspec(dllexport) LRESULT CALLBACK MsgHookProc(int code, WPARAM wParam, LPARAM lParam) {
	if(code == HC_ACTION) {
		if (!::gs_msgVpConnect) {
			::gs_msgVpConnect = ::RegisterWindowMessage(MSG_VP_CONNECT);
		}

		auto msg = reinterpret_cast<MSG*>(lParam);
		if(msg->message == ::gs_msgVpConnect) {
			switch(msg->wParam) {
			case VPC_MSG_CALLBACKWND:
				::gs_hCallBackWnd = reinterpret_cast<HWND>(msg->lParam);
				break;
			case VPC_MSG_ENABLEHOOK:
				::gs_isHookAction = msg->lParam != VPC_HOOK_DISABLE;
				::gs_isStartAction = true;
				::gs_isStartSpeach = true;
				::gs_isEndAction = false;
				::gs_startTime = ::gs_isHookAction ? ::GetTickCount64() : 0;
				break;
			}
			msg->message = WM_NULL;
		} else switch(msg->message) {
		// 処理中マウスが対象にいないため WM_MOUSELEAVE が来る
		// 握りつぶす
		case WM_MOUSELEAVE:
			if(::gs_isHookAction) {
				msg->message = WM_NULL;
			}
			break;
		case WM_PAINT:
			if(::gs_isHookAction && ::gs_isStartAction && (!::gs_isStartSpeach || !::gs_isEndAction)) {
				/*
				if((::GetTickCount64() - ::gs_startTime) < 500) {
					break;
				}
				*/
				if(::gs_timer) {
					::KillTimer(msg->hwnd, ::gs_timer);
					::gs_timer = NULL;
				}
				::gs_timer = ::SetTimer(msg->hwnd, 0, 1000, ::SpeechTimerProc);
			}
			break;
		}
	}
	return CallNextHookEx(::gs_hMsgHook, code, wParam, lParam);
}

extern "C" __declspec(dllexport) bool WINAPI HookVoicePeak(HWND hVoicePeak) {
	::gs_hWndHook = ::SetWindowsHookEx(
		WH_CALLWNDPROC,
		reinterpret_cast<HOOKPROC>(HookWndProc),
		::g_module,
		0);
	::gs_hMsgHook = ::SetWindowsHookEx(
		WH_GETMESSAGE,
		reinterpret_cast<HOOKPROC>(MsgHookProc),
		::g_module,
		0);
	//return ::gs_hMsgHook;
	if(::gs_hWndHook && ::gs_hMsgHook) {
		return true;
	} else {
		if(::gs_hWndHook) {
			::UnhookWindowsHookEx(::gs_hWndHook);
			::gs_hWndHook = NULL;
		}
		if(::gs_hMsgHook) {
			::UnhookWindowsHookEx(::gs_hMsgHook);
			::gs_hMsgHook = NULL;
		}
		return false;
	}
}

extern "C" __declspec(dllexport) bool WINAPI UnhookVoicePeak() {
	if(!::gs_hMsgHook) {
		return true;
	}

	::UnhookWindowsHookEx(::gs_hWndHook);
	::UnhookWindowsHookEx(::gs_hMsgHook);
	::gs_hWndHook = NULL;
	::gs_hMsgHook = NULL;

	return true;
}