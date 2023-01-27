// このコードは実験場なので使ってない宣言とか沢山

using Microsoft.VisualBasic.Devices;
using Microsoft.VisualBasic.Logging;
using System;
using System.Drawing;
using System.Reflection;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows.Forms;


namespace Yarukizero.Net.VocePeakConnect.Lab;

static class Extension {
	public static T WriteLine<T>(this T @this) {
		Console.WriteLine(@this);
		return @this;
	}

	public static T WriteLine<T>(this T @this, string format) {
		Console.WriteLine(format, @this);
		return @this;
	}
}

class Program {
	[DllImport("user32.dll", CharSet = CharSet.Unicode)]
	static extern int RegisterWindowMessage(string lpString);
	[DllImport("kernel32.dll")]
	static extern IntPtr OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);
	[DllImport("kernel32.dll", SetLastError = true)]
	static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, IntPtr dwSize, int flAllocationType, int flProtect);
	[DllImport("kernel32.dll")]
	static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, string lpBuffer, IntPtr nSize, IntPtr lpNumberOfBytesWritten);
	[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError =true)]
	static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);
	[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
	static extern IntPtr GetModuleHandle(string lpLibFileName);
	[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
	static extern IntPtr LoadLibrary(string lpLibFileName);

	[DllImport("kernel32.dll")]
	static extern IntPtr CreateRemoteThread(
		IntPtr hProcess,
		IntPtr lpThreadAttributes, int dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter,
		int dwCreationFlags, IntPtr lpThreadId);

	[DllImport("user32.dll", CharSet = CharSet.Unicode)]
	static extern IntPtr PostMessage(IntPtr hwnd, int msg, IntPtr wp, IntPtr lp);
	[DllImport("user32.dll", CharSet = CharSet.Unicode)]
	static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wp, IntPtr lp);
	[DllImport("user32.dll", CharSet = CharSet.Unicode)]
	static extern IntPtr SendMessage(IntPtr hwnd, int msg, string wp, int lp);

	[DllImport("imm32.dll")]
	static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hwnd);

	[DllImport("imm32.dll")]

	static extern IntPtr ImmGetContext(IntPtr hIme);
	[DllImport("imm32.dll")]

	static extern bool ImmReleaseContext(IntPtr hwnd, IntPtr himc);

	[DllImport("imm32.dll")]
	static extern bool ImmSetOpenStatus(IntPtr himc,　bool unnamedParam2);

	[DllImport("imm32.dll", CharSet = CharSet.Unicode)]

	static extern bool ImmSetCompositionString(IntPtr himc, int dwIndex, string lpComp, int dwCompLen, IntPtr lpRead, int dwReadLen);

	[DllImport("user32.dll", CharSet = CharSet.Unicode)]
	static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

	[DllImport("user32.dll", CharSet = CharSet.Unicode)]
	static extern IntPtr FindWindow(string pClassName, string pWindowName);

	[DllImport("user32.dll")]
	static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
	[DllImport("user32.dll")]
	static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
	[DllImport("user32.dll")]
	static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);

	[DllImport("user32.dll")]
	static extern bool SetForegroundWindow(IntPtr hWnd);
	[DllImport("user32.dll")]
	static extern bool SetFocus(IntPtr hWnd);
	[DllImport("user32.dll")]
	static extern bool SetCursorPos(int X, int Y);
	[DllImport("user32.dll")]
	static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);


	struct IMECHARPOSITION {
		public int dwSize;
		public int dwCharPos;
		public int pt_x;
		public int pt_y;
		public int cLineHeight;
		public int rcDocument_0;
		public int rcDocument_1;
		public int rcDocument_2;
		public int rcDocument_3;
	}

	private const int SCS_SETSTR = 0x0001 | 0x0008;
	private const int WM_IME_COMPOSITION = 0x10F;
	private const int WM_IME_CONTROL = 0x283;
	private const int IMC_GETCONVERSIONMODE = 0x0001;
	private const int IMC_SETCONVERSIONMODE = 0x0002;
	private const int IMC_GETSENTENCEMODE = 0x0003;
	private const int IMC_SETSENTENCEMODE = 0x0004;
	private const int IMC_GETOPENSTATUS = 0x0005;
	private const int IMC_SETOPENSTATUS = 0x0006;

	private const int IMC_GETCANDIDATEPOS = 0x0007;
	private const int IMC_SETCANDIDATEPOS = 0x0008;
	private const int IMC_GETCOMPOSITIONFONT = 0x0009;
	private const int IMC_SETCOMPOSITIONFONT = 0x000A;
	private const int IMC_GETCOMPOSITIONWINDOW = 0x000B;
	private const int IMC_SETCOMPOSITIONWINDOW = 0x000C;
	private const int IMC_GETSTATUSWINDOWPOS = 0x000F;
	private const int IMC_SETSTATUSWINDOWPOS = 0x0010;
	private const int IMC_CLOSESTATUSWINDOW = 0x0021;
	private const int IMC_OPENSTATUSWINDOW = 0x0022;
	private const int GCS_COMPREADSTR = 1;
	private const int GCS_COMPREADATTR = 2;
	private const int GCS_COMPREADCLAUSE = 4;
	private const int GCS_COMPSTR = 8;
	private const int GCS_COMPATTR = 0x10;
	private const int GCS_COMPCLAUSE = 0x20;
	private const int GCS_CURSORPOS = 0x80;
	private const int GCS_DELTASTART = 0x100;
	private const int GCS_RESULTREADSTR = 0x200;
	private const int GCS_RESULTREADCLAUSE = 0x400;
	private const int GCS_RESULTSTR = 0x800;
	private const int GCS_RESULTCLAUSE = 0x1000;
	private const int CS_INSERTCHAR = 0x2000;

	private const int PROCESS_CREATE_THREAD = 0x0002;
	private const int PROCESS_VM_OPERATION = 0x0008;
	private const int PROCESS_VM_WRITE = 0x0020;

	private const int MEM_COMMIT = 0x1000;
	private const int PAGE_READWRITE = 0x4;


	private const int HC_ACTION = 0;
	private const int WH_MOUSE = 7;
	private const int WH_MOUSE_LL = 14;

	private const int WM_LBUTTONDOWN = 0x201;
	private const int WM_LBUTTONUP = 0x202;
	private const int WM_MBUTTONDOWN = 0x207;
	private const int WM_MBUTTONUP = 0x208;
	private const int WM_RBUTTONDOWN = 0x204;
	private const int WM_RBUTTONUP = 0x205;
	private const int WM_MOUSEMOVE = 0x200;
	private const int WM_MOUSEWHEEL = 0x20A;
	private const int WM_MOUSELEAVE = 0x02A3;
	private const int WM_XBUTTONDOWN = 0x20B;
	private const int WM_XBUTTONUP = 0x20C;
	private const int WM_KEYDOWN = 0x0100;
	private const int WM_KEYUP = 0x0101;
	private const int WM_CHAR = 0x0102;
	private const int WM_IME_CHAR = 0x286;

	private const int VK_RETURN = 0x0D;
	private const int VK_CONTROL = 0x11;
	private const int VK_HOME = 0x24;
	private const int VK_DELETE = 0x2E;
	private const int VK_A = 0x41;

	[StructLayout(LayoutKind.Sequential)]
	struct POINT {
		public int x;
		public int y;
	}
	struct RECT {
		public int left;
		public int top;
		public int right;
		public int bottom;
	}

	[StructLayout(LayoutKind.Sequential)]
	struct MSLLHOOKSTRUCT {
		public POINT pt;
		public int mouseData;
		public int flags;
		public int time;
		public IntPtr dwExtraInfo;
	}

	[DllImport("kernel32.dll")]
	static extern int GetCurrentThreadId();
	[DllImport("user32.dll")]
	static extern bool AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);

	[DllImport("user32.dll")]
	private static extern bool GetCursorPos(ref POINT pt);
	[DllImport("user32.dll")]
	private static extern IntPtr SetWindowsHookEx(int idHook, MouseProc lpfn, IntPtr hMod, int dwThreadId);
	[DllImport("user32.dll")]
	private static extern IntPtr CallNextHookEx(IntPtr hHook, int nCode, IntPtr wParam, ref MSLLHOOKSTRUCT lParam);
	[DllImport("user32.dll")]
	private static extern bool UnhookWindowsHookEx(IntPtr hHook);
	[UnmanagedFunctionPointer(CallingConvention.Cdecl)]
	private delegate IntPtr MouseProc(int nCode, IntPtr wParam, ref MSLLHOOKSTRUCT lParam);

	private delegate bool HookVoicePeakProc(IntPtr hwnd);
	private delegate bool UnhookVoicePeakProc(IntPtr hwnd);

	[STAThread]
	static void Main(string[] args) {
		IntPtr.Size.WriteLine("ポインタサイズ={0}");

		var target = FindWindow(null, "VOICEPEAK");
		//var tidMe = GetCurrentThreadId();
		var tidTarget = GetWindowThreadProcessId(target, out var pid);

		var msg = RegisterWindowMessage("yarukizero-vp-connect");
		GetWindowRect(target, out var rc);
		SetWindowPos(target, IntPtr.Zero, rc.left, rc.top, 1024, 877, 0);
		GetClientRect(target, out rc);

		Thread.Sleep(3000);
		var width = rc.right - rc.left;
		Speech(target, msg, width, "てすとなのだ");
		Speech(target, msg, width, "ずんだもんなのだ");
		Speech(target, msg, width, "ずんだもちたべたいのだ");
		Speech(target, msg, width, "ずんだもんなのだ");
		return;
		/*
		この先DLLインジェクションでなんとかしようとしていた後
		if (SendMessage(target, RegisterWindowMessage("yarukizero-vp-connect"), (IntPtr)1, (IntPtr)1).WriteLine().ToInt32() != 0) {
			goto next;
		}

		var hproc = OpenProcess(
			PROCESS_CREATE_THREAD | PROCESS_VM_OPERATION | PROCESS_VM_WRITE,
			false,
			pid).WriteLine();
		var dll = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "x64", "Debug", "vocepeak-hook.dll");
		var pRemoteLib = VirtualAllocEx(
			hproc,
			IntPtr.Zero,
			(IntPtr)(dll.Length * 2 + 2),
			MEM_COMMIT,
			PAGE_READWRITE).WriteLine();
		WriteProcessMemory(
			hproc,
			pRemoteLib,
			dll,
			(IntPtr)(dll.Length * 2) + 2,
			IntPtr.Zero).WriteLine();
		var pLoadLib = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryW").WriteLine();
		CreateRemoteThread(
			hproc, IntPtr.Zero, 0,
			pLoadLib, pRemoteLib,
			0, IntPtr.Zero).WriteLine();
	next:;
		"start".WriteLine();
		//*/
		//AttachThreadInput(tidMe, tidTarget, true).WriteLine();
	}
	static void Speech(
		IntPtr target,
		int msg,
		int clientWidth,
		string speachText) {

		static void click(IntPtr hwnd, int x, int y) {
			var pos = x | (y << 16);
			PostMessage(hwnd, WM_LBUTTONDOWN, (IntPtr)1, (IntPtr)pos).WriteLine();
			PostMessage(hwnd, WM_LBUTTONUP, (IntPtr)0, (IntPtr)pos).WriteLine();
		}

		static void keyboard(IntPtr hwnd, int keycode) {
			PostMessage(hwnd, WM_KEYDOWN, (IntPtr)keycode, (IntPtr)(0x000000001));
			Thread.Sleep(50);
			PostMessage(hwnd, WM_KEYUP, (IntPtr)keycode, (IntPtr)0xC00000001);
			Thread.Sleep(100);
		}

		click(target, 400, 140);

		// テキスト送信
		foreach (var c in speachText) {
			SendMessage(target, WM_IME_CHAR, (IntPtr)c, IntPtr.Zero).WriteLine();
		}
		Thread.Sleep(50 * speachText.Length);

		// 再生
		var dll = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "vocepeak-hook", "bin", "x64", "Debug", "vocepeak-hook.dll");
		var hHookDll = LoadLibrary(dll);
		var hookFunc = Marshal.GetDelegateForFunctionPointer<HookVoicePeakProc>(GetProcAddress(hHookDll, "HookVoicePeak"));
		var unhookFunc = Marshal.GetDelegateForFunctionPointer<UnhookVoicePeakProc>(GetProcAddress(hHookDll, "UnhookVoicePeak"));

		hookFunc(target).WriteLine();
		SendMessage(target, msg, (IntPtr)1, (IntPtr)1).WriteLine();
		click(target, clientWidth / 2 + 165, 20);
		SendMessage(target, msg, (IntPtr)1, (IntPtr)0).WriteLine();
		unhookFunc(target).WriteLine();

		// 適当に再生終了をまつ
		Thread.Sleep(5 * 1000);

		// 後始末
#if !true
		click(target, 400, 140);
		Thread.Sleep(100);
		PostMessage(target, WM_KEYDOWN, (IntPtr)VK_CONTROL, (IntPtr)0x401d001);
		PostMessage(target, WM_KEYDOWN, (IntPtr)VK_A, (IntPtr)0x001E001);
		PostMessage(target, WM_KEYDOWN, (IntPtr)VK_CONTROL, (IntPtr)0x401d001);
		PostMessage(target, WM_KEYUP, (IntPtr)VK_A, (IntPtr)0xC01E001);
		PostMessage(target, WM_KEYUP, (IntPtr)VK_CONTROL, (IntPtr)0xC01D001);
		//Console.ReadLine();
		PostMessage(target, WM_KEYDOWN, (IntPtr)VK_DELETE, (IntPtr)0x0153001);
		Thread.Sleep(100);
		PostMessage(target, WM_KEYUP, (IntPtr)VK_DELETE, (IntPtr)0xC153001);
		//Console.ReadLine();
#elif true
		click(target, 400, 140);
		if(!string.IsNullOrEmpty(speachText)) {
			keyboard(target, VK_HOME);
			foreach(var _ in speachText) {
				keyboard(target, VK_DELETE);
			}
		}
#else
		click(target, 6, 140);
		PostMessage(target, WM_KEYDOWN, (IntPtr)VK_RETURN, (IntPtr)0x0153001);
		Sleep(100);
		PostMessage(target, WM_KEYUP, (IntPtr)VK_RETURN, (IntPtr)0xC153001);
#endif
	}
}