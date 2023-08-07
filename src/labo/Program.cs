// このコードは実験場なので使ってない宣言とか沢山

using Lucene.Net.Analysis.Ja.TokenAttributes;
using Lucene.Net.Analysis.Ja;
using Lucene.Net.Analysis.TokenAttributes;
using Lucene.Net.Analysis;
using Microsoft.VisualBasic.Devices;
using Microsoft.VisualBasic.Logging;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Lucene.Net.Analysis.Ja.Dict;
using System.ComponentModel;

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
	[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
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
	static extern bool ImmSetOpenStatus(IntPtr himc, bool unnamedParam2);

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

	private const int WM_SETFOCUS = 0x0007;
	private const int WM_KILLFOCUS = 0x0008;
	private const int WM_ACTIVATE = 0x0006;
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

	private const int VK_SPACE = 0x20;
	private const int VK_RETURN = 0x0D;
	private const int VK_CONTROL = 0x11;
	private const int VK_HOME = 0x24;
	private const int VK_DELETE = 0x2E;
	private const int VK_A = 0x41;

	const int WA_ACTIVE = 1;
	const int WA_CLICKACTIVE = 2;
	const int WA_INACTIVE = 0;

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
		/*
		English2Kana.Init();
		Console.WriteLine(English2Kana.Convert("aioueo"));
		Console.WriteLine(English2Kana.Convert("testなのだ"));
		Console.WriteLine(English2Kana.Convert("ずんだもんtestなのだ"));
		Console.WriteLine(English2Kana.Convert("VoicePeak"));
		Console.ReadLine();
		*/
		Impl01();
	}

	static void Impl01() {
		IntPtr.Size.WriteLine("ポインタサイズ={0}");
		var p = System.Diagnostics.Process.GetProcesses().Where(x => {
			try {
				return x.MainModule?.ModuleName?.ToLower() == "voicepeak.exe";
			}
			catch(Exception e) when ((e is Win32Exception) || (e is InvalidOperationException)) {
				return false;
			}
		}).FirstOrDefault();
		if(p == null) {
			return;
		}
		var target = p.MainWindowHandle;
		Console.WriteLine(p.MainWindowTitle);
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
		string speechText) {

		static void click(IntPtr hwnd, int x, int y) {
			var pos = x | (y << 16);
			PostMessage(hwnd, WM_LBUTTONDOWN, (IntPtr)1, (IntPtr)pos).WriteLine();
			PostMessage(hwnd, WM_LBUTTONUP, (IntPtr)0, (IntPtr)pos).WriteLine();
		}

		static void keyboard(IntPtr hwnd, int keycode) {
			PostMessage(hwnd, WM_KEYDOWN, (IntPtr)keycode, (IntPtr)(0x000000001));
			//Thread.Sleep(50);
			PostMessage(hwnd, WM_KEYUP, (IntPtr)keycode, (IntPtr)0xC00000001);
			//Thread.Sleep(100);
		}

		click(target, 400, 140);

		// テキスト送信
		foreach(var c in speechText) {
			SendMessage(target, WM_IME_CHAR, (IntPtr)c, IntPtr.Zero).WriteLine();
		}
		Thread.Sleep(50 * speechText.Length);

		// 再生
		var dll = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "voicepeak-hook", "bin", "x64", "Debug", "voicepeak-hook.dll");
		var hHookDll = LoadLibrary(dll);
		var hookFunc = Marshal.GetDelegateForFunctionPointer<HookVoicePeakProc>(GetProcAddress(hHookDll, "HookVoicePeak"));
		var unhookFunc = Marshal.GetDelegateForFunctionPointer<UnhookVoicePeakProc>(GetProcAddress(hHookDll, "UnhookVoicePeak"));

		hookFunc(target).WriteLine();
		SendMessage(target, msg, (IntPtr)2, (IntPtr)1).WriteLine();
		click(target, clientWidth / 2 + 125, 20);
		click(target, clientWidth / 2 + 165, 20);
		SendMessage(target, msg, (IntPtr)2, (IntPtr)0).WriteLine();
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
		if(!string.IsNullOrEmpty(speechText)) {
			keyboard(target, VK_HOME);
			foreach(var _ in speechText) {
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

	[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
	static extern IntPtr CreateFile(string pszFileName, int dwAccess, int dwShare, IntPtr psa, int dwCreatDisposition, int dwFlagsAndAttributes, IntPtr hTemplate);

	[DllImport("kernel32.dll")]
	static extern bool ReadFile(IntPtr hFile, byte[] pBuffer, int nNumberOfBytesToRead, out int pNumberOfBytesRead, IntPtr pOverlapped);

	[DllImport("kernel32.dll")]
	static extern int GetFileSize(IntPtr hFile, IntPtr pFileSizeHigh);
	[DllImport("kernel32.dll")]
	static extern bool CloseHandle(IntPtr hObject);

	[DllImport("kernel32.dll")]
	static extern int WaitForSingleObject(IntPtr hHandle, int dwMilliseconds);
	const int WAIT_TIMEOUT = 0x102;


	static void Impl02() {
		Speech2("てすとなのだ");
		Speech2("ずんだもんなのだ");
		Speech2("ずんだもちたべたいのだ");
		Speech2("ずんだもんなのだ");
		Speech2("ずん ずん ずん ずん");
	}

	static void Speech2(string speech) {
		speech.WriteLine();

		var voicePeak = @"C:\Program Files\VOICEPEAK\voicepeak.exe";
		var fileName = "vp.wav";

		if(File.Exists(fileName)) {
			File.Delete(fileName);
		}

		IntPtr hFile;
		unchecked {
			hFile = CreateFile(
				fileName,
				(int)0x80000000,
				0x00000001 | 2,
				IntPtr.Zero,
				1,
				0,
				IntPtr.Zero);
			hFile.WriteLine();
		}

		var sw = new Stopwatch();
		sw.Start();
		var p = Process.Start(new ProcessStartInfo() {
			FileName = voicePeak,
			Arguments = $"-s \"{speech}\" -n Zundamon -o {fileName}",
			RedirectStandardOutput = true,
		});

		var bufferedWaveProvider = new BufferedWaveProvider(new WaveFormat(48000, 16, 1));
		var wavProvider = new VolumeWaveProvider16(bufferedWaveProvider);
		//wavProvider.Volume = 0.1f;
		// TODO using
		var mmDevice = new MMDeviceEnumerator()
			.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);

		WasapiOut wavPlayer = null;
		try {
			Task.Run(() => {
				byte[] b = new byte[76800]; // 1秒間のデータサイズ
				var pos = 0;
				var len = 0;
				var sw = new Stopwatch();
				while(ReadFile(
					hFile,
					b,
					b.Length,
					out var ret,
					IntPtr.Zero
						)) {
					if(0 < ret) {
						if(pos == 0) {
							sw.Start();
							int head = 104; // voicepeakが出力するデータ領域開始アドレス
							bufferedWaveProvider.AddSamples(b, head, ret - head);
						} else {
							bufferedWaveProvider.AddSamples(b, 0, ret);
						}
						pos += ret;
					}
					var exit = WaitForSingleObject(p.Handle, 1000);
					if(exit == WAIT_TIMEOUT) {
						continue;
					}
					if(p.ExitCode != 0) {
						return;
					}
					if(len == 0) {
						len = GetFileSize(hFile, IntPtr.Zero);
					}
					if(ret == 0 && len <= pos) {
						sw.Stop();
						while(sw.ElapsedMilliseconds < (pos / 76800d) * 1000) {
							sw.Start();
							Thread.Sleep(1);
							sw.Stop();
						}
						wavPlayer?.Stop();
						break;
					}
				}
			});

			wavPlayer = new WasapiOut(mmDevice, AudioClientShareMode.Shared, false, 200);
			wavPlayer.Init(wavProvider);
			wavPlayer.Play();

			while(wavPlayer.PlaybackState == PlaybackState.Playing) {
				if((WaitForSingleObject(p.Handle, 10) != WAIT_TIMEOUT) && (p.ExitCode != 0)) {
					break;
				}
				System.Threading.Thread.Sleep(100);
			}
			p.WaitForExit();
			//p.ExitCode.WriteLine();
			sw.Stop();
			sw.ElapsedMilliseconds.WriteLine();
		}
		finally {
			if(hFile != IntPtr.Zero) {
				CloseHandle(hFile);
			}
			wavPlayer?.Dispose();
		}
	}

	static void Impl03() {
		var target = FindWindow(null, "VOICEPEAK");
		var juce = FindWindow(null, "JUCEWindow");
		//var tidMe = GetCurrentThreadId();
		var tidTarget = GetWindowThreadProcessId(target, out var pid);

		var msg = RegisterWindowMessage("yarukizero-vp-connect");
		GetWindowRect(target, out var rc);
		SetWindowPos(target, IntPtr.Zero, rc.left, rc.top, 1024, 877, 0);
		GetClientRect(target, out rc);

		Thread.Sleep(3000);
		var width = rc.right - rc.left;
		Speech3(target, juce, msg, "てすとなのだ");
		Speech3(target, juce, msg, "ずんだもんなのだ");
		Speech3(target, juce, msg, "ずんだもちたべたいのだ");
		Speech3(target, juce, msg, "ずんだもんなのだ");
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

	static void Speech3(
		IntPtr target,
		IntPtr juce,
		int msg,
		string speechText) {

		static void keyboard(IntPtr hwnd, int keycode) {
			PostMessage(hwnd, WM_KEYDOWN, (IntPtr)keycode, (IntPtr)(0x000000001));
			//Thread.Sleep(50);
			PostMessage(hwnd, WM_KEYUP, (IntPtr)keycode, (IntPtr)0xC00000001);
			//Thread.Sleep(100);
		}


		speechText.WriteLine();
		SendMessage(juce, WM_ACTIVATE, (IntPtr)WA_CLICKACTIVE, IntPtr.Zero);

		SendMessage(target, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);
		// テキスト送信
		foreach(var c in speechText) {
			SendMessage(target, WM_IME_CHAR, (IntPtr)c, IntPtr.Zero);
		}
		Thread.Sleep(50 * speechText.Length);

		SendMessage(target, WM_KILLFOCUS, IntPtr.Zero, IntPtr.Zero);
		// 再生
		SendMessage(target, WM_IME_CHAR, (IntPtr)VK_SPACE, IntPtr.Zero);

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
		SendMessage(target, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);
		if(!string.IsNullOrEmpty(speechText)) {
			keyboard(target, VK_HOME);
			foreach(var _ in speechText) {
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

	static class English2Kana {
		private const string KuromojiDicFile = "kuromoji.dic";
		private const string KanaDicFile = "english.dic";
		private static UserDictionary? s_userDictionary;
		private static Dictionary<string, string> s_kanaDic = new Dictionary<string, string>();

		public static void Init() {
			var dir = Path.GetDirectoryName(typeof(English2Kana).Assembly.Location) ?? "";
			try {
				foreach(var it in File.ReadAllLines(Path.Combine(dir, KanaDicFile))) {
					if(it.Any() && (it.Split('\t') is string[] s) && (s.Count() == 2)) {
						s_kanaDic.Add(s[0].ToLower(), s[1]);
					}
				}
			}
			catch(IOException) { }

			try {
				s_userDictionary = new UserDictionary(new StringReader(File.ReadAllText(Path.Combine(dir, KuromojiDicFile))));
			}
			catch(Exception) { }
		}


		public static string Convert(string text) {
			static string kana(string text) {
				if(s_kanaDic.TryGetValue(text.ToLower(), out var result)) {
					return result;
				} else {
					return ToKana(text);
				}
			}

			if(s_userDictionary == null) {
				return text;
			}

			var r = new StringBuilder();
			using var reader = new StringReader(text);
			using var tokenizer = new JapaneseTokenizer(reader, s_userDictionary, false, JapaneseTokenizerMode.SEARCH);

			using var ts = new TokenStreamComponents(tokenizer, tokenizer).TokenStream;
			ts.Reset();
			while(ts.IncrementToken()) {
				var term = ts.GetAttribute<ICharTermAttribute>();
				if(Regex.Match(term.ToString(), "^[a-z]+$", RegexOptions.IgnoreCase).Success) {
					var read = ts.AddAttribute<IReadingAttribute>().GetReading();
					r.Append(read switch {
						string v => v,
						_ => kana(term.ToString()),
					});
				} else {
					r.Append(term.ToString());
				}
			}
			return r.ToString();
		}

		private static string ToKana(string text) {
			static bool boin(char c) => c switch {
				'a' => true,
				'i' => true,
				'u' => true,
				'e' => true,
				'o' => true,
				_ => false,
			};

			var r = new StringBuilder();
			var last = '\0';
			IEnumerable<char> s = text.ToLower();
			while(s.Any()) {
				if(4 <= s.Count()) {
					var c1 = s.First();
					var c2 = s.Skip(1).First();
					var c3 = s.Skip(2).First();
					var c4 = s.Skip(3).First();
					last = c4;

					if((c1 == 't') && (c2 == 'c') && (c3 == 'h') && (c4 == 'i')) {
						r.Append("っち");
						s.Skip(4);
						continue;
					}

					if(boin(last) && (c1 == 'h')) {
						last = c1;
						if((c2 == 's') && (c3 == 'h') && (c4 == 'i')) {
							r.Append("ー");
							s = s.Skip(1);
							continue;
						} else if((c2 == 'c') && (c3 == 'h') && (c4 == 'i')) {
							r.Append("ー");
							s = s.Skip(1);
							continue;
						} else if((c2 == 't') && (c3 == 'h') && (c4 == 'i')) {
							r.Append("ー");
							s = s.Skip(1);
							continue;
						}
					}
				}
				if(3 <= s.Count()) {
					var c1 = s.First();
					var c2 = s.Skip(1).First();
					var c3 = s.Skip(2).First();
					last = c3;
					if((c1 == 's') && (c2 == 'h') && (c3 == 'i')) {
						r.Append("し");
						s.Skip(3);
						continue;
					} else if((c1 == 'c') && (c2 == 'h') && (c3 == 'i')) {
						r.Append("ち");
						s.Skip(3);
						continue;
					} else if((c1 == 't') && (c2 == 's') && (c3 == 'u')) {
						r.Append("つ");
						s.Skip(3);
						continue;
					}

					if((c1 == 'c') && (c2 == 'h') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ちゃ",
							'i' => "ち",
							'u' => "ちゅ",
							'e' => "ちぇ",
							'o' => "ちょ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'c') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ちゃ",
							'i' => "ちぃ",
							'u' => "ちゅ",
							'e' => "ちぇ",
							'o' => "ちょ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 't') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ちゃ",
							'i' => "ちぃ",
							'u' => "ちゅ",
							'e' => "ちぇ",
							'o' => "ちょ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					}


					if((c1 == 'm') && (c2 == 'm') && boin(c3)) {
						last = c1;
						r.Append("ん");
						s.Skip(1);
						continue;
					} else if((c1 == 'm') && (c2 == 'b') && boin(c3)) {
						last = c1;
						r.Append("ん");
						s.Skip(1);
						continue;
					} else if((c1 == 'm') && (c2 == 'p') && boin(c3)) {
						last = c1;
						r.Append("ん");
						s.Skip(1);
						continue;
					}


					if(boin(last) && (c1 == 'h')) {
						last = c1;
						if((c2 switch {
							'k' => true,
							's' => true,
							't' => true,
							'n' => true,
							'h' => true,
							'm' => true,
							'y' => true,
							'r' => true,
							'w' => true,
							'g' => true,
							'z' => true,
							'd' => true,
							'b' => true,
							'p' => true,
							'j' => true,
							_ => false,
						}) && boin(c3)) {
							r.Append("ー");
							s = s.Skip(1);
							continue;
						}
					}
				}

				if(2 <= s.Count()) {
					var c1 = s.First();
					var c2 = s.Skip(1).First();
					last = c2;
					if((c1 switch {
						'k' => true,
						's' => true,
						't' => true,
						'n' => true,
						'h' => true,
						'm' => true,
						'y' => true,
						'r' => true,
						'w' => true,
						'g' => true,
						'z' => true,
						'd' => true,
						'b' => true,
						'p' => true,
						_ => false,
					}) && boin(c2)) {
						r.Append(c1 switch {
							'k' => c2 switch {
								'a' => "か",
								'i' => "き",
								'u' => "く",
								'e' => "け",
								'o' => "こ",
								_ => throw new InvalidOperationException(),
							},
							's' => c2 switch {
								'a' => "さ",
								'i' => "し",
								'u' => "す",
								'e' => "せ",
								'o' => "そ",
								_ => throw new InvalidOperationException(),
							},
							't' => c2 switch {
								'a' => "た",
								'i' => "ち",
								'u' => "つ",
								'e' => "て",
								'o' => "と",
								_ => throw new InvalidOperationException(),
							},
							'n' => c2 switch {
								'a' => "な",
								'i' => "に",
								'u' => "ぬ",
								'e' => "ね",
								'o' => "の",
								_ => throw new InvalidOperationException(),
							},
							'h' => c2 switch {
								'a' => "は",
								'i' => "ひ",
								'u' => "ふ",
								'e' => "へ",
								'o' => "ほ",
								_ => throw new InvalidOperationException(),
							},
							'm' => c2 switch {
								'a' => "ま",
								'i' => "み",
								'u' => "む",
								'e' => "め",
								'o' => "も",
								_ => throw new InvalidOperationException(),
							},
							'y' => c2 switch {
								'a' => "や",
								'i' => "い",
								'u' => "ゆ",
								'e' => "え",
								'o' => "よ",
								_ => throw new InvalidOperationException(),
							},
							'r' => c2 switch {
								'a' => "ら",
								'i' => "り",
								'u' => "る",
								'e' => "れ",
								'o' => "ろ",
								_ => throw new InvalidOperationException(),
							},
							'w' => c2 switch {
								'a' => "わ",
								'i' => "うぃ",
								'u' => "う",
								'e' => "うぇ",
								'o' => "を",
								_ => throw new InvalidOperationException(),
							},
							'g' => c2 switch {
								'a' => "が",
								'i' => "ぎ",
								'u' => "ぐ",
								'e' => "げ",
								'o' => "ご",
								_ => throw new InvalidOperationException(),
							},
							'z' => c2 switch {
								'a' => "ざ",
								'i' => "じ",
								'u' => "ず",
								'e' => "ぜ",
								'o' => "ぞ",
								_ => throw new InvalidOperationException(),
							},
							'd' => c2 switch {
								'a' => "だ",
								'i' => "ぢ",
								'u' => "づ",
								'e' => "で",
								'o' => "ど",
								_ => throw new InvalidOperationException(),
							},
							'b' => c2 switch {
								'a' => "ば",
								'i' => "び",
								'u' => "ぶ",
								'e' => "べ",
								'o' => "ぼ",
								_ => throw new InvalidOperationException(),
							},
							'p' => c2 switch {
								'a' => "ぱ",
								'i' => "ぴ",
								'u' => "ぷ",
								'e' => "ぺ",
								'o' => "ぽ",
								_ => throw new InvalidOperationException(),
							},
							_ => throw new InvalidOperationException(),
						});
						s = s.Skip(2);
						continue;
					} else if((c1 == 'j') && boin(c2)) {
						r.Append(c1 switch {
							'a' => "じゃ",
							'i' => "じぃ",
							'u' => "じゅ",
							'e' => "じぇ",
							'o' => "じょ",
							_ => false,
						});
						s = s.Skip(2);
						continue;
					}
				}

				{
					var c1 = s.First();
					last = c1;
					s = s.Skip(1);


					r.Append(c1 switch {
						'a' => "あ",
						'b' => "べ",
						'c' => "く",
						'd' => "で",
						'e' => "え",
						'f' => "ふ",
						'g' => "ぐ",
						'h' => "ち",
						'i' => "い",
						'j' => "じ",
						'k' => "け",
						'l' => "る",
						'm' => "む",
						'n' => "ん",
						'o' => "お",
						'p' => "ぷ",
						'q' => "く",
						'r' => "ら",
						's' => "す",
						't' => "て",
						'u' => "う",
						'v' => "ヴ",
						'w' => "う",
						'x' => "",
						'y' => "い",
						'z' => "じ",

						'-' => "ー",
						_ => " ", // 区切るためスペースを入れる
					});
				}
			}
			return r.ToString();
		}
	}
}
