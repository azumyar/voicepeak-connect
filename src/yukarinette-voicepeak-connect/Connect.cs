using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using static System.Net.Mime.MediaTypeNames;

namespace Yarukizero.Net.Yularinette.VoicePeakConnect {
	internal class Connect : IDisposable {
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		static extern int RegisterWindowMessage(string lpString);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern nint PostMessage(nint hwnd, int msg, nint wp, nint lp);
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern nint SendMessage(nint hwnd, int msg, nint wp, nint lp);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern nint LoadLibrary(string lpLibFileName);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
		static extern nint GetProcAddress(nint hModule, string lpProcName);

		[DllImport("user32.dll")]
		private static extern bool GetClientRect(nint hWnd, out RECT lpRect);
		[DllImport("user32.dll")]
		private static extern bool GetWindowRect(nint hWnd, out RECT lpRect);
		[DllImport("user32.dll")]
		private static extern bool SetWindowPos(nint hWnd, nint hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern int GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern nint CreateFileMapping(nint hFile, nint lpFileMappingAttributes, int flProtect, int dwMaximumSizeHigh, int dwMaximumSizeLow, string lpName);
		[DllImport("kernel32.dll")]
		private static extern nint MapViewOfFile(nint hFileMappingObject, int dwDesiredAccess, int dwFileOffsetHigh, int dwFileOffsetLow, nint dwNumberOfBytesToMap);
		[DllImport("kernel32.dll")]
		private static extern nint UnmapViewOfFile(nint hFileMappingObject);
		[DllImport("kernel32.dll")]
		private static extern nint CloseHandle(nint hObject);
		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern nint lstrcpy(nint str1, string str2);


		private const int WM_KILLFOCUS = 0x0008; 
		private const int WM_LBUTTONDOWN = 0x201;
		private const int WM_LBUTTONUP = 0x202;
		private const int WM_KEYDOWN = 0x0100;
		private const int WM_KEYUP = 0x0101;
		private const int WM_IME_CHAR = 0x286;
		private const int MK_LBUTTON = 0x01;
		private const int VK_HOME = 0x24;
		private const int VK_DELETE = 0x2E;
		private const int VK_SPACE = 0x20;

		private const int PAGE_READWRITE = 0x04;
		private const int FILE_MAP_WRITE = 0x00000002;

		[StructLayout(LayoutKind.Sequential)]
		struct POINT {
			public int x;
			public int y;
		}
		[StructLayout(LayoutKind.Sequential)]
		struct RECT {
			public int left;
			public int top;
			public int right;
			public int bottom;
		}

		[DllImport("dwmapi.dll")]
		private static extern uint DwmRegisterThumbnail(nint dest, nint src, out nint thumb);
		[DllImport("dwmapi.dll")]
		private static extern uint DwmUnregisterThumbnail(nint thumb);
		[DllImport("dwmapi.dll")]
		private static extern uint DwmUpdateThumbnailProperties(nint hThumb, in DWM_THUMBNAIL_PROPERTIES props);
		[StructLayout(LayoutKind.Sequential)]
		struct DWM_THUMBNAIL_PROPERTIES {
			public int dwFlags;
			public RECT rcDestination;
			public RECT rcSource;
			public byte opacity;
			public bool fVisible;
			public bool fSourceClientAreaOnly;
		}
		private const int DWM_TNP_RECTDESTINATION = 0x00000001;
		private const int DWM_TNP_RECTSOURCE = 0x00000002;
		private const int DWM_TNP_OPACITY = 0x00000004;
		private const int DWM_TNP_VISIBLE = 0x00000008;
		private const int DWM_TNP_SOURCECLIENTAREAONLY = 0x00000010;

		private delegate bool HookVoicePeakProc(IntPtr hwnd);
		private delegate bool UnhookVoicePeakProc(IntPtr hwnd);

		private const int VPC_MSG_CALLBACKWND = 1;
		private const int VPC_MSG_ENABLEHOOK = 2;
		private const int VPC_MSG_ENDSPEECH = 3;
		private const int VPC_MSG_ENABLEHOOK2 = 4;
		private const int VPC_HOOK_ENABLE = 1;
		private const int VPC_HOOK_DISABLE = 0;

		private const int WM_USER = 0x400;
		private const int VPCM_ENDSPEECH = (WM_USER + VPC_MSG_ENDSPEECH);
		private readonly string VpConnectMessage = "yarukizero-vp-connect";
		private readonly string MapNameHokk2 = "yarukizero-vp-connect.hook2";
		private readonly string MapNameCapture = "yarukizero-net-yukarinette.audio-capture";

		class MessageWindow : System.Windows.Window {
			private readonly Connect con;
			private System.Windows.Interop.HwndSource source;

			public IntPtr Handle { get; }

			public MessageWindow(Connect con) {
				this.con = con;
				this.ShowInTaskbar = false;
				this.Opacity = 0;
				this.Title = "voicepeak-connect";


				var helper = new System.Windows.Interop.WindowInteropHelper(this);
				this.Handle = helper.EnsureHandle();
				this.source = System.Windows.Interop.HwndSource.FromHwnd(helper.Handle);
				this.source.AddHook(WndProc);
				this.Loaded += async (_, _) => {
					await Task.Delay(100);
					this.Hide();
				};
			}

			private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
				if(msg == YACM_STARTUP) {
					this.con.captureWindow = lParam;
				} else if(msg == YACM_CAPTURE) {
					this.con.sync.Set();
					handled = true;
				}
				return IntPtr.Zero;
			}
		}

		class DwmWindow : System.Windows.Window {
			private readonly Connect con;

			public IntPtr Handle { get; }

			public DwmWindow(Connect con) {
				this.con = con;
				this.ShowInTaskbar = false;
				this.Title = "ボイスピミラー";


				var helper = new System.Windows.Interop.WindowInteropHelper(this);
				this.Handle = helper.EnsureHandle();
				/*
				this.Loaded += async (_, _) => {
					await Task.Delay(100);
					this.Hide();
				};
				*/
			}
		}

		private readonly AutoResetEvent sync = new AutoResetEvent(false);

		private int connectionMsg;
		private MessageWindow window;
		//private DwmWindow dwmWindow;
		private IntPtr formHandle;
		private nint dwmThumb;

		private IntPtr hHookModule;
		private HookVoicePeakProc hookFunc;
		private UnhookVoicePeakProc unhookFunc;

		private IntPtr hVoicePeak;
		private int voicePeakWidth;

		private readonly string dllPath;
		private readonly string exePath;
		private readonly string iniPath;
		private readonly int defaultWaitSec = 50;

		private Process captureProcess;
		private nint captureWindow;
		private const int WM_APP = 0x8000;
		private const int YACM_STARTUP = WM_APP + 1;
		private const int YACM_SHUTDOWN = WM_APP + 2;
		private const int YACM_GETSTATE = WM_APP + 4;
		private const int YACM_CAPTURE = WM_APP + 5;

		public Connect() {
			this.connectionMsg = RegisterWindowMessage(VpConnectMessage);
			this.exePath = Path.Combine(
				AppDomain.CurrentDomain.BaseDirectory,
				"Plugins",
				"Yarukizero.Net.VoicePeakConnect",
				"audio-capture.exe");
			/*
			this.dllPath = Path.Combine(
				AppDomain.CurrentDomain.BaseDirectory,
				"Plugins",
				"Yarukizero.Net.VoicePeakConnect",
				"voicepeak-hook.dll");
			*/
			this.iniPath = Path.Combine(
				AppDomain.CurrentDomain.BaseDirectory,
				"Plugins",
				"Yarukizero.Net.VoicePeakConnect",
				"voicepeak-connet.ini");

			this.window = new MessageWindow(this);
			this.window.Show();
			this.formHandle = this.window.Handle;
			//this.dwmWindow = new DwmWindow(this);
		}

		public void Dispose() {
			this.EndCaptureVoicePeak();
			this.window?.Close();
		}

		public bool BeginHook() {
			return true;
		}

		public bool BeginVoicePeak() {
			var p = System.Diagnostics.Process.GetProcesses().Where(x => {
				try {
					return x.MainModule?.ModuleName?.ToLower() == "voicepeak.exe";
				}
				catch(Exception e) when((e is Win32Exception) || (e is InvalidOperationException)) {
					return false;
				}
			}).FirstOrDefault();
			if(p == null) {
				return false;
			}

			this.hVoicePeak = p.MainWindowHandle;
			GetWindowRect(this.hVoicePeak, out var rc);
			SetWindowPos(this.hVoicePeak, IntPtr.Zero, rc.left, rc.top, 1024, 877, 0);
			GetClientRect(this.hVoicePeak, out rc);
			this.voicePeakWidth = rc.right - rc.left;

			if(this.captureProcess == null) {
				var hMapObj = CreateFileMapping(
					(IntPtr)(-1), IntPtr.Zero, PAGE_READWRITE,
					0, 8 * 2,
					MapNameCapture);
				var ptr = MapViewOfFile(hMapObj, FILE_MAP_WRITE, 0, 0, IntPtr.Zero);
				Marshal.WriteIntPtr(ptr, 0, (IntPtr)p.Id);
				Marshal.WriteIntPtr(ptr, 8, this.window.Handle);
				UnmapViewOfFile(ptr);
				this.captureProcess = Process.Start(this.exePath);
			}

			/* 無理ぽいので保留
			this.dwmWindow.Dispatcher.BeginInvoke(() => {
				this.dwmWindow.Show();
				DwmRegisterThumbnail(this.dwmWindow.Handle, this.hVoicePeak, out this.dwmThumb);
				DwmUpdateThumbnailProperties(this.dwmThumb, new DWM_THUMBNAIL_PROPERTIES() {
					dwFlags = DWM_TNP_RECTDESTINATION | DWM_TNP_VISIBLE | DWM_TNP_SOURCECLIENTAREAONLY,
					rcDestination = new RECT() {
						left = 0,
						top = 0,
						right = rc.right - rc.left,
						bottom = rc.bottom - rc.top
					},
					fSourceClientAreaOnly = true,
					fVisible = true
				});
			});
			*/
			return true;
		}

		public void EndCaptureVoicePeak() {
			if(this.captureProcess != null) {
				PostMessage(this.captureWindow, YACM_SHUTDOWN, 0, 0);
				this.captureProcess.Dispose();
				this.captureProcess = null;
			}
		}
		
		public void Speech(string text) {
			static void click(IntPtr hwnd, int x, int y) {
				var pos = x | (y << 16);
				PostMessage(hwnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, (IntPtr)pos);
				PostMessage(hwnd, WM_LBUTTONUP, IntPtr.Zero, (IntPtr)pos);
			}

			static void keyboard(IntPtr hwnd, int keycode) {
				PostMessage(hwnd, WM_KEYDOWN, (IntPtr)keycode, (IntPtr)(0x000000001));
				PostMessage(hwnd, WM_KEYUP, (IntPtr)keycode, (IntPtr)0xC00000001);
			}

			this.sync.Reset();
			PostMessage(
				this.captureWindow,
				YACM_CAPTURE,
				0,
				0);

			var len = text.Length;
			click(this.hVoicePeak, 400, 140);
			foreach(var c in text) {
				SendMessage(this.hVoicePeak, WM_IME_CHAR, (IntPtr)c, IntPtr.Zero);
			}
			Thread.Sleep(50 * text.Length);

			SendMessage(this.hVoicePeak, WM_KEYDOWN, VK_HOME, 0x000000001);
			SendMessage(this.hVoicePeak, WM_KEYUP, VK_HOME, unchecked((int)0xC00000001));
			Thread.Sleep(50);
			// フォーカスを削除してカーソルのWM_PAINTを抑制する
			//SendMessage(this.hVoicePeak, WM_KILLFOCUS, 0, 0);
			click(this.hVoicePeak, this.voicePeakWidth / 2 + 125, 20);
			click(this.hVoicePeak, this.voicePeakWidth / 2 + 165, 20);
			//SendMessage(this.hVoicePeak, WM_IME_CHAR, (IntPtr)VK_SPACE, IntPtr.Zero);

			{
				var waitSec = GetPrivateProfileInt(
					"plugin",
					"waittime_speech",
					this.defaultWaitSec,
					this.iniPath);
				this.sync.WaitOne(waitSec * 1000); // フリーズ防止のためデフォルト50秒で解除する
			}

			click(this.hVoicePeak, 400, 140);
			if(!string.IsNullOrEmpty(text)) {
				// 残ることがあるらしいので3週Deleteを打つ
				for(var i = 0; i < 3; i++) {
					keyboard(this.hVoicePeak, VK_HOME);
					foreach(var _ in text) {
						keyboard(this.hVoicePeak, VK_DELETE);
					}
				}
			}
		}
	}
}