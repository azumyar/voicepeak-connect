using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Yarukizero.Net.Yularinette.VoicePeakConnect {
	internal class Connect : IDisposable {
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		static extern int RegisterWindowMessage(string lpString);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern IntPtr PostMessage(IntPtr hwnd, int msg, IntPtr wp, IntPtr lp);
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wp, IntPtr lp);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern IntPtr LoadLibrary(string lpLibFileName);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
		static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern IntPtr FindWindow(string pClassName, string pWindowName);

		[DllImport("user32.dll")]
		private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
		[DllImport("user32.dll")]
		private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
		[DllImport("user32.dll")]
		private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern int GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern IntPtr CreateFileMapping(IntPtr hFile, IntPtr lpFileMappingAttributes, int flProtect, int dwMaximumSizeHigh, int dwMaximumSizeLow, string lpName);
		[DllImport("kernel32.dll")]
		private static extern IntPtr MapViewOfFile(IntPtr hFileMappingObject, int dwDesiredAccess, int dwFileOffsetHigh, int dwFileOffsetLow, IntPtr dwNumberOfBytesToMap);
		[DllImport("kernel32.dll")]
		private static extern IntPtr UnmapViewOfFile(IntPtr hFileMappingObject);
		[DllImport("kernel32.dll")]
		private static extern IntPtr CloseHandle(IntPtr hObject);
		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern IntPtr lstrcpy(IntPtr str1, string str2);


		private const int WM_LBUTTONDOWN = 0x201;
		private const int WM_LBUTTONUP = 0x202;
		private const int WM_KEYDOWN = 0x0100;
		private const int WM_KEYUP = 0x0101;
		private const int WM_IME_CHAR = 0x286;
		private const int MK_LBUTTON = 0x01;
		private const int VK_HOME = 0x24;
		private const int VK_DELETE = 0x2E;

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
				if(msg == VPCM_ENDSPEECH) {
					this.con.sync.Set();
					handled = true;
				}
				return IntPtr.Zero;
			}
		}

		private readonly AutoResetEvent sync = new AutoResetEvent(false);

		private int connectionMsg;
		private MessageWindow window;
		private IntPtr formHandle;

		private IntPtr hHookModule;
		private HookVoicePeakProc hookFunc;
		private UnhookVoicePeakProc unhookFunc;

		private IntPtr hVoicePeak;
		private int voicePeakWidth;

		private readonly string dllPath;
		private readonly string iniPath;
		private readonly int defaultWaitSec = 50;

		public Connect() {
			this.connectionMsg = RegisterWindowMessage(VpConnectMessage);
			this.dllPath = Path.Combine(
				AppDomain.CurrentDomain.BaseDirectory,
				"Plugins",
				"Yarukizero.Net.VoicePeakConnect",
				"voicepeak-hook.dll");
			this.iniPath = Path.Combine(
				AppDomain.CurrentDomain.BaseDirectory,
				"Plugins",
				"Yarukizero.Net.VoicePeakConnect",
				"voicepeak-connet.ini");

			this.window = new MessageWindow(this);
			this.window.Show();
			this.formHandle = this.window.Handle;
		}

		public void Dispose() {
			this.window?.Close();
		}

		public bool BeginHook() {
			if(this.hHookModule == IntPtr.Zero) {
				try {
					this.hHookModule = LoadLibrary(this.dllPath);
					this.hookFunc = Marshal.GetDelegateForFunctionPointer<HookVoicePeakProc>(GetProcAddress(this.hHookModule, "HookVoicePeak"));
					this.unhookFunc = Marshal.GetDelegateForFunctionPointer<UnhookVoicePeakProc>(GetProcAddress(this.hHookModule, "UnhookVoicePeak"));
				}
				catch(ArgumentException _) { // Marshal.GetDelegateForFunctionPointerが投げる
					return false;
				}
			}
			return (this.hHookModule != IntPtr.Zero)
				&& (this.hookFunc != null)
				&& (this.unhookFunc != null);
		}


		public bool BeginVoicePeak() {
			this.hVoicePeak = FindWindow(null, "VOICEPEAK");
			if(this.hVoicePeak == IntPtr.Zero) {
				return false;
			}

			GetWindowRect(this.hVoicePeak, out var rc);
			SetWindowPos(this.hVoicePeak, IntPtr.Zero, rc.left, rc.top, 1024, 877, 0);
			GetClientRect(this.hVoicePeak, out rc);
			this.voicePeakWidth = rc.right - rc.left;

			return true;
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

			this.hookFunc(this.hVoicePeak);
			PostMessage(
				this.hVoicePeak, this.connectionMsg,
				(IntPtr)VPC_MSG_CALLBACKWND,
				this.formHandle);

			var hMapObj = CreateFileMapping(
				(IntPtr)(-1), IntPtr.Zero, PAGE_READWRITE,
				0, text.Length * 2 + 2,
				MapNameHokk2);
			var ptr = MapViewOfFile(hMapObj, FILE_MAP_WRITE, 0, 0, IntPtr.Zero);
			lstrcpy(ptr, text);
			UnmapViewOfFile(ptr);
			PostMessage(
				this.hVoicePeak, this.connectionMsg,
				(IntPtr)VPC_MSG_ENABLEHOOK2,
				(IntPtr)VPC_HOOK_ENABLE);

			{
				var waitSec = GetPrivateProfileInt(
					"plugin",
					"waittime",
					this.defaultWaitSec,
					this.iniPath);
				this.sync.WaitOne(waitSec * 1000); // フリーズ防止のためデフォルト50秒で解除する
			}

			PostMessage(this.hVoicePeak, this.connectionMsg,
				(IntPtr)VPC_MSG_ENABLEHOOK2,
				(IntPtr)VPC_HOOK_DISABLE);
			this.unhookFunc(this.hVoicePeak);
			CloseHandle(hMapObj);

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