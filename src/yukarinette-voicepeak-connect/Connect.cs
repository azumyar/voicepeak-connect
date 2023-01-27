using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace Yarukizero.Net.Yularinette.VociePeakConnect {
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


		private const int WM_LBUTTONDOWN = 0x201;
		private const int WM_LBUTTONUP = 0x202;
		private const int WM_KEYDOWN = 0x0100;
		private const int WM_KEYUP = 0x0101;
		private const int WM_IME_CHAR = 0x286;
		private const int MK_LBUTTON = 0x01;
		private const int VK_HOME = 0x24;
		private const int VK_DELETE = 0x2E;

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
		private const int VPC_HOOK_ENABLE = 1;
		private const int VPC_HOOK_DISABLE = 0;


		class MessageForm : Form {
			private readonly AutoResetEvent sync;

			public MessageForm(AutoResetEvent sync) {
				this.sync = sync;
				this.ShowInTaskbar = false;
				this.Opacity = 0;
			}

			protected override void WndProc(ref Message m) {
				switch(m.Msg) {
				case VPC_MSG_ENDSPEECH:
					this.sync.Reset();
					break;
				}
				base.WndProc(ref m);
			}
		}

		private readonly AutoResetEvent sync = new AutoResetEvent(false);

		private int connectionMsg;
		private MessageForm form;
		private IntPtr formHandle;

		private IntPtr hHookModule;
		private HookVoicePeakProc hookFunc;
		private UnhookVoicePeakProc unhookFunc;

		private IntPtr hVoicePeak;
		private int voicePeakWidth;

		public Connect() {
			this.form = new MessageForm(this.sync);
			this.connectionMsg = RegisterWindowMessage("yarukizero-vp-connect");
		}

		public void Dispose() {
			this.form?.Dispose();
		}

		public bool BeginSpeech() {
			if(!this.form.IsHandleCreated) {
				this.form.Show();
				this.formHandle = this.form.Handle;
			}
			return true;
		}

		public bool BeginHook() {
			if(this.hHookModule == IntPtr.Zero) {
				try {
					this.hHookModule = LoadLibrary(
						Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
						"Plugins",
						"Yarukizero.Net.VoicePeakConnect",
						"voicepeak-hook.dll"));
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
				Thread.Sleep(50);
				PostMessage(hwnd, WM_KEYUP, (IntPtr)keycode, (IntPtr)0xC00000001);
				Thread.Sleep(100);
			}

			click(this.hVoicePeak, 400, 140);
			foreach(var c in text) {
				SendMessage(this.hVoicePeak, WM_IME_CHAR, (IntPtr)c, IntPtr.Zero);
			}
			Thread.Sleep(50 * text.Length);

			this.hookFunc(this.hVoicePeak);
			SendMessage(
				this.hVoicePeak, this.connectionMsg,
				(IntPtr)VPC_MSG_CALLBACKWND,
				this.formHandle);
			SendMessage(
				this.hVoicePeak, this.connectionMsg,
				(IntPtr)VPC_MSG_ENABLEHOOK,
				(IntPtr)VPC_HOOK_ENABLE);
			click(this.hVoicePeak, this.voicePeakWidth / 2 + 165, 20);

			this.sync.WaitOne(5000); // フリーズ防止のため5秒で解除する

			SendMessage(this.hVoicePeak, this.connectionMsg,
				(IntPtr)VPC_MSG_ENABLEHOOK,
				(IntPtr)VPC_HOOK_DISABLE);
			unhookFunc(this.hVoicePeak);

			click(this.hVoicePeak, 400, 140);
			if(!string.IsNullOrEmpty(text)) {
				keyboard(this.hVoicePeak, VK_HOME);
				foreach(var _ in text) {
					keyboard(this.hVoicePeak, VK_DELETE);
				}
			}
		}
	}
}