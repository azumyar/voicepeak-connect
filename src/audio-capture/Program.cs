using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Yarukizero.Net.Yularinette.VoicePeakConnect;
using System.Reflection;

namespace Yarukizero.Net.Yularinette.AudioCapture;
internal static class Program {
	class MessageForm : Form {
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		public static extern nint PostMessage(nint hwnd, int msg, nint wParam, nint lParam);

		private const int WM_APP = 0x8000;
		private const int YACM_STARTUP = WM_APP + 1;
		private const int YACM_SHUTDOWN = WM_APP + 2;
		private const int YACM_GETSTATE = WM_APP + 4;
		private const int YACM_CAPTURE = WM_APP + 5;

		private const int YACSTATE_NONE = 0;
		private const int YACSTATE_FAIL = 1;
		private const int YACSTATE_INITILIZED = 3;

		private nint reciveWnd;
		private int targetProcess;
		private bool isInit = false;
		private ApplicationCapture? capture;

		public MessageForm(int targetProcess, nint reciveWnd) {
			this.FormBorderStyle = FormBorderStyle.None;
			this.Opacity = 0;
			this.ControlBox = false;

			this.targetProcess = targetProcess;
			this.reciveWnd = reciveWnd;
		}

		protected override void WndProc(ref Message m) {
			switch(m.Msg) {
			case YACM_STARTUP:
				//this.targetProcess = m.WParam.ToInt32();
				//this.reciveWnd = m.LParam;

				break;
			case YACM_SHUTDOWN:
				this.Close();
				break;
			case YACM_GETSTATE:
				if(isInit) {
					m.Result = this.capture switch {
						null => YACSTATE_FAIL,
						_ => YACSTATE_INITILIZED,
					};
				} else {
					m.Result = YACSTATE_NONE;
				}
				break;
			case YACM_CAPTURE:
				if(this.capture != null) {
					var index = m.WParam.ToInt32();
					this.capture.Start();
					Task.Run(() => {
						this.capture.Wait();
						this.capture.Stop();
						PostMessage(reciveWnd, YACM_CAPTURE, index, 0);
					});
				}
				break;
			}

			base.WndProc(ref m);
		}

		protected override async void OnLoad(EventArgs e) {
			base.OnLoad(e);
			ApplicationCapture.UiInitilize();
			PostMessage(reciveWnd, YACM_STARTUP, 0, this.Handle);
			var _ = Task.Run(async () => {
				this.capture = await ApplicationCapture.Get(this.targetProcess);
				this.isInit = true;
			});
			await Task.Delay(500);
			this.Hide();
		}

		protected override void OnFormClosed(FormClosedEventArgs e) {
			base.OnFormClosed(e);
			Application.Exit();
		}
	}
	private const string MapNameCapture = "yarukizero-net-yukarinette.audio-capture";
	[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
	private static extern nint CreateFileMapping(nint hFile, nint lpFileMappingAttributes, int flProtect, int dwMaximumSizeHigh, int dwMaximumSizeLow, string lpName);
	[DllImport("kernel32.dll")]
	private static extern nint MapViewOfFile(nint hFileMappingObject, int dwDesiredAccess, int dwFileOffsetHigh, int dwFileOffsetLow, nint dwNumberOfBytesToMap);
	[DllImport("kernel32.dll")]
	private static extern nint UnmapViewOfFile(nint hFileMappingObject);

	private const int PAGE_READWRITE = 0x04;
	private const int FILE_MAP_READ = 0x00000004;
	private const int ERROR_ALREADY_EXISTS = 183;
	[STAThread]
	static void Main() {
		var hMapObj = CreateFileMapping(
			(IntPtr)(-1), IntPtr.Zero, PAGE_READWRITE,
			0, 8 * 2,
			MapNameCapture);
		if(Marshal.GetLastWin32Error() == ERROR_ALREADY_EXISTS) {
			var ptr = MapViewOfFile(hMapObj, FILE_MAP_READ, 0, 0, IntPtr.Zero);
			var pId = Marshal.ReadInt32(ptr, 0);
			var hwnd = Marshal.ReadIntPtr(ptr, 8);
			UnmapViewOfFile(ptr);

			ApplicationConfiguration.Initialize();
			Application.Run(new MessageForm(pId, hwnd));
		}
	}
}
