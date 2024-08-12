using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.CoreAudioApi;
using NAudio.Wasapi.CoreAudioApi.Interfaces;
using NAudio.Wave;
using System.Threading;
using System.Drawing;

namespace Yarukizero.Net.Yularinette.VoicePeakConnect;
internal class ApplicationCapture {

	public static class Api {
		public static readonly Guid IID_AudioClient = new("1CB9AD4C-DBFA-4c32-B178-C2F568A703B2");

		[StructLayout(LayoutKind.Sequential)]
		public struct AUDIOCLIENT_ACTIVATION_PARAMS {
			public int ActivationType;
			// AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS 
			public int TargetProcessId;
			public int ProcessLoopbackMode;
		}
		public const int AUDIOCLIENT_ACTIVATION_TYPE_DEFAULT = 0;
		public const int AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK = 1;

		public const int PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE = 0;
		public const int PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE = 1;

		[DllImport("Mmdevapi.dll")]
		public static extern uint ActivateAudioInterfaceAsync(
			[MarshalAs(UnmanagedType.LPWStr)]
		string deviceInterfacePath,
			[MarshalAs(UnmanagedType.LPStruct)]
		Guid riid,
			in PROPVARIANT activationParams,
			IActivateAudioInterfaceCompletionHandler completionHandler,
			out IActivateAudioInterfaceAsyncOperation activationOperation);

		[StructLayout(LayoutKind.Sequential)]
		public struct PROPVARIANT {
			public short vt;
			public short r1;
			public short r2;
			public short r3;
			public int blob_cbSize;
			public nint blob_pBlobData;
			public int dmy;
		}
		public const int VT_BLOB = 65;
		public const string VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK = "VAD\\Process_Loopback";

		public const int AUDCLNT_STREAMFLAGS_CROSSPROCESS = 0x00010000;
		public const int AUDCLNT_STREAMFLAGS_LOOPBACK = 0x00020000;
		public const int AUDCLNT_STREAMFLAGS_EVENTCALLBACK = 0x00040000;
		public const int AUDCLNT_STREAMFLAGS_NOPERSIST = 0x00080000;
		public const int AUDCLNT_STREAMFLAGS_RATEADJUST = 0x00100000;
		public const int AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM = unchecked((int)0x80000000);
		public const int AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY = 0x08000000;

		public const int COINIT_MULTITHREADED = 0x0;
		public const int COINIT_APARTMENTTHREADED = 0x2;
		public const int COINIT_DISABLE_OLE1DDE = 0x4;
		public const int COINIT_SPEED_OVER_MEMORY = 0x8;

		[DllImport("Ole32.dll")]
		public static extern uint CoInitialize(nint pvReserved);
		[DllImport("Ole32.dll")]
		public static extern uint CoInitializeEx(nint pvReserved, int dwCoInit);
		public const uint MF_SDK_VERSION = 0x0002;
		public const uint MF_API_VERSION = 0x0070;
		public const uint MF_VERSION = (MF_SDK_VERSION << 16 | MF_API_VERSION);
		public const int MFSTARTUP_NOSOCKET = 0x1;
		public const int MFSTARTUP_LITE = (MFSTARTUP_NOSOCKET);
		public const int MFSTARTUP_FULL = 0;
		[DllImport("Mfplat.dll")]
		public static extern uint MFStartup(uint Version, int dwFlags);
	}


	class AudioCapture : IActivateAudioInterfaceCompletionHandler {
		private AutoResetEvent initilize_event = new AutoResetEvent(false);
		private IAudioClient? audioClient = default;

		public static AudioClient? Initilize(int pid) {
			AudioCapture result = new();
			Api.AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = new() {
				ActivationType = Api.AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK,
				TargetProcessId = pid,
				ProcessLoopbackMode = Api.PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE,
			};
			var sizeAudioclientActivationParams = Marshal.SizeOf<Api.AUDIOCLIENT_ACTIVATION_PARAMS>();
			var pAudioclientActivationParams = Marshal.AllocHGlobal(sizeAudioclientActivationParams);
			try {
				Marshal.StructureToPtr<Api.AUDIOCLIENT_ACTIVATION_PARAMS>(
					audioclientActivationParams,
					pAudioclientActivationParams,
					true);
				Api.PROPVARIANT activateParams = new();
				activateParams.vt = Api.VT_BLOB;
				activateParams.blob_cbSize = sizeAudioclientActivationParams;
				activateParams.blob_pBlobData = pAudioclientActivationParams;
				var IAudioClient = Api.IID_AudioClient;
				var IID_IAudioClient2 = new Guid("726778CD-F60A-4eda-82DE-E47610CD78AA");
				result.initilize_event.Reset();
				var ret = Api.ActivateAudioInterfaceAsync(
					Api.VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
					IAudioClient,
					activateParams,
					result,
					out var asyncOp);
				System.Diagnostics.Debug.WriteLine($"ActivateAudioInterfaceAsync={ret:x}");
				if(ret == 0) {
					result.initilize_event.WaitOne();
				}
			}
			finally {
				Marshal.FreeHGlobal(pAudioclientActivationParams);
			}
			return new(result.audioClient);
		}

		public void ActivateCompleted(IActivateAudioInterfaceAsyncOperation activateOperation) {
			activateOperation.GetActivateResult(
				out var activateResult,
				out var activatedInterface);
			if((activateResult == 0) && (activatedInterface is IAudioClient client)) {
				this.audioClient = client;
			}
			this.initilize_event.Set();
		}

		public bool IsInitilized => this.audioClient != null;
	}

	private int bytesPerFrame;
	private EventWaitHandle? frameEventWaitHandle;
	private WaveFormat? waveFormat;
	private AudioClient audioClient;
	private CaptureState captureState;
	private Thread? captureThread;
	private DateTime? speakStartTime = default;
	private DateTime? speakEndTime = default;
	private bool isSpeak = false;
	private readonly AutoResetEvent speakWait = new(false);

	private ApplicationCapture(AudioClient audioClient) {
		this.audioClient = audioClient;
	}

	private static bool s_initilized = false;
	private static SynchronizationContext? uiContext;


	public static void UiInitilize() {
		if(s_initilized) {
			return;
		}

		// WPFの場合
		//SynchronizationContext.SetSynchronizationContext(new System.Windows.Threading.DispatcherSynchronizationContext());
		if(SynchronizationContext.Current == null) {
			throw new NullReferenceException();
		}
		uiContext = SynchronizationContext.Current;

		Api.CoInitialize(0);
		Api.MFStartup(Api.MF_VERSION, Api.MFSTARTUP_NOSOCKET);
		s_initilized = true;
	}

	public static async Task<ApplicationCapture?> Get(int pid) {
		var capture = default(ApplicationCapture?);
		var ev = new AutoResetEvent(false);
		ev.Reset();
		var audioClient = AudioCapture.Initilize(pid);
		if(audioClient != null) {
			capture = new ApplicationCapture(audioClient);
			capture.waveFormat = new WaveFormat(
				rate: 16000,
				bits: 16,
				channels: 1);
			//WasapiCapture
			var num = 200000;
			capture.audioClient.Initialize(
				AudioClientShareMode.Shared,
				AudioClientStreamFlags.EventCallback
					| AudioClientStreamFlags.Loopback
					| AudioClientStreamFlags.AutoConvertPcm
				,
				num,
				0,
				capture.waveFormat,
				Guid.Empty);
			capture.frameEventWaitHandle = new EventWaitHandle(
				initialState: false,
				EventResetMode.AutoReset);
			capture.audioClient.SetEventHandle(capture.frameEventWaitHandle.SafeWaitHandle.DangerousGetHandle());
			long bufferSize = unchecked((uint)audioClient.BufferSize);
			capture.bytesPerFrame = capture.waveFormat.Channels * capture.waveFormat.BitsPerSample / 8;
		}
		return capture;
	}

	public void Start() {
		if(captureState != 0) {
			throw new InvalidOperationException("Previous recording still in progress");
		}

		isSpeak = false;
		speakStartTime = speakEndTime = null;
		speakWait.Reset();
		captureState = CaptureState.Starting;
		captureThread = new Thread(() => {
			CaptureThread(audioClient);
		}) {
			IsBackground = true
		};
		captureThread.Start();
	}

	public void Wait() {
		speakWait.WaitOne();
	}

	public void Stop() {
		if(captureState != 0) {
			captureState = CaptureState.Stopping;
		}
	}

	private void CaptureThread(AudioClient client) {
		try {
			DoRecording(client);
		}
		finally {
			client.Stop();
		}
		captureThread = null;
		captureState = CaptureState.Stopped;
	}

	private void DoRecording(AudioClient client) {
		var bufferSize = unchecked((uint)client.BufferSize);
		var num = (long)(1000.0 * (double)bufferSize / (double)waveFormat.SampleRate);
		var timeout = (int)(3 * num / 10000);
		var audioCaptureClient = client.AudioCaptureClient;
		client.Start();
		if(captureState == CaptureState.Starting) {
			captureState = CaptureState.Capturing;
		}

		while(captureState == CaptureState.Capturing) {
			frameEventWaitHandle?.WaitOne(timeout, exitContext: false);

			if(!ReadNextPacket(audioCaptureClient, (rms) => {
				if(rms) {
					if(speakStartTime.HasValue) {
						if(100 < (DateTime.Now - speakStartTime.Value).TotalMilliseconds) {
							return true;
						}
					} else {
						speakStartTime = DateTime.Now;
					}
				} else {
					speakStartTime = null;
				}
				return false;
			})) {
				break;
			}

			if(captureState != CaptureState.Capturing) {
				return;
			}
		}

		while(captureState == CaptureState.Capturing) {
			frameEventWaitHandle?.WaitOne(timeout, exitContext: false);

			if(!ReadNextPacket(audioCaptureClient, (rms) => {
				if(!rms) {
					if(speakEndTime.HasValue) {
						if(750 < (DateTime.Now - speakEndTime.Value).TotalMilliseconds) {
							return true;
						}
					} else {
						speakEndTime = DateTime.Now;
					}
				} else {
					speakEndTime = null;
				}
				return false;
			})) {
				break;
			}

			if(captureState != CaptureState.Capturing) {
				break;
			}
		}
		speakWait.Set();
	}


	private bool ReadNextPacket(AudioCaptureClient capture, Func<bool, bool> speakCheck) {
		int nextPacketSize = capture.GetNextPacketSize();
		while((nextPacketSize != 0) && (captureState == CaptureState.Capturing)) {
			var buffer = capture.GetBuffer(
				out int numFramesToRead,
				out AudioClientBufferFlags bufferFlags);
			var size = numFramesToRead * bytesPerFrame;
			if(this.recodeBytes.Length < size) {
				this.recodeBytes = new byte[size];
			}
			Marshal.Copy(buffer, this.recodeBytes, 0, size);
			capture.ReleaseBuffer(numFramesToRead);
			nextPacketSize = capture.GetNextPacketSize();
			if(speakCheck(Rms(this.recodeBytes.Take(size)))) {
				return false;
			}
		}
		return true;
	}
	private byte[] recodeBytes = Array.Empty<byte>();

	private bool Rms(IEnumerable<byte> buf) {
		// RMS計算
		// sqrt(sum(S_i^2)/n
		long s = 0;
		for(var i = 0; i < buf.Count(); i += 2) {
			long pcm = buf.ElementAt(i) << 8 | buf.ElementAt(i + 1);
			// なんか混じってくるので0xffffは捨てる
			if(pcm != 0xffff) {
				s += pcm * pcm;
			}

		}
		double rms = (0 < s) switch {
			true => Math.Sqrt(s) / (buf.Count() / 2),
			false => 0,
		};
		if(0 < rms) {
			System.Diagnostics.Debug.WriteLine(rms);
		}

		var cf = 100 < rms;
		var now = DateTime.Now;
		return cf;
	}
}

