using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading.Tasks;
using TTSEngineLib;
using System.Threading;
using System.Xml;

namespace Yarukizero.Net.Sapi.VoicePeakConnect;

internal static class GuidConst {
	public const string InterfaceGuid = "C10F4E0B-52DC-48D0-B44F-6F6EBBC72542";
	public const string ClassGuid = "7F2A34CE-4848-4D60-9D46-921628AED3A3";
}

[ComVisible(true)]
[Guid(GuidConst.InterfaceGuid)]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IVoicePeakConnectTTSEngine : ISpTTSEngine, ISpObjectWithToken { }


[ComVisible(true)]
[Guid(GuidConst.ClassGuid)]
public class VoicePeakConnectTTSEngine : IVoicePeakConnectTTSEngine {
	private const ushort WAVE_FORMAT_PCM = 1;

	private static readonly Guid SPDFID_WaveFormatEx = new Guid("C31ADBAE-527F-4ff5-A230-F62BB61FF70C");
	private static readonly Guid SPDFID_Text = new Guid ("7CEEF9F9-3D13-11d2-9EE7-00C04F797396");

	[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
	private static extern IntPtr CreateFile(string pszFileName, int dwAccess, int dwShare, IntPtr psa, int dwCreatDisposition, int dwFlagsAndAttributes, IntPtr hTemplate);

	[DllImport("kernel32.dll")]
	private static extern bool ReadFile(IntPtr hFile, byte[] pBuffer, int nNumberOfBytesToRead, out int pNumberOfBytesRead, IntPtr pOverlapped);

	[DllImport("kernel32.dll")]
	private static extern int GetFileSize(IntPtr hFile, IntPtr pFileSizeHigh);
	[DllImport("kernel32.dll")]
	private static extern bool CloseHandle(IntPtr hObject);

	[DllImport("kernel32.dll")]
	private static extern int WaitForSingleObject(IntPtr hHandle, int dwMilliseconds);
	private const int WAIT_TIMEOUT = 0x102;
	private const int GENERIC_WRITE = 0x40000000;
	private const int GENERIC_READ = unchecked((int)0x80000000);
	private const int FILE_SHARE_READ = 0x00000001;
	private const int FILE_SHARE_WRITE = 0x00000002;
	private const int FILE_SHARE_DELETE = 0x00000004;
	private const int CREATE_NEW = 1;
	private const int CREATE_ALWAYS = 2;
	private const int OPEN_EXISTING = 3;
	private const int OPEN_ALWAYS = 4;
	private const int TRUNCATE_EXISTING = 5;

	[Flags]
	enum SPVESACTIONS {
		SPVES_CONTINUE = 0,
		SPVES_ABORT = (1 << 0),
		SPVES_SKIP = (1 << 1),
		SPVES_RATE = (1 << 2),
		SPVES_VOLUME = (1 << 3)
	}

	enum SPEVENTENUM {
		SPEI_UNDEFINED = 0,
		SPEI_START_INPUT_STREAM = 1,
		SPEI_END_INPUT_STREAM = 2,
		SPEI_VOICE_CHANGE = 3,
		SPEI_TTS_BOOKMARK = 4,
		SPEI_WORD_BOUNDARY = 5,
		SPEI_PHONEME = 6,
		SPEI_SENTENCE_BOUNDARY = 7,
		SPEI_VISEME = 8,
		SPEI_TTS_AUDIO_LEVEL = 9,
		SPEI_TTS_PRIVATE = 15,
		SPEI_MIN_TTS = 1,
		SPEI_MAX_TTS = 15,
		SPEI_END_SR_STREAM = 34,
		SPEI_SOUND_START = 35,
		SPEI_SOUND_END = 36,
		SPEI_PHRASE_START = 37,
		SPEI_RECOGNITION = 38,
		SPEI_HYPOTHESIS = 39,
		SPEI_SR_BOOKMARK = 40,
		SPEI_PROPERTY_NUM_CHANGE = 41,
		SPEI_PROPERTY_STRING_CHANGE = 42,
		SPEI_FALSE_RECOGNITION = 43,
		SPEI_INTERFERENCE = 44,
		SPEI_REQUEST_UI = 45,
		SPEI_RECO_STATE_CHANGE = 46,
		SPEI_ADAPTATION = 47,
		SPEI_START_SR_STREAM = 48,
		SPEI_RECO_OTHER_CONTEXT = 49,
		SPEI_SR_AUDIO_LEVEL = 50,
		SPEI_SR_RETAINEDAUDIO = 51,
		SPEI_SR_PRIVATE = 52,
		SPEI_ACTIVE_CATEGORY_CHANGED = 53,
		SPEI_RESERVED5 = 54,
		SPEI_RESERVED6 = 55,
		SPEI_MIN_SR = 34,
		SPEI_MAX_SR = 55,
		SPEI_RESERVED1 = 30,
		SPEI_RESERVED2 = 33,
		SPEI_RESERVED3 = 63
	}

	private const ulong SPFEI_FLAGCHECK = (1u << (int)SPEVENTENUM.SPEI_RESERVED1) | (1u << (int)SPEVENTENUM.SPEI_RESERVED2);
	private const ulong SPFEI_ALL_TTS_EVENTS = 0x000000000000FFFEul | SPFEI_FLAGCHECK;
	private const ulong SPFEI_ALL_SR_EVENTS = 0x003FFFFC00000000ul | SPFEI_FLAGCHECK;
	private const ulong SPFEI_ALL_EVENTS = 0xEFFFFFFFFFFFFFFFul;

	private ulong SPFEI(SPEVENTENUM SPEI_ord) => (1ul << (int)SPEI_ord) | SPFEI_FLAGCHECK;

	enum SPEVENTLPARAMTYPE {
		SPET_LPARAM_IS_UNDEFINED = 0,
		SPET_LPARAM_IS_TOKEN = (SPET_LPARAM_IS_UNDEFINED + 1),
		SPET_LPARAM_IS_OBJECT = (SPET_LPARAM_IS_TOKEN + 1),
		SPET_LPARAM_IS_POINTER = (SPET_LPARAM_IS_OBJECT + 1),
		SPET_LPARAM_IS_STRING = (SPET_LPARAM_IS_POINTER + 1)
	}

	private static readonly string DefaultVoicePeakPath = @"C:\Program Files\VOICEPEAK\voicepeak.exe";
	private static readonly string KeyVoicePeakPath = "x-voicepeak-path";
	private static readonly string KeyVoicePeakNarrator = "x-voicepeak-narrator";
	private static readonly string KeyVoicePeakPitch = "x-voicepeak-pitch";
	private static readonly string KeyVoicePeakEmotion = "x-voicepeak-emotion";

	private ISpObjectToken? token;
	private string? voicePeakPath = null;
	private string? voicePeakNarrator = null;
	private string? voicePeakEmotion = null;
	private string? voicePeakPitch = null;
	private System.Media.SoundPlayer? player = null;

	public void Speak(uint dwSpeakFlags, ref Guid rguidFormatId, ref WAVEFORMATEX pWaveFormatEx, ref SPVTEXTFRAG pTextFragList, ISpTTSEngineSite pOutputSite) {
		static bool deleteTmp(string tmpWaveFile) {
			try {
				if(File.Exists(tmpWaveFile)) {
					File.Delete(tmpWaveFile);
				}
				return true;
			}
			catch(IOException) {
				return false;
			}
		}
		static uint output(ISpTTSEngineSite output, byte[] data) {
			var pWavData = IntPtr.Zero;
			try {
				if(data.Length == 0) {
					output.Write(pWavData, 0u, out var written);
					return written;
				} else {
					pWavData = Marshal.AllocCoTaskMem(data.Length);
					Marshal.Copy(data, 0, pWavData, data.Length);
					output.Write(pWavData, (uint)data.Length, out var written);
					return written;
				}
			}
			finally {
				if(pWavData != IntPtr.Zero) {
					Marshal.FreeCoTaskMem(pWavData);
				}
			}
		}
		void play(string resourceName) {
			if(this.player == null) {
				this.player = new System.Media.SoundPlayer();
			}
			player.Stream = typeof(VoicePeakConnectTTSEngine)
				.Assembly
				.GetManifestResourceStream(resourceName);
			player.Play();
		}

		if(rguidFormatId == SPDFID_Text) {
			return;
		}

		var optSpeed = $"";
		{
			pOutputSite.GetRate(out var spd);
			var n = Math.Max(Math.Min(1d, spd / 10d), -1d);
			optSpeed = n switch {
				var v when (0 < n) => $" --speed {100 + (int)(n * 100)} ",
				var v when (n < 0) => $" --speed {100 - (int)(n * 50)} ",
				_ => ""
			};
		}
		var voicePeakExe = this.voicePeakPath ?? DefaultVoicePeakPath;
		var optNarrator = this.voicePeakNarrator switch {
			string s when !string.IsNullOrEmpty(s) => $" -n \"{s}\" ",
			_ => ""
		};
		var optEmotion = this.voicePeakEmotion switch {
			string s when !string.IsNullOrEmpty(s) => $" -e {s} ",
			_ => ""
		};
		var optPitch = this.voicePeakPitch switch {
			string s when s == "0" =>  "",
			string s when !string.IsNullOrEmpty(s) =>  $" --pitch {s} ",
			_ => ""
		};

		if(!File.Exists(voicePeakExe)) {
			play($"{typeof(VoicePeakConnectTTSEngine).Namespace}.Resources.vp-notfound.wav");
		}

		try {
			var writtenWavLength = 0UL;
			var currentTextList = pTextFragList;
			while(true) {
				if(currentTextList.State.eAction == SPVACTIONS.SPVA_ParseUnknownTag) {
					goto next;
				}
				var text = Regex.Replace(
					currentTextList.pTextStart,
					@"<.+?>",
					"",
					RegexOptions.IgnoreCase);
				if(string.IsNullOrWhiteSpace(text)) {
					goto next;
				}
				if(((SPVESACTIONS)pOutputSite.GetActions()).HasFlag(SPVESACTIONS.SPVES_ABORT)) {
					return;
				}
				AddEventToSAPI(pOutputSite, currentTextList.pTextStart, text, writtenWavLength);

				if(!File.Exists(voicePeakExe)) {
					// 無音をしゃべらせる
					writtenWavLength += output(pOutputSite, new byte[4]);
					goto next;
				}

				var tmpWaveFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.wav");
				if(!deleteTmp(tmpWaveFile)) {
					goto next;
				}
				// 先にファイルを開いておくことでVoicePeakのロックを回避する
				var hFile = CreateFile(
					tmpWaveFile,
					GENERIC_READ,
					FILE_SHARE_READ | FILE_SHARE_WRITE,
					IntPtr.Zero,
					CREATE_NEW,
					0,
					IntPtr.Zero);
				var ms = new MemoryStream();
				var p = Process.Start(new ProcessStartInfo() {
					FileName = voicePeakExe,
					Arguments = $"-s \"{text}\"{optNarrator} -o {tmpWaveFile}{optSpeed}{optEmotion}{optPitch}",
				});
				if(p == null) {
					play($"{typeof(VoicePeakConnectTTSEngine).Namespace}.Resources.vp-error.wav");

					// 無音をしゃべらせる
					writtenWavLength += output(pOutputSite, new byte[4]);
					goto next;
				}

				try {
					var t = Task.Run(() => {
						byte[] b = new byte[76800]; // 1秒間のデータサイズ
						var pos = 0;
						var len = 0;
						while(ReadFile(
							hFile,
							b, b.Length,
							out var ret,
							IntPtr.Zero)) {

							if(0 < ret) {
								if(pos == 0) {
									var head = 104; // voicepeakが出力するデータ領域開始アドレス
									ms.Write(b, head, ret - head);
								} else {
									ms.Write(b, 0, ret);
								}
								pos += ret;
							}
							var exit = WaitForSingleObject(p.Handle, 1);
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
								break;
							}
						}
					});
					p.WaitForExit();
					t.Wait();

					if(p.ExitCode == 0) {
						writtenWavLength += output(pOutputSite, ms.ToArray());
					} else {
						play($"{typeof(VoicePeakConnectTTSEngine).Namespace}.Resources.vp-error.wav");

						// 無音をしゃべらせる
						writtenWavLength += output(pOutputSite, new byte[4]);
					}
				}
				finally {
					if(hFile != IntPtr.Zero) {
						CloseHandle(hFile);
					}
					ms.DisposeAsync();
				}
				deleteTmp(tmpWaveFile);

			next:
				if(currentTextList.pNext == IntPtr.Zero) {
					break;
				} else {
					currentTextList = Marshal.PtrToStructure<SPVTEXTFRAG>(currentTextList.pNext);
				}
			}
		}
		catch {
			play($"{typeof(VoicePeakConnectTTSEngine).Namespace}.Resources.unknown-error.wav");
			throw;
		}
	}

	private void AddEventToSAPI(ISpTTSEngineSite outputSite, string allText, string speakTargetText, ulong writtenWavLength) {
		outputSite.GetEventInterest(out var ev);
		var list = new List<SPEVENT>();
		var wParam = (uint)speakTargetText.Length;
		var lParam = allText.IndexOf(speakTargetText);
		if((ev & SPFEI(SPEVENTENUM.SPEI_SENTENCE_BOUNDARY)) == SPFEI(SPEVENTENUM.SPEI_SENTENCE_BOUNDARY)) {
			list.Add(new SPEVENT() { 
				eEventId = (ushort)SPEVENTENUM.SPEI_SENTENCE_BOUNDARY,
				elParamType = (ushort)SPEVENTLPARAMTYPE.SPET_LPARAM_IS_UNDEFINED,
				wParam = wParam,
				lParam = lParam,
				ullAudioStreamOffset = writtenWavLength
			});
		}
		if((ev & SPFEI(SPEVENTENUM.SPEI_WORD_BOUNDARY)) == SPFEI(SPEVENTENUM.SPEI_WORD_BOUNDARY)) {
			list.Add(new SPEVENT() {
				eEventId = (ushort)SPEVENTENUM.SPEI_WORD_BOUNDARY,
				elParamType = (ushort)SPEVENTLPARAMTYPE.SPET_LPARAM_IS_UNDEFINED,
				wParam = wParam,
				lParam = lParam,
				ullAudioStreamOffset = writtenWavLength
			});
		}
		if(list.Any()) {
			var arr = list.ToArray();
			outputSite.AddEvents(ref arr[0], (uint)arr.Length);
		}
	}

	public void GetOutputFormat(ref Guid pTargetFmtId, ref WAVEFORMATEX pTargetWaveFormatEx, out Guid pOutputFormatId, IntPtr ppCoMemOutputWaveFormatEx) {
		pOutputFormatId = SPDFID_WaveFormatEx;
		var wf = new WAVEFORMATEX() {
			wFormatTag = WAVE_FORMAT_PCM,
			nChannels = 1,
			cbSize = 0,
			nSamplesPerSec = 48000,
			wBitsPerSample = 16,
			nBlockAlign = 1 * 16 / 8, // チャンネル * bps / 8
			nAvgBytesPerSec = 48000 * (1 * 16 / 8), // サンプリングレート / ブロックアライン
		};

		var p = Marshal.AllocCoTaskMem(Marshal.SizeOf(wf));
		Marshal.StructureToPtr(wf, p, false);
		Marshal.WriteIntPtr(ppCoMemOutputWaveFormatEx, p);
	}

	public void SetObjectToken(ISpObjectToken pToken) {
		string get(string key) {
			pToken.GetStringValue(key, out var s);
			return s;
		}

		this.token = pToken;
		this.voicePeakPath = get(KeyVoicePeakPath);
		this.voicePeakNarrator = get(KeyVoicePeakNarrator);
		this.voicePeakEmotion = get(KeyVoicePeakEmotion);
		this.voicePeakPitch = get(KeyVoicePeakPitch);
	}


	public void GetObjectToken(ref ISpObjectToken? ppToken) {
		ppToken = token;
	}

	[ComRegisterFunction()]
	public static void RegisterClass(string _) {
		static string safePath(string name) => Regex.Replace(name, @"[\s,/\:\*\?""\<\>\|]", "");
		var entry = @"SOFTWARE\Microsoft\Speech\Voices\Tokens";
		var prefix = "TTS_YARUKIZERO_VOICEPAEK";
		var narrators = "";

		if(!File.Exists(DefaultVoicePeakPath)) {
			return;
		}

		try {
			var p = Process.Start(new ProcessStartInfo() {
				FileName = DefaultVoicePeakPath,
				Arguments = $"--list-narrator",
				RedirectStandardOutput = true,
			});
			if(p != null) {
				p.WaitForExit();
				narrators = p.StandardOutput.ReadToEnd();
			}
		}
		catch {
			return;
		}

		// 一度情報を破棄する
		InitializeRegistry();
		foreach(var name in narrators.Replace("\r\n", "\n").Split('\n').Where(x => !string.IsNullOrWhiteSpace(x))) {
			using(var registryKey = Registry.LocalMachine.CreateSubKey($@"{entry}\{prefix}-{safePath(name)}")) {
				registryKey.SetValue("", $"VOICEPAEK {name}");
				registryKey.SetValue("411", $"VOICEPAEK {name}");
				registryKey.SetValue("CLSID", $"{{{GuidConst.ClassGuid}}}");
				registryKey.SetValue(KeyVoicePeakPath, DefaultVoicePeakPath);
				registryKey.SetValue(KeyVoicePeakNarrator, name);
				registryKey.SetValue(KeyVoicePeakPitch, "0");
				registryKey.SetValue(KeyVoicePeakEmotion, "");
			}
			using(var registryKey = Registry.LocalMachine.CreateSubKey($@"{entry}\{prefix}-{safePath(name)}\Attributes")) {
				registryKey.SetValue("Age", "Teen"); // とれないのでここはてきとー
				registryKey.SetValue("Vendor", "AHS");
				registryKey.SetValue("Language", "411");
				registryKey.SetValue("Gender", "Female"); // とれないのでここもてきとー
				registryKey.SetValue("Name", $"VOICEPAEK {name}");
			}
		}
	}

	[ComUnregisterFunction()]
	public static void UnregisterClass(string _) {
		InitializeRegistry();
	}

	private static void InitializeRegistry() {
		using(var regTokensKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Speech\Voices\Tokens\", true)) {
			if(regTokensKey == null) {
				return;
			}
			foreach(var name in regTokensKey.GetSubKeyNames()) {
				using(var regKey = regTokensKey.OpenSubKey(name)) {
					if(regKey?.GetValue("CLSID") is string id && id == $"{{{GuidConst.ClassGuid}}}") {
						regTokensKey.DeleteSubKeyTree(name);
					}
				}
			}
		}
	}
}
