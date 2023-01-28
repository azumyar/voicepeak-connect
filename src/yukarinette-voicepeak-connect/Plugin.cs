using Yukarinette;
using System;
using System.Windows.Forms;
using System.Numerics;
using System.Resources;

namespace Yarukizero.Net.Yularinette.VociePeakConnect {
	public class Plugin : IYukarinetteInterface {
		public override string Name { get; } = "VOICEPEAK メッセージ連携";

		private Connect con = null;
		private System.Media.SoundPlayer player;

		public override void Loaded() {
			if(con == null) {
				con = new Connect();
			}
		}

		public override void SpeechRecognitionStart() {
			try {
				if(!con.BeginSpeech()) {
					throw new YukarinetteException("スピーチ初期化に失敗しました");
				}
				if(!con.BeginHook()) {
					throw new YukarinetteException("コンポーネントが見つからないあるいは不正");
				}
				if(!con.BeginVoicePeak()) {
					if(this.player == null) {
						this.player = new System.Media.SoundPlayer();
					}
					player.Stream = typeof(Plugin).Assembly.GetManifestResourceStream(
						$"{typeof(Plugin).Namespace}.Resources.vp-notfound.wav");
					player.Play();
				}
			}
			catch(YukarinetteException) {
				throw;
			}
			catch(Exception e) {
				throw new YukarinetteException(e);
			}
		}

		public override void Speech(string text) {
			try {
				con.Speech(text);
			}
			catch(Exception e) {
				throw new YukarinetteException(e);
			}
		}

		public override void Closed() {
			this.con?.Dispose();
			this.con = null;
		}
	}
}