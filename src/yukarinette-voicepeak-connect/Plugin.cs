using Yukarinette;
using System;
using System.Windows.Forms;
using System.Numerics;
using System.Resources;
using System.IO;
using System.Windows.Media.Imaging;

namespace Yarukizero.Net.Yularinette.VociePeakConnect {
	public class Plugin : IYukarinetteInterface {
		public override string Name { get; } = "VOICEPEAK メッセージ連携";
		public override System.Windows.Media.ImageSource Icon => icon;

		private Connect con = null;
		private System.Media.SoundPlayer player;
		private System.Windows.Media.ImageSource icon;

		public override void Loaded() {
			var bmp = new System.Windows.Media.Imaging.BitmapImage();
			bmp.BeginInit();
			bmp.CacheOption = BitmapCacheOption.OnLoad;
			bmp.StreamSource = typeof(Plugin).Assembly.GetManifestResourceStream($"{typeof(Plugin).Namespace}.Resources.icon.png");
			bmp.EndInit();
			this.icon = bmp;

			if(con == null) {
				try {
					con = new Connect();
				}
				catch(Exception e) {
					throw new YukarinetteException(e);
				}
			}
		}

		public override void Closed() {
			this.con?.Dispose();
			this.con = null;
		}

		public override void SpeechRecognitionStart() {
			try {
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
	}
}