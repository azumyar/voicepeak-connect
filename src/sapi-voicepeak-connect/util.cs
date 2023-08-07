using Lucene.Net.Analysis.Ja.Dict;
using Lucene.Net.Analysis.Ja.TokenAttributes;
using Lucene.Net.Analysis.Ja;
using Lucene.Net.Analysis.TokenAttributes;
using Lucene.Net.Analysis;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Yarukizero.Net.Sapi.VoicePeakConnect {
	static class English2Kana {
		private const string KuromojiDicFile = "kuromoji.dic";
		private const string KanaDicFile = "english.dic";
		private static UserDictionary? s_userDictionary;
		private static Dictionary<string, string> s_kanaDic = new Dictionary<string, string>();

		public static bool IsInited { get; private set; } = false;

		public static void Init() {
			var dir = Path.GetDirectoryName(typeof(English2Kana).Assembly.Location) ?? "";
			try {
				foreach(var it in File.ReadAllLines((Path.Combine(dir, KanaDicFile)))) {
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
			IsInited = true;
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
						r.Append("ッチ");
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
					if((c1 == 't') && (c2 == 's') && (c3 == 'u')) {
						r.Append("ツ");
						s.Skip(3);
						continue;
					}　else if((c1 == 'k') && (c2 == 'w') && (c3 == 'a')) {
						r.Append("クヮ");
						s.Skip(3);
						continue;
					} else if((c1 == 'g') && (c2 == 'w') && (c3 == 'a')) {
						r.Append("グヮ");
						s.Skip(3);
						continue;
					}


					if((c1 == 'k') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "キャ",
							'i' => "キィ",
							'u' => "キュ",
							'e' => "キェ",
							'o' => "キョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 's') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "シャ",
							'i' => "シィ",
							'u' => "シュ",
							'e' => "シェ",
							'o' => "ショ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'c') && (c2 == 'h') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "チャ",
							'i' => "チ",
							'u' => "チュ",
							'e' => "チェ",
							'o' => "チョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'c') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "チャ",
							'i' => "チィ",
							'u' => "チュ",
							'e' => "チェ",
							'o' => "チョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 't') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "チャ",
							'i' => "チィ",
							'u' => "チュ",
							'e' => "チェ",
							'o' => "チョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'n') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ニャ",
							'i' => "ニィ",
							'u' => "ニュ",
							'e' => "ニェ",
							'o' => "ニョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'h') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ヒャ",
							'i' => "ヒィ",
							'u' => "ヒュ",
							'e' => "ヒェ",
							'o' => "ヒョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'm') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ミャ",
							'i' => "ミィ",
							'u' => "ミュ",
							'e' => "ミェ",
							'o' => "ミョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'r') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "リャ",
							'i' => "リィ",
							'u' => "リュ",
							'e' => "リェ",
							'o' => "リョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'g') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ギャ",
							'i' => "ギィ",
							'u' => "ギュ",
							'e' => "ギェ",
							'o' => "ギョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'd') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ヂャ",
							'i' => "ヂィ",
							'u' => "ヂュ",
							'e' => "ヂェ",
							'o' => "ヂョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'b') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ビャ",
							'i' => "ビィ",
							'u' => "ビュ",
							'e' => "ビェ",
							'o' => "ビョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					} else if((c1 == 'p') && (c2 == 'y') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "ピャ",
							'i' => "ピィ",
							'u' => "ピュ",
							'e' => "ピェ",
							'o' => "ピョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					}


					if((c1 == 's') && (c2 == 'h') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "シャ",
							'i' => "シ",
							'u' => "シュ",
							'e' => "シェ",
							'o' => "ショ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					}else if((c1 == 'c') && (c2 == 'h') && boin(c3)) {
						r.Append(c3 switch {
							'a' => "チャ",
							'i' => "チ",
							'u' => "チュ",
							'e' => "チェ",
							'o' => "チョ",
							_ => throw new InvalidOperationException(),
						});
						s.Skip(3);
						continue;
					}

					if((c1 == 'm') && (c2 == 'm') && boin(c3)) {
						last = c1;
						r.Append("ン");
						s.Skip(1);
						continue;
					} else if((c1 == 'm') && (c2 == 'b') && boin(c3)) {
						last = c1;
						r.Append("ン");
						s.Skip(1);
						continue;
					} else if((c1 == 'm') && (c2 == 'p') && boin(c3)) {
						last = c1;
						r.Append("ン");
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
								'a' => "カ",
								'i' => "キ",
								'u' => "ク",
								'e' => "ケ",
								'o' => "コ",
								_ => throw new InvalidOperationException(),
							},
							's' => c2 switch {
								'a' => "サ",
								'i' => "シ",
								'u' => "ス",
								'e' => "セ",
								'o' => "ソ",
								_ => throw new InvalidOperationException(),
							},
							't' => c2 switch {
								'a' => "タ",
								'i' => "チ",
								'u' => "ツ",
								'e' => "テ",
								'o' => "ト",
								_ => throw new InvalidOperationException(),
							},
							'n' => c2 switch {
								'a' => "ナ",
								'i' => "ニ",
								'u' => "ヌ",
								'e' => "ネ",
								'o' => "ノ",
								_ => throw new InvalidOperationException(),
							},
							'h' => c2 switch {
								'a' => "ハ",
								'i' => "ヒ",
								'u' => "フ",
								'e' => "ヘ",
								'o' => "ホ",
								_ => throw new InvalidOperationException(),
							},
							'm' => c2 switch {
								'a' => "マ",
								'i' => "ミ",
								'u' => "ム",
								'e' => "メ",
								'o' => "モ",
								_ => throw new InvalidOperationException(),
							},
							'y' => c2 switch {
								'a' => "ヤ",
								'i' => "イ",
								'u' => "ユ",
								'e' => "エ",
								'o' => "ヨ",
								_ => throw new InvalidOperationException(),
							},
							'r' => c2 switch {
								'a' => "ラ",
								'i' => "リ",
								'u' => "ル",
								'e' => "レ",
								'o' => "ロ",
								_ => throw new InvalidOperationException(),
							},
							'w' => c2 switch {
								'a' => "ワ",
								'i' => "ウィ",
								'u' => "ウ",
								'e' => "ウェ",
								'o' => "ヲ",
								_ => throw new InvalidOperationException(),
							},
							'g' => c2 switch {
								'a' => "ガ",
								'i' => "ギ",
								'u' => "グ",
								'e' => "ゲ",
								'o' => "ゴ",
								_ => throw new InvalidOperationException(),
							},
							'z' => c2 switch {
								'a' => "ザ",
								'i' => "ジ",
								'u' => "ズ",
								'e' => "ゼ",
								'o' => "ゾ",
								_ => throw new InvalidOperationException(),
							},
							'd' => c2 switch {
								'a' => "ダ",
								'i' => "ヂ",
								'u' => "ヅ",
								'e' => "デ",
								'o' => "ド",
								_ => throw new InvalidOperationException(),
							},
							'b' => c2 switch {
								'a' => "バ",
								'i' => "ビ",
								'u' => "ブ",
								'e' => "ベ",
								'o' => "ボ",
								_ => throw new InvalidOperationException(),
							},
							'p' => c2 switch {
								'a' => "パ",
								'i' => "ピ",
								'u' => "プ",
								'e' => "ペ",
								'o' => "ポ",
								_ => throw new InvalidOperationException(),
							},
							_ => throw new InvalidOperationException(),
						});
						s = s.Skip(2);
						continue;
					} else if((c1 == 'j') && boin(c2)) {
						r.Append(c1 switch {
							'a' => "ジャ",
							'i' => "ジイ",
							'u' => "ジュ",
							'e' => "ジェ",
							'o' => "ジョ",
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
						'a' => "ア",
						'b' => "ベ",
						'c' => "グ",
						'd' => "デ",
						'e' => "エ",
						'f' => "フ",
						'g' => "グ",
						'h' => "チ",
						'i' => "イ",
						'j' => "ジ",
						'k' => "ケ",
						'l' => "ル",
						'm' => "ム",
						'n' => "ン",
						'o' => "オ",
						'p' => "プ",
						'q' => "ク",
						'r' => "ラ",
						's' => "ス",
						't' => "テ",
						'u' => "ウ",
						'v' => "ヴ",
						'w' => "ウ",
						'x' => "",
						'y' => "イ",
						'z' => "ジ",

						'-' => "ー",
						_ => " ", // 区切るためスペースを入れる
					});
				}
			}
			return r.ToString();
		}
	}
}
