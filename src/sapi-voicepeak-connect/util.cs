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
