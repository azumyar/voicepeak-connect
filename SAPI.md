# voicepeak-connect SAPI

AHSサポートから**VOICEPEAK 1.2.1のCLI機能はテスト機能のままででAHSとしては本バージョンでサポートしていない**という回答をもらったので現在連続して発声するとボイスピークが不正終了することがある事象は対応できません。今後に期待。  
  
普通のSAPI実装です。普通にCLIで連携するので早くありません。  
棒読みちゃんなどのSAPI対応クライアントからVOICEPEAKと連携します。  
https://install.appcenter.ms/users/azumyar/apps/sapi-voicepeak/distribution_groups/canary  

## インストール/アンインストール
### 32bit版(棒読みちゃんを使う場合はこちら)
インストールはx64windows-install-x86.batを管理者権限で実行してください。   
アンインストールはx64windows-uninstall-x86.batを管理者権限で実行してください。  
  
 棒読みちゃんを使う場合かつ棒読みちゃんの機能を使用しない場合**x64windows-install-as-x86.bat**でインストールすると単体で喋るように構成します。長文の時に発声速度の向上が得られますが、残響や音量など棒読みちゃんが管理している機能は使えません。
 
### 64bit版
インストールはx64windows-install-x64.batを管理者権限で実行してください。   
アンインストールはx64windows-uninstall-x64.batを管理者権限で実行してください。

### 共通の仕様
.NET6のランタイムをインストールしてください。  
  
VOICEPEAKがC:\Program Files\VOICEPEAK\voicepeak.exeに存在する場合インストールしているVOICEPEAKから話者情報を抜き出し登録します。それ以外の場所にVOICEPEAKをインストールした場合は手動で設定する必要があります。設定ツールは現時点では用意していません。また性別/年齢はCLIからは不明のため固定値で設定されます。  
  
ボイスを追加した場合再度インストールを実行することで再初期化されます。

## 謝辞
コードベースとして[shigobu/SAPIForVOICEVOX](https://github.com/shigobu/SAPIForVOICEVOX)を使用させていただきました。

