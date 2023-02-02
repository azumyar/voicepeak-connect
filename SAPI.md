# voicepeak-connect SAPI

voicepeak-connect のSAPI実装です。  
棒読みちゃんなどのSAPI対応クライアントからVOICEPEAKと連携します  
https://install.appcenter.ms/users/azumyar/apps/sapi-voicepeak/distribution_groups/canary  
バイナリは現時点で棒読みちゃんをターゲットにしているためx86版のみの配布になっています。

## インストール
x64windows-install-x86.batを管理者権限で実行してください。  
VOICEPEAKがC:\Program Files\VOICEPEAK\voicepeak.exeに存在する場合インストールしているVOICEPEAKから話者情報を抜き出し登録します。それ以外の場所にVOICEPEAKをインストールした場合は手動で設定する必要があります。設定ツールは現時点では用意していません。また性別/年齢はCLIからは不明のため固定値で設定されます。  
  
ボイスを追加した場合再度x64windows-install-x86.batを実行することで再初期化されます。

## アンインストール
x64windows-uninstall-x86.batを管理者権限で実行してください。

## 謝辞
コードベースとして[shigobu/SAPIForVOICEVOX](https://github.com/shigobu/SAPIForVOICEVOX)を使用させていただきました。

