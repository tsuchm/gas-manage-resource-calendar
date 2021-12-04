# Google Apps Script to Manage Calendars of Resources

この記事は[Google Apps Script Advent Calendar 2021](https://qiita.com/advent-calendar/2021/google-apps-script)の7日目の記事です

Google Workspace では，会議室や教室などの共有リソースを Google カレンダーを使って予約できるように[設定する](https://support.google.com/a/answer/1686462)ことが可能です．

組織内の共有リソースの予約システムが必要になることは多いので，この機能を利用しようとしたのですが，1つ問題が発生しました．共有リソースの[追加権限を移譲する](https://workspaceupdates.googleblog.com/2021/02/new-calendar-admin-privilege-hierarchy.html)ことは可能ですし，また共有リソースを[一括登録する](https://support.google.com/a/answer/1033925)ことも可能なので，構成員なら誰でも利用できる共有リソースなら問題なく利用できそうです．しかし、筆者の組織では、一部の構成員しか利用できない共有リソース（例えば、A学科の教員・学生だけが予約できるセミナー室など）が大半であり，共有リソース毎にアクセス権を設定する必要があります．その上，筆者の知る限り，アクセス権を設定する権限を移譲する方法が提供されていないため，組織が大きくなると管理工数が非現実的になってしまいます．このスクリプトは，この管理作業を自動化します．

## What you can do

Google スプレッドシートで作成した共有リソースの一覧表（[サンプル](https://docs.google.com/spreadsheets/d/17B878jFYrAdxMbcGf4YNSYkKGNZOeRGaXIUuZQRMuds/edit?usp=sharing)）に基づいて，共有リソースを一括作成できます．よって，このスプレッドシートを共有するアカウントを適切に設定することによって，共有リソースを追加・設定する権限を移譲できることになります．

## Usage

(1) 自組織向けの共有リソースの一覧表（[サンプル](https://docs.google.com/spreadsheets/d/17B878jFYrAdxMbcGf4YNSYkKGNZOeRGaXIUuZQRMuds/edit?usp=sharing)）を作成する．

Buildings シートが建物の一覧表であり，一括登録用のCSVファイルの[フォーマット](https://support.google.com/a/answer/1033925#calendar-format)の抜粋である．
作成済みの建物の一覧表は[管理コンソール > ディレクトリ > ビルディングとリソース > リソース管理](https://admin.google.com/ac/calendarresources/buildings)から一括ダウンロードできるので，ダウンロードしたCSVファイルを Buildings シートに貼り付けると，現状の設定内容そのままの一覧表を作成できる．

Resources シートが共有リソースの一覧表である．一括登録用のCSVファイルの[フォーマット](https://support.google.com/a/answer/1033925#calendar-format)と基本的な部分は同じであるが，幾つかの相違点がある．

第1に，`Resource User`列が追加されている．`Resource User`列には，当該共有リソースを利用できる構成員グループの名前を指定する．構成員グループの名前とグループアドレスの対応関係は ResourceUserGroups シートに登録しておく．

第2に，`Building Id`列に代えて，`Building Name`列がある．建物IDと建物名の対応関係は，Buildings シートに登録しておく．

第3に，`Resource Category`列は日本語化されている．`Resource Category`の英名と和名の対応関係は，ResourceCategories シートに登録しておく．

第4に，`Resource Calendar`列が追加されている．この列は，当該共有リソースのカレンダーに対するURLである．スクリプトにより更新されるので，空欄で良い．

`Resource Id`列は，当該共有リソースを識別するためのIDである．新規に共有リソースを追加する場合は空欄にしておく．
作成済みの共有リソースの一覧表は[管理コンソール > ディレクトリ > ビルディングとリソース > リソース管理](https://admin.google.com/ac/calendarresources/resources)から一括ダウンロードできるので，ダウンロードしたCSVファイルを Resources シートに内容を調整しながら貼り付けると，現状の設定内容そのままの一覧表を作成できる．

(2) ドメイン全体の管理者として [Apps Script](https://script.google.com) にアクセス．新しいプロジェクトを用意して code.js を配置．

(3) サービスとして，以下の4つを有効化．

 * Admin SDK API service
 * Google Calendar API service
 * Google Sheets API service
 * Gmail API service

(4) code.js 先頭のカスタマーID，一覧表のスプレッドシートID，ドメイン名を正しく設定する．

(5) 動作確認としては，たとえば showAllResources() を実行してみて，現在の共有リソースが全てきちんと表示されるかを見る．

(6) updateResourcesAndBuildings() を手動実行してみて，一覧表のスプレッドシートの共有リソースが全て一括作成されることを確認する．なお，一覧表に存在していない共有リソースを削除したい場合は，updateResourcesAndBuidlings() 内の deleteResources() のコメントアウトを外す．

(7) updateResourcesAndBuildings() を定期的に実行するよう[トリガーを設定する](https://developers.google.com/apps-script/guides/triggers/installable)．

## License

GPL3.
