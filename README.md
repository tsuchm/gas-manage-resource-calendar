# Google Apps Script to Manage Calendars of Resources

Google Workspace では，会議室や教室などの共有リソースを Google カレンダーを使って予約できるように[設定する](https://support.google.com/a/answer/1686462)ことが可能です．

組織内で各種共有リソースの予約システムが必要になることは多いので，この機能を使いたいのですが，1つ問題があります．共有リソースを追加する[権限を移譲する](https://workspaceupdates.googleblog.com/2021/02/new-calendar-admin-privilege-hierarchy.html)ことは可能なのですが，追加された共有リソースのアクセス権を設定する権限を移譲する方法が見つかりませんでした．構成員なら誰でも使って良いという共有リソースばかりであれば問題ないのですが，構成員の一部のグループだけが利用できる共有リソースを設定する必要があると，常にドメイン全体の管理者の操作が必要になってしまいます．組織が大きくなると管理工数が非現実的になってしまいますので，この管理を自動化するスクリプトを作成しました．

## What you can do

[Google スプレッドシートで作成した共有リソースの一覧表](https://docs.google.com/spreadsheets/d/17B878jFYrAdxMbcGf4YNSYkKGNZOeRGaXIUuZQRMuds/edit?usp=sharing)に基づいて，共有リソースを一括作成できます．よって，このスプレッドシートを共有するアカウントを適切に設定することによって，共有リソースを追加・設定する権限を移譲できることになります．

## Usage

(1) [自組織向けの共有リソースの一覧表](https://docs.google.com/spreadsheets/d/17B878jFYrAdxMbcGf4YNSYkKGNZOeRGaXIUuZQRMuds/edit?usp=sharing)を作成．作成したスプレッドシートのID をメモしておく．

(2) ドメイン全体の管理者として [Apps Script](https://script.google.com) にアクセス．新しいプロジェクトを用意して，code.js を配置．

(3) サービスとして，以下の3つを有効化．

 * Admin SDK API service
 * Google Calendar API service
 * Google Sheets API service

(4) code.js 先頭のカスタマーID，一覧表のスプレッドシートのID，ドメイン名を正しく設定する．

(5) 動作確認としては，たとえば showAllResources() を実行してみて，現在の共有リソースが全てきちんと表示されるかを見る．

(6) updateResourcesAndBuildings() を手動実行してみて，一覧表のスプレッドシートの共有リソースが全て一括作成されることを確認する．なお，一覧表に存在していない共有リソースを削除したい場合は，updateResourceAndBuidlings() 内の deleteResources() のコメントアウトを外す．

(7) updateResourcesAndBuildings() を定期的に実行するよう[トリガーを設定する](https://developers.google.com/apps-script/guides/triggers/installable)．

## License

GPL3.
