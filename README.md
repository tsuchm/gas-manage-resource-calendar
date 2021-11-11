# Google Apps Script to Manage Calendars of Resources

Google Workspace では，会議室や教室などの共有リソースを Google カレンダーを使って予約できるように[設定する](https://support.google.com/a/answer/1686462)ことが可能です．

組織内で各種共有リソースの予約システムが必要になることは多いので，この機能を使いたいのですが，1つ問題があります．共有リソースを追加する[権限を移譲する](https://workspaceupdates.googleblog.com/2021/02/new-calendar-admin-privilege-hierarchy.html)ことは可能なのですが，追加された共有リソースのアクセス権を設定する権限を移譲する方法が見つかりませんでした．構成員なら誰でも使って良いという共有リソースばかりであれば問題ないのですが，構成員の一部のグループだけが利用できる共有リソースを設定する必要があると，常にドメイン全体の管理者の操作が必要になってしまいます．

https://docs.google.com/spreadsheets/d/17B878jFYrAdxMbcGf4YNSYkKGNZOeRGaXIUuZQRMuds/edit?usp=sharing

