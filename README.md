[日本語](README.md)

<h1 align="center">
    <a href="https://github.com/overdrive1708/ExcelFileNumberToName">
        <img alt="ExcelFileNumberToName" src="docs/images/AppIconReadme.png" width="50" height="50">
    </a><br>
    ExcelFileNumberToName
</h1>

<h2 align="center">
    エクセルファイル内の数字を名称に変換するアプリケーション
</h2>

<div align="center">
    <img alt="csharp" src="https://img.shields.io/badge/csharp-blue.svg?style=plastic&logo=csharp">
    <img alt="dotnet" src="https://img.shields.io/badge/.NET-blue.svg?style=plastic&logo=dotnet">
    <img alt="license" src="https://img.shields.io/github/license/overdrive1708/ExcelFileNumberToName?style=plastic">
    <br>
    <img alt="repo size" src="https://img.shields.io/github/repo-size/overdrive1708/ExcelFileNumberToName?style=plastic&logo=github">
    <img alt="release" src="https://img.shields.io/github/release/overdrive1708/ExcelFileNumberToName?style=plastic&logo=github">
    <img alt="download" src="https://img.shields.io/github/downloads/overdrive1708/ExcelFileNumberToName/total?style=plastic&logo=github&color=brightgreen">
    <img alt="open issues" src="https://img.shields.io/github/issues-raw/overdrive1708/ExcelFileNumberToName?style=plastic&logo=github&color=brightgreen">
    <img alt="closed issues" src="https://img.shields.io/github/issues-closed-raw/overdrive1708/ExcelFileNumberToName?style=plastic&logo=github&color=brightgreen">
</div>

## ダウンロード方法

[GitHubのReleases](https://github.com/overdrive1708/ExcelFileNumberToName/releases)にあるLatestのAssetsよりExcelFileNumberToName_vx.x.x.zipをダウンロードしてください｡

## 初回セットアップ

｢ExcelFileNumberToName.exe｣を一回起動して終了してください｡

｢Settings.json｣と｢ConvertRules.csv｣が生成されます｡

[設定例](docs/設定例Settings.json)を参考に｢Settings.json｣を記載してください｡

[設定例](docs/設定例ConvertRules.csv)を参考に｢ConvertRules.csv｣を記載してください｡

｢Settings.json｣で設定する設定項目は以下のとおりです｡
| 設定項目 | 設定内容 |
| --- | --- |
| PresetName | プリセットを識別するための名前を設定します｡ |
| ExaminationFileKeyword | 調査ファイルの一覧に登録する際のキーワードを設定します｡<br>調査ファイルに調査対象フォルダもしくは調査対象ファイルをドラッグ&ドロップした際に､設定したキーワードを含むファイルのみを調査対象とします｡ |
| ExaminationTargets | 調査対象のセルを特定するための情報を設定します｡後述する説明を参照してください｡ |

｢Settings.jsonのExaminationTargets｣で設定する設定項目は以下のとおりです｡
| 設定項目 | 設定内容 |
| --- | --- |
| Sheet | 変換したい数字のセルを特定するためのシート名を設定してください｡ |
| Cell | 変換したい数字のセルを特定するためのセル名を設定してください｡<br>"A1"や"A1:A5"などで指定してください｡ |
| Memo | メモ欄です｡セルの意味などのメモにご利用ください｡ |

｢ConvertRules.csv｣はCSV形式で｢数値,名称｣の形式で変換ルールを規定してください｡

## 使用方法

1. ｢ExcelFileNumberToName.exe｣を起動してください｡
1. 調査設定をプリセットから選択してください｡
1. 調査ファイルに調査対象フォルダもしくは調査対象ファイルをドラッグ&ドロップしてください｡
1. ｢調査実施｣ボタンを押下してください｡

## 開発環境

- Microsoft Visual Studio Community 2022

## 使用しているライブラリ

詳細は[NOTICE.md](NOTICE.md)を参照してください｡

## ライセンス

このプロジェクトはMITライセンスです。  
詳細は [LICENSE](LICENSE) を参照してください。

## 作者

[overdrive1708](https://github.com/overdrive1708)
