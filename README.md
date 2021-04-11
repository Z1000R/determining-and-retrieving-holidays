# VBAによる祝日・休日の判定および取得
<!-- TOC -->

- [VBAによる祝日・休日の判定および取得](#vbaによる祝日休日の判定および取得)
  - [概要](#概要)
  - [応用例](#応用例)
  - [使用方法](#使用方法)
  - [メソッド・プロパティ](#メソッドプロパティ)
    - [祝日定義](#祝日定義)
      - [getNationalHolidayInfoMD](#getnationalholidayinfomd)
      - [getNationalHolidayInfoWN](#getnationalholidayinfown)
    - [休日定義](#休日定義)
      - [getCompanyHolidayInfoW](#getcompanyholidayinfow)
      - [getCompanyHolidayInfoMD](#getcompanyholidayinfomd)
      - [getCompanyHolidayInfoWN](#getcompanyholidayinfown)
      - [getCompanyHolidayInfoMDExclude](#getcompanyholidayinfomdexclude)
      - [getCompanyHolidayInfoWNExclude](#getcompanyholidayinfownexclude)
    - [メソッド](#メソッド)
      - [getCompanyHolidays](#getcompanyholidays)
      - [getCompanyHolidayName](#getcompanyholidayname)
      - [reInitialize](#reinitialize)
    - [データ精製状況](#データ精製状況)
      - [InitializedLastYear](#initializedlastyear)
  - [開発環境](#開発環境)
  - [ライセンス](#ライセンス)
  - [Link](#link)
    - [ブログ](#ブログ)
    - [祝日関連（Wikipedia）](#祝日関連wikipedia)
    - [祝日関連（内閣府）](#祝日関連内閣府)

<!-- /TOC -->

## 概要
- VBAを使用して、クラスモジュール内に祝日情報定義を持つことで、インターネットやサーバー等の外部データを参照せずに、ローカル環境のみで、祝日・休日判定を行うことができる。
- 祝日定義は、法改正により祝日が変わらなければ、原則として変える必要はない。
- 休日定義は、以下の方法で指定できる（適用開始年、適用終了年を含む）
  - 曜日固定
  - 月日固定
  - 月、週、曜日固定
- 休日除外も、以下の方法で指定できる（適用開始年、適用終了年を含む）
  - 月日固定
  - 月、週、曜日固定

## 応用例
本クラスモジュールを呼び出して使用する例として、以下のような事が考えられる。（別途処理が必要）
- 指定日の祝日・休日判定
- 稼働日数のカウント
- 第N営業日の取得
- 休日一覧の作成（Excel WORKDAY関数用などに使用）
- カレンダーの作成

## 使用方法
1. CCompanyHoliday をダウンロードして、使用するOffice製品のVBEで、インポートして下さい。
1. 休日定義の内容を確認して、使用する環境に合わせて設定して下さい。
1. サンプルコードを参考に、目的の処理を実装して下さい。

## メソッド・プロパティ
### 祝日定義
#### getNationalHolidayInfoMD
- 月日固定の祝日情報生成
#### getNationalHolidayInfoWN
- 月週曜日固定の祝日情報生成
### 休日定義
#### getCompanyHolidayInfoW
- 曜日固定の休日情報生成  
#### getCompanyHolidayInfoMD
- 月日固定の休日情報生成  
#### getCompanyHolidayInfoWN
- 月週曜日固定の休日情報生成  
#### getCompanyHolidayInfoMDExclude
- 月日固定の休日除外情報生成  
#### getCompanyHolidayInfoWNExclude
- 月週曜日固定の休日除外情報生成  
### メソッド
#### getCompanyHolidays
- 指定年の祝日を配列に格納して返す  
#### getCompanyHolidayName
- 指定日の祝日名を返す

#### reInitialize
- 指定年までの祝日データを生成させる  

### データ精製状況
#### InitializedLastYear
- 何年までの祝日データが生成されているか

## 開発環境
Windows 10 Home/Pro 64bit  
Microsoft 365 Solo 64bit  
Microsoft Office 2013 32bit  

## ライセンス
Mit ライセンス

## Link
### ブログ
[VBAによる祝日判定および祝日取得](https://z1000s.hatenablog.com/entry/2018/05/28/221451)  
[VBAによる「祝日判定処理」を「休日判定処理」に拡張してみた](https://z1000s.hatenablog.com/entry/2018/09/09/164513)  
[VBAで休日判定処理を使って、指定営業日数後の日付を取得する](https://z1000s.hatenablog.com/entry/2018/09/09/172227)  
[VBAで休日判定処理を使って、Excelワークシートに休日カレンダーを作る](https://z1000s.hatenablog.com/entry/2018/09/09/181213)  

### 祝日関連（Wikipedia）
[国民の祝日](https://ja.wikipedia.org/wiki/%E5%9B%BD%E6%B0%91%E3%81%AE%E7%A5%9D%E6%97%A5)  
[振替休日](https://ja.wikipedia.org/wiki/%E6%8C%AF%E6%9B%BF%E4%BC%91%E6%97%A5)  
[国民の休日](https://ja.wikipedia.org/wiki/%E5%9B%BD%E6%B0%91%E3%81%AE%E4%BC%91%E6%97%A5)  
[国民の祝日に関する法律](https://ja.wikipedia.org/wiki/%E5%9B%BD%E6%B0%91%E3%81%AE%E7%A5%9D%E6%97%A5%E3%81%AB%E9%96%A2%E3%81%99%E3%82%8B%E6%B3%95%E5%BE%8B)  
[春分の日](https://ja.wikipedia.org/wiki/%E6%98%A5%E5%88%86%E3%81%AE%E6%97%A5)  
[秋分の日](https://ja.wikipedia.org/wiki/%E7%A7%8B%E5%88%86%E3%81%AE%E6%97%A5)  

### 祝日関連（内閣府）
[国民の祝日について](https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html)  