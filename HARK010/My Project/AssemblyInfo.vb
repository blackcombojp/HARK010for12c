Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' アセンブリに関する一般情報は以下の属性セットをとおして制御されます。
' アセンブリに関連付けられている情報を変更するには、
' これらの属性値を変更してください。

' アセンブリ属性の値を確認します

<Assembly: AssemblyTitle("汎用データ検索ツール for 物流管理")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyCompany("アイティーアイ株式会社")>
<Assembly: AssemblyProduct("HARK010.exe")>
<Assembly: AssemblyCopyright("2021 ITI Corporation")>
<Assembly: AssemblyTrademark("ITI Corporation")>


<Assembly: ComVisible(False)>

'このプロジェクトが COM に公開される場合、次の GUID が typelib の ID になります
<Assembly: Guid("805cecd3-82c8-4325-b4a9-e3a4d555ef00")>

' アセンブリのバージョン情報は次の 4 つの値で構成されています:
'
'      メジャー バージョン
'      マイナー バージョン
'      ビルド番号
'      Revision
'
' すべての値を指定するか、次を使用してビルド番号とリビジョン番号を既定に設定できます
' 既定値にすることができます:
' <Assembly: AssemblyVersion("1.0.*")>

<Assembly: AssemblyVersion("0.0.3.4")>
<Assembly: AssemblyFileVersion("0.0.3.4")>
<Assembly: log4net.Config.XmlConfigurator(ConfigFile:="log4net.config", Watch:=True)>