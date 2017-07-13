Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' アセンブリに関する一般情報は以下の属性セットをとおして制御されます。 
' アセンブリに関連付けられている情報を変更するには、
' これらの属性値を変更してください。

' アセンブリ属性の値を確認します。

<Assembly: AssemblyDescription("MocaWin")>
<Assembly: AssemblyCompany("MiYABiS")> 
<Assembly: AssemblyProduct("MocaWin")> 
<Assembly: AssemblyCopyright("Copyright © 2011 MiYABiS All Rights Reserved.")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(True)> 

'このプロジェクトが COM に公開される場合、次の GUID がタイプ ライブラリの ID になります。
<Assembly: Guid("c97df286-05e9-4bcd-9e04-dded13f5a719")>

' アセンブリのバージョン情報は、以下の 4 つの値で構成されています:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' すべての値を指定するか、下のように '*' を使ってビルドおよびリビジョン番号を 
' 既定値にすることができます:


' プログラム要素が CLS (Common Language Specification) に準拠しているかどうかを示します
<Assembly: System.CLSCompliant(True)>

#If net20 Then
<Assembly: AssemblyVersion("2.0.0")>
<Assembly: AssemblyFileVersion("2.0.0")>
<Assembly: AssemblyTitle("Moca.NET Win .NET 2.0")>
<Assembly: AssemblyInformationalVersion("2.0.0 .NET 2.0")>
#End If
#If net35 Then
<Assembly: AssemblyVersion("3.5.0")> 
<Assembly: AssemblyFileVersion("3.5.0")> 
<Assembly: AssemblyTitle("Moca.NET Win .NET 3.5")>
<Assembly: AssemblyInformationalVersion("3.5.0 .NET 3.5")>
#End If
#If net40 Then
<Assembly: AssemblyVersion("4.0.0")>
<Assembly: AssemblyFileVersion("4.0.0")>
<Assembly: AssemblyTitle("Moca.NET Win .NET 4.0")>
<Assembly: AssemblyInformationalVersion("4.0.0 .NET 4.0")>
#End If
#If net45 Then
<Assembly: AssemblyVersion("4.5.0")> 
<Assembly: AssemblyFileVersion("4.5.0")> 
<Assembly: AssemblyTitle("Moca.NET Win .NET 4.5")>
<Assembly: AssemblyInformationalVersion("4.5.0 .NET 4.5")>
#End If
#If net452 Then
<Assembly: AssemblyVersion("4.5.2")> 
<Assembly: AssemblyFileVersion("4.5.2")> 
<Assembly: AssemblyTitle("Moca.NET Win .NET 4.5.2")>
<Assembly: AssemblyInformationalVersion("4.5.2 .NET 4.5.2")>
#End If
#If net46 Then
<Assembly: AssemblyVersion("4.6.0")>
<Assembly: AssemblyFileVersion("4.6.0")>
<Assembly: AssemblyTitle("Moca.NET Win .NET 4.6")>
<Assembly: AssemblyInformationalVersion("4.6.0 .NET 4.6")>
#End If
#If net462 Then
<Assembly: AssemblyVersion("4.6.2")>
<Assembly: AssemblyFileVersion("4.6.2")>
<Assembly: AssemblyTitle("Moca.NET Win .NET 4.6.2")>
<Assembly: AssemblyInformationalVersion("4.6.2 .NET 4.6.2")>
#End If
#If net47 Then
<Assembly: AssemblyVersion("4.7.0")>
<Assembly: AssemblyFileVersion("4.7.0")>
<Assembly: AssemblyTitle("Moca.NET Win .NET 4.7")>
<Assembly: AssemblyInformationalVersion("4.7.0 .NET 4.7")>
#End If
