using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;
using Microsoft.Office.Tools.Excel;

// アセンブリに関する一般情報は以下の属性セットをとおして制御されます。 
// アセンブリに関連付けられている情報を変更するには、
// これらの属性値を変更してください。
[assembly: AssemblyTitle("excelsvn")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("excelsvn")]
[assembly: AssemblyCopyright("Copyright ©  2009")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// ComVisible を false に設定すると、その型はこのアセンブリ内で COM コンポーネントから 
// 参照不可能になります。COM からこのアセンブリ内の型にアクセスする場合は、 
// その型の ComVisible 属性を true に設定してください。
[assembly: ComVisible(false)]

// 次の GUID は、このプロジェクトが COM に公開される場合の、typelib の ID です
[assembly: Guid("46aa7db5-282f-4a86-a5a5-25cfbe7d0e75")]

// アセンブリのバージョン情報は、以下の 4 つの値で構成されています:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// すべての値を指定するか、下のように '*' を使ってビルドおよびリビジョン番号を 
// 既定値にすることができます:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

// 
// ExcelLocale1033 属性は Excel オブジェクト モデルに渡されるロケールを設定
// します。ExcelLocale1033 を true に設定すると、Excel オブジェクト モデルがすべての 
// ロケールで同じ動作をするようになり、Visual Basic for Applications の動作と一致 
// します。ExcelLocale1033 を false に設定すると、ユーザーによってロケール
// 設定が異なる場合、Excel オブジェクト モデルが異なる動作をするようになり、Visual Studio Tools for Office 
// Version 2003 の動作と一致します。これにより、数式名および日付形式などの 
// ロケール情報に、予期しない結果が生じる可能性があります。
// 
[assembly: ExcelLocale1033(true)]

[assembly: SecurityTransparent()]
