
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization

Namespace Serialization

	''' <summary>
	''' オブジェクトをXMLファイルにシリアル化又は逆シリアル化するための抽象クラス
	''' </summary>
	''' <remarks>
	''' ファイルを開く又は保存するときのファイル選択ダイアログ画面を表示するメソッドあり
	''' </remarks>
	Public MustInherit Class XMLSerialier
		Inherits DataSerializer

		''' <summary>ファイル選択ダイアログ画面で使用する拡張子フィルター</summary>
		Private _dlgFilter As String
		''' <summary>ファイル選択ダイアログ画面で使用する初期フォルダ</summary>
		Private _dlgInitialDirectory As String

#Region " XmlIgnoreAttribute "

		''' <summary>
		''' ファイル選択ダイアログ画面で使用する拡張子フィルタープロパティ
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		<XmlIgnoreAttribute()> _
		Public Property DlgFilter() As String
			Get
				If _dlgFilter.Length = 0 Then
					_dlgFilter = "XML File(*.xml)|*.xml"
				End If
				Return _dlgFilter
			End Get
			Set(ByVal value As String)
				_dlgFilter = value
			End Set
		End Property

		''' <summary>
		''' ファイル選択ダイアログ画面で使用する初期フォルダプロパティ
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		<XmlIgnoreAttribute()> _
		Public Property DlgInitialDirectory() As String
			Get
				If _dlgInitialDirectory.Length = 0 Then
					_dlgInitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
				End If
				Return _dlgInitialDirectory
			End Get
			Set(ByVal value As String)
				_dlgInitialDirectory = value
			End Set
		End Property

#End Region

#Region " Select File "

		''' <summary>
		''' 保存するファイル名を入力する
		''' </summary>
		''' <param name="title">ダイアログ画面のタイトル</param>
		''' <param name="defaultFilename">デフォルトのファイル名</param>
		''' <returns></returns>
		''' <remarks>
		''' あらかじめフィルターと初期フォルダを指定してください。<br/>
		''' 指定しない場合は、デフォルト値を使用します。<br/>
		''' <list>
		''' <item>
		''' <code>フィルター</code>
		''' <description>XML File(*.xml)|*.xml</description>
		''' </item>
		''' <item>
		''' <code>初期フォルダ</code>
		''' <description>マイ ドキュメント フォルダ</description>
		''' </item>
		''' </list>
		''' </remarks>
		Public Function SelectSaveFile(ByVal title As String, ByVal defaultFilename As String) As String
			dataFilename = String.Empty

			Using dlg As New SaveFileDialog
				dlg.RestoreDirectory = True
				dlg.Title = title
				dlg.Filter = _dlgFilter
				dlg.FileName = defaultFilename
				dlg.InitialDirectory = _dlgInitialDirectory

				' ファイル選択
				If dlg.ShowDialog() = DialogResult.OK Then
					dataFilename = dlg.FileName
				End If

				Return dataFilename
			End Using
		End Function

		''' <summary>
		''' 開くファイル名を入力する
		''' </summary>
		''' <param name="title">ダイアログ画面のタイトル</param>
		''' <param name="defaultFilename">デフォルトのファイル名</param>
		''' <returns></returns>
		''' <remarks>
		''' あらかじめフィルターと初期フォルダを指定してください。<br/>
		''' 指定しない場合は、デフォルト値を使用します。<br/>
		''' <list>
		''' <item>
		''' <code>フィルター</code>
		''' <description>XML File(*.xml)|*.xml</description>
		''' </item>
		''' <item>
		''' <code>初期フォルダ</code>
		''' <description>マイ ドキュメント フォルダ</description>
		''' </item>
		''' </list>
		''' </remarks>
		Public Function SelectOpenFile(ByVal title As String, ByVal defaultFilename As String) As String
			dataFilename = String.Empty

			Using dlg As New OpenFileDialog
				dlg.RestoreDirectory = True
				dlg.Title = title
				dlg.Filter = _dlgFilter
				dlg.FileName = defaultFilename
				dlg.InitialDirectory = _dlgInitialDirectory

				' ファイル選択
				If dlg.ShowDialog() = DialogResult.OK Then
					dataFilename = dlg.FileName
				End If

				Return dataFilename
			End Using
		End Function

#End Region

	End Class

End Namespace
