
Imports System.Threading

Namespace Win

    ''' <summary>
    ''' 画面ハードコピー
    ''' </summary>
    ''' <remarks>
    ''' 画面全体のハードコピー又は、特定のWindowのハードコピーを取得します。<br/>
    ''' 取得したイメージは一旦、ユーザーの作業フォルダへWindowCopy.pngとして保存します。<br/>
    ''' 画面全体のハードコピー取得時は、メッセージボックス表示される可能性を考慮して0.5秒後に画面イメージを取得します。<br/>
    ''' 印刷又はプレビューする場合は、<see cref="PrintScreen.Copy"></see>実行前に、<see cref="PrintScreen.AutoPrint"></see>又は<see cref="PrintScreen.Preview"></see>を設定します。<br/>
    ''' 両方指定されているときは、プレビューを優先します。<br/>
    ''' 印刷時はデフォルトのプリンタへ印刷します。<br/>
    ''' <example>
    ''' 画面全体を取得し、プレビューする場合
    ''' <code lang="vb">
    ''' Dim ps As New PrintScreen
    ''' ps.Preview = True
    ''' ps.Copy()
    ''' </code>
    ''' </example>
    ''' </remarks>
	Public Class PrintScreen

		''' <summary>イメージ出力ファイル名</summary>
		Private Const C_FILENAME As String = "WindowCopy.png"

		''' <summary>対象となるWindow</summary>
		Private _ctrl As Control
		''' <summary>プレビューかどうか</summary>
		Private _preview As Boolean
		''' <summary>ハードコピーの出力ファイルパス</summary>
		Private _filename As String
		''' <summary>自動印刷するかどうか</summary>
		Private _autoPrint As Boolean

#Region " コンストラクタ "

		Public Sub New()
		End Sub

#End Region
#Region " Property "

		''' <summary>
		''' 自動印刷するかどうか
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property AutoPrint() As System.Boolean
			Get
				Return _autoPrint
			End Get
			Set(ByVal value As System.Boolean)
				_autoPrint = value
			End Set
		End Property

		''' <summary>
		''' プレビューするかどうか
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Preview() As System.Boolean
			Get
				Return _preview
			End Get
			Set(ByVal value As System.Boolean)
				_preview = value
			End Set
		End Property

		''' <summary>
		''' 画面イメージ出力ファイル名
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Filename() As System.String
			Get
				Return _filename
			End Get
			Set(ByVal value As System.String)
				_filename = value
			End Set
		End Property

#End Region
#Region " Method "

		''' <summary>
		''' 画面ハードコピー
		''' </summary>
		''' <remarks></remarks>
		Public Sub Copy()
			'メソッドをスレッドプールのキューに追加する
			ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf _doSomething))
		End Sub

		''' <summary>
		''' 画面ハードコピー
		''' </summary>
		''' <param name="ctrl">対象となる画面</param>
		''' <remarks></remarks>
		Public Sub Copy(ByVal ctrl As Control)
			_ctrl = ctrl
			_copyControl()
		End Sub

		''' <summary>
		''' スレッドで実行するメソッド
		''' </summary>
		''' <param name="obj"></param>
		''' <remarks></remarks>
		Private Sub _doSomething(ByVal obj As Object)
			Dim tim As New System.Timers.Timer

			AddHandler tim.Elapsed, AddressOf _copyTimer

			tim.Interval = 500
			tim.AutoReset = False	' 一度だけ
			tim.Enabled = True		' timer.Start()と同じ
		End Sub

		''' <summary>
		''' タイマーで実行するメソッド
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		''' <remarks></remarks>
		Private Sub _copyTimer(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
			Dim tim As System.Timers.Timer

			tim = DirectCast(sender, System.Timers.Timer)
			tim.Enabled = False

			_copyScreen()
		End Sub

		''' <summary>
		''' 画面全体のハードコピー
		''' </summary>
		''' <remarks></remarks>
		Private Sub _copyScreen()
			' Bitmapの作成
			Dim bmp As New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
			' Graphicsの作成
			Dim g As Graphics = Graphics.FromImage(bmp)
			' 画面全体をコピーする
			g.CopyFromScreen(New Point(0, 0), New Point(0, 0), bmp.Size)
			' 解放
			g.Dispose()

			' ファイルに保存する
			_filename = System.IO.Path.Combine(System.IO.Path.GetTempPath(), C_FILENAME)
			bmp.Save(_filename)
			' 後始末
			bmp.Dispose()

			' 印刷
			_print()
		End Sub

		''' <summary>
		''' 指定されたWindowのハードコピー
		''' </summary>
		''' <remarks></remarks>
		Private Sub _copyControl()
			' コントロールの外観を描画するBitmapの作成
			Dim bmp As New Bitmap(_ctrl.Width, _ctrl.Height)
			' キャプチャする
			_ctrl.DrawToBitmap(bmp, New Rectangle(0, 0, _ctrl.Width, _ctrl.Height))
			' ファイルに保存する
			_filename = System.IO.Path.Combine(System.IO.Path.GetTempPath(), C_FILENAME)
			bmp.Save(_filename)
			' 後始末
			bmp.Dispose()

			' 印刷
			_print()
		End Sub

		''' <summary>
		''' 印刷
		''' </summary>
		''' <remarks></remarks>
		Private Sub _print()
			' PrintDocumentオブジェクトの作成
			Dim pd As New System.Drawing.Printing.PrintDocument

			' PrintPageイベントハンドラの追加
			AddHandler pd.PrintPage, AddressOf pd_PrintPage

			' ページ設定
			Dim paperIdx As Integer = 0
			For Each ps As System.Drawing.Printing.PaperSize In pd.PrinterSettings.PaperSizes
				If ps.PaperName.IndexOf("A4") > -1 Then
					Exit For
				End If
				paperIdx += 1
			Next

			pd.DefaultPageSettings.PaperSize = pd.PrinterSettings.PaperSizes(paperIdx)
			pd.DefaultPageSettings.Landscape = True
			pd.DefaultPageSettings.Color = True

			If _preview Then
				' PrintPreviewDialogオブジェクトの作成
				Dim ppd As New PrintPreviewDialog
				' プレビューするPrintDocumentを設定
				ppd.Document = pd

				ppd.StartPosition = FormStartPosition.Manual
				ppd.SetDesktopBounds(0, 0, 800, 600)
				ppd.UseAntiAlias = True

				' 印刷プレビューダイアログを表示する
				ppd.ShowDialog()

				Exit Sub
			End If

			' 印刷しないときはここまで
			If Not _autoPrint Then
				Exit Sub
			End If

			' 印刷を開始する
			pd.Print()
		End Sub

		''' <summary>
		''' 印刷イベント
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		''' <remarks></remarks>
		Private Sub pd_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
			Dim zoom As Single = 1
			Dim padding As Integer = 20

			' 画像を読み込む
			Dim img As Image = Image.FromFile(_filename)

			Try
				' 用紙サイズに合わせて幅と高さを算出
				If img.Width > e.Graphics.VisibleClipBounds.Width Then
					zoom = e.Graphics.VisibleClipBounds.Width / img.Width
				End If
				If (img.Height + padding) * zoom > e.Graphics.VisibleClipBounds.Height Then
					zoom = e.Graphics.VisibleClipBounds.Height / (img.Height + padding)
				End If

				' 日付
				e.Graphics.DrawString(Now.ToString("yyyy/MM/dd (dddd) tt hh:mm:ss"), New Font("MS Gothic", 11), Brushes.Black, New Point(0, 0))

				' 画像を描画する
				e.Graphics.DrawImage(img, 0, padding, img.Width * zoom, img.Height * zoom)

				' 次のページがないことを通知する
				e.HasMorePages = False
			Finally
				' 後始末をする
				img.Dispose()
			End Try
		End Sub

#End Region

	End Class

End Namespace
