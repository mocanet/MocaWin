
Namespace Win

	''' <summary>
	''' タイトルバー無しウィンドウの移動制御
	''' </summary>
	''' <remarks></remarks>
	Public Class NoTitleBarWindowMover

#Region " Declre "

		Private _frm As Form

		''' <summary>マウスのクリック位置を記憶</summary>
		Private _mousePoint As Point

#End Region

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="frm">移動対象の画面</param>
		''' <remarks></remarks>
		Public Sub New(ByVal frm As Form)
			_frm = frm
		End Sub

#End Region

		''' <summary>
		''' 移動対象とするコントロールを追加
		''' </summary>
		''' <param name="controls"></param>
		''' <remarks></remarks>
		Public Sub AppendControls(ByVal ParamArray controls() As Control)
			For Each ctrl As Control In controls
				ctrl.Cursor = Cursors.SizeAll
				AddHandler ctrl.MouseDown, AddressOf Form_MouseDown
				AddHandler ctrl.MouseMove, AddressOf Form_MouseMove
			Next
		End Sub

		''' <summary>
		''' マウスダウンイベント
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		''' <remarks></remarks>
		Private Sub Form_MouseDown(sender As Object, e As MouseEventArgs)
			If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
				' 位置を記憶する
				_mousePoint = New Point(e.X, e.Y)
			End If
		End Sub

		''' <summary>
		''' マウス移動イベント
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		''' <remarks></remarks>
		Private Sub Form_MouseMove(sender As Object, e As MouseEventArgs)
			If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
				_frm.Left += e.X - _mousePoint.X
				_frm.Top += e.Y - _mousePoint.Y
			End If
		End Sub

	End Class

End Namespace
