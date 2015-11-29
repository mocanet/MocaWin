
Namespace Win

	''' <summary>
	''' StatusStripに現在の時間を表示する
	''' </summary>
	''' <remarks></remarks>
	<System.Windows.Forms.Design.ToolStripItemDesignerAvailability( _
	 System.Windows.Forms.Design.ToolStripItemDesignerAvailability.StatusStrip)> _
	Public Class ToolStripStatusDateTime
		Inherits ToolStripStatusLabel

		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>タイマー</summary>
		Private _timer As Timer

		Private _format As String

		Public Sub New()
			'初期化
			components = New System.ComponentModel.Container()
			Me.Text = Me.GetFormattedDateTimeText()
			_format = "t"
			_timer = New Timer(components)
			_timer.Interval = 100
			'Tickイベントハンドラの追加
			AddHandler _timer.Tick, AddressOf _timer_Tick
			_timer.Enabled = True
		End Sub

		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing And Not (components Is Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		''' <summary>
		''' 日時のフォーマット
		''' </summary>
		<System.Configuration.DefaultSettingValue("t")> _
		Public Property Format() As String
			Get
				Return _format
			End Get
			Set(ByVal value As String)
				_format = value
			End Set
		End Property

		<System.ComponentModel.Browsable(False)> _
		Public Overrides Property [Text]() As String
			Get
				Return MyBase.Text
			End Get
			Set(ByVal value As String)
				MyBase.Text = value
				Me.OnTextChanged(Nothing)
			End Set
		End Property

		Private Sub _timer_Tick(ByVal sender As Object, ByVal e As EventArgs)
			'Textを変更する
			Dim txt As String = Me.GetFormattedDateTimeText()
			If Me.Text <> txt Then
				Me.Text = txt
			End If
		End Sub

		Private Function GetFormattedDateTimeText() As String
			Return DateTime.Now.ToString(Me._format)
		End Function

	End Class

End Namespace
