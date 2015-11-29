
Namespace Win

	''' <summary>
	''' タブコントロールの制御を行うクラス
	''' </summary>
	''' <remarks></remarks>
	Public Class TabControlHelper

		''' <summary>操作するタブコントロール</summary>
		Private _tab As TabControl

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="tab">操作するタブコントロール</param>
		''' <remarks></remarks>
		Public Sub New(ByVal tab As TabControl)
			Me._tab = tab
		End Sub

#End Region

		''' <summary>
		''' タブページとして画面を追加する
		''' </summary>
		''' <param name="frm">追加するフォーム</param>
		''' <remarks></remarks>
		Public Sub AddPage(ByVal frm As Form)
			' トップレベルではタブに追加出来ないのでFalseにする。
			frm.TopLevel = False

			' タブに追加する
			Me._tab.TabPages.Add(frm.Name, frm.Text)
			Me._tab.TabPages.Item(frm.Name).Controls.Add(frm)

			' 各プロパティの設定
			frm.ControlBox = False
			frm.Text = String.Empty
			'frm.WindowState = FormWindowState.Maximized
			frm.Dock = DockStyle.Fill
			frm.FormBorderStyle = FormBorderStyle.None

			'フォームを表示する
			frm.Show()
			'最前面へ移動
			frm.BringToFront()
		End Sub

		''' <summary>
		''' タブページとして画面を追加する
		''' </summary>
		''' <param name="frmType">追加するフォームの型</param>
		''' <remarks></remarks>
		Public Sub AddPage(ByVal frmType As Type)
			Dim frm As Form

			frm = DirectCast(Activator.CreateInstance(frmType), Form)

			AddPage(frm)
		End Sub

	End Class

End Namespace
