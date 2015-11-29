
Imports System.Runtime.InteropServices

Namespace Win

	''' <summary>
	''' ウィンドウやコントロールにアニメーションさせる
	''' </summary>
	''' <remarks>
	''' user32.dll の <c>AnimateWindow</c> APIを使って表示させます。
	''' </remarks>
	Public Class AnimateWindow

#Region " Declare "

#Region " AnimateWindow API "

		''' <summary>
		''' AnimateWindow関数にて利用されるフラグ
		''' </summary>
		''' <remarks></remarks>
		<Flags()>
		Private Enum AnimateWindowFlags
			AW_HOR_POSITIVE = &H1
			AW_HOR_NEGATIVE = &H2
			AW_VER_POSITIVE = &H4
			AW_VER_NEGATIVE = &H8
			AW_CENTER = &H10
			AW_HIDE = &H10000
			AW_ACTIVATE = &H20000
			AW_SLIDE = &H40000
			AW_BLEND = &H80000
		End Enum

		''' <summary>
		''' AnimateWindow関数
		''' </summary>
		''' <param name="hwnd">ハンドル</param>
		''' <param name="time">アニメーションを行う時間</param>
		''' <param name="flags">挙動を表すフラグ</param>
		''' <returns></returns>
		''' <remarks></remarks>
		<DllImport("user32.dll")> _
		Private Shared Function AnimateWindow(ByVal hwnd As IntPtr, ByVal time As Integer, ByVal flags As AnimateWindowFlags) As Boolean
		End Function

#End Region

		''' <summary>
		'''アニメーション時間のデフォルト値
		''' </summary>
		''' <remarks></remarks>
		Private Const C_DEFAULT_TIME As Integer = 200

		''' <summary>
		''' アニメーションの方向
		''' </summary>
		''' <remarks></remarks>
		Public Enum DirectionType As Integer
			''' <summary>左から</summary>
			Left = AnimateWindowFlags.AW_HOR_POSITIVE
			''' <summary>右から</summary>
			Right = AnimateWindowFlags.AW_HOR_NEGATIVE
			''' <summary>上から</summary>
			Top = AnimateWindowFlags.AW_VER_POSITIVE
			''' <summary>下から</summary>
			Bottom = AnimateWindowFlags.AW_VER_NEGATIVE
		End Enum

#End Region

#Region " Slide "

		''' <summary>
		''' スライドしながら表示
		''' </summary>
		''' <param name="val"></param>
		''' <remarks></remarks>
		Public Sub Slide(ByVal val As Control, Optional ByVal direction As DirectionType = DirectionType.Left, Optional ByVal time As Integer = C_DEFAULT_TIME)
			AnimateWindow(val.Handle, time, CType(direction, AnimateWindowFlags) Or AnimateWindowFlags.AW_SLIDE)
			val.Show()
		End Sub

		''' <summary>
		''' スライドしながら非表示
		''' </summary>
		''' <param name="val"></param>
		''' <remarks></remarks>
		Public Sub SlideClose(ByVal val As Control, Optional ByVal direction As DirectionType = DirectionType.Left, Optional ByVal time As Integer = C_DEFAULT_TIME)
			AnimateWindow(val.Handle, time, CType(direction, AnimateWindowFlags) Or AnimateWindowFlags.AW_SLIDE Or AnimateWindowFlags.AW_HIDE)
			val.Hide()
		End Sub

#End Region
#Region " FadeIn "
		''' <summary>
		''' フェードインしながら表示
		''' </summary>
		''' <param name="val"></param>
		''' <remarks></remarks>
		Public Sub FadeIn(ByVal val As Control, Optional ByVal direction As DirectionType = DirectionType.Left, Optional ByVal time As Integer = C_DEFAULT_TIME)
			AnimateWindow(val.Handle, time, CType(direction, AnimateWindowFlags) Or AnimateWindowFlags.AW_BLEND)
			val.Show()
		End Sub

		''' <summary>
		''' フェードインしながら非表示
		''' </summary>
		''' <param name="val"></param>
		''' <remarks></remarks>
		Public Sub FadeInClose(ByVal val As Control, Optional ByVal direction As DirectionType = DirectionType.Left, Optional ByVal time As Integer = C_DEFAULT_TIME)
			AnimateWindow(val.Handle, time, CType(direction, AnimateWindowFlags) Or AnimateWindowFlags.AW_BLEND Or AnimateWindowFlags.AW_HIDE)
			val.Hide()
		End Sub

#End Region
#Region " Center "

		''' <summary>
		''' 中央から徐々に表示
		''' </summary>
		''' <param name="val"></param>
		''' <remarks></remarks>
		Public Sub Center(ByVal val As Control, Optional ByVal direction As DirectionType = DirectionType.Left, Optional ByVal time As Integer = C_DEFAULT_TIME)
			AnimateWindow(val.Handle, time, CType(direction, AnimateWindowFlags) Or AnimateWindowFlags.AW_CENTER)
			val.Show()
		End Sub

		''' <summary>
		''' 中央から徐々に非表示
		''' </summary>
		''' <param name="val"></param>
		''' <remarks></remarks>
		Public Sub CenterClose(ByVal val As Control, Optional ByVal direction As DirectionType = DirectionType.Left, Optional ByVal time As Integer = C_DEFAULT_TIME)
			AnimateWindow(val.Handle, time, CType(direction, AnimateWindowFlags) Or AnimateWindowFlags.AW_CENTER Or AnimateWindowFlags.AW_HIDE)
			val.Hide()
		End Sub

#End Region

	End Class

End Namespace
