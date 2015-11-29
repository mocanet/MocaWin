
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Threading

Imports Moca.Exceptions

Namespace Win

	''' <summary>
	''' コントロールを扱う上で便利なメソッドたち
	''' </summary>
	''' <remarks></remarks>
	Public Class WinUtil

		''' <summary>メッセージボックスの表示スタイル</summary>
		Public Enum MessageBoxType As Integer
			''' <summary>情報メッセージアイコン    【表示形態：OKﾎﾞﾀﾝのみ】</summary>
			InformationOK = 0
			''' <summary>情報メッセージアイコン    【表示形態：はい・いいえﾎﾞﾀﾝのみ】</summary>
			InformationYesNo
			''' <summary>問い合わせアイコン        【表示形態：OKﾎﾞﾀﾝのみ】</summary>
			QuestionOK
			''' <summary>問い合わせアイコン        【表示形態：はい・いいえﾎﾞﾀﾝのみ】</summary>
			QuestionYesNo
			''' <summary>問い合わせアイコン        【表示形態：はい・いいえ・キャンセルﾎﾞﾀﾝのみ】</summary>
			QuestionYesNoCancel
			''' <summary>警告アイコン              【表示形態：OKﾎﾞﾀﾝのみ】</summary>
			CriticalOK
			''' <summary>警告アイコン              【表示形態：はい・いいえﾎﾞﾀﾝのみ】</summary>
			CriticalYesNO
			''' <summary>注意アイコン              【表示形態：OKﾎﾞﾀﾝのみ】</summary>
			ExclamationOK
			''' <summary>注意アイコン              【表示形態：はい・いいえﾎﾞﾀﾝのみ】</summary>
			ExclamationYesNo
			''' <summary>表示アイコンなし          【表示形態：OK・ﾍﾙﾌﾟﾎﾞﾀﾝのみ】</summary>
			ErrorOkHelp
			''' <summary>表示アイコンなし          【表示形態：OKﾎﾞﾀﾝのみ】</summary>
			ErrorTransactionOk
		End Enum

		''' <summary>アプリケーションでキャッチしきれていない例外をキャッチした時のリスナー</summary>
		Protected Shared appExceptionListener As IApplicationExceptionListener

		''' <summary>
		''' 入力系コントロールの自動クリア
		''' </summary>
		''' <param name="parent">親コントロール</param>
		''' <param name="bindingClear">バインディングのクリア指定</param>
		''' <param name="noExist">クリアしないコントロールの指定</param>
		''' <remarks>
		''' 指定されたコントロール内のすべてのコントロールを列挙し、入力内容をクリアする。<br/>
		''' <c>bindingClear</c> が True のときは、データがバインディングされているものはクリアされます。<br/>
		''' <c>noExist</c> に指定されたコントロールはクリア対象外となります。<br/>
		''' <list>
		''' <item><code>TextBoxBase からの派生型</code><description>Text をクリアする</description></item>
		''' <item><code>ListControl からの派生型</code><description>DataSource と Items をクリアする</description></item>
		''' <item><code>CheckBox の場合</code><description>Checked を False にする</description></item>
		''' <item><code>RadioButton の場合</code><description>Checked を False にする</description></item>
		''' <item><code>DateTimePicker の場合</code><description>Value を Date.Today にする</description></item>
		''' <item><code>NullableDateTimePicker の場合</code><description>Value を Nothing にする</description></item>
		''' </list>
		''' </remarks>
		Public Shared Sub ClearInput(ByVal parent As Control, Optional ByVal bindingClear As Boolean = False, Optional ByVal noExist As IList = Nothing)
			'Debug.Print(parent.ToString & ", " & parent.Name)
			If noExist IsNot Nothing Then
				If noExist.Contains(parent) Then
					Exit Sub
				End If
			End If

			' 子コントロールがあるとき
			If parent.Controls.Count > 0 Then
				For Each cControl As Control In parent.Controls
					ClearInput(cControl, bindingClear, noExist)
				Next
			End If

			' バインディングクリア
			If bindingClear Then
				parent.DataBindings.Clear()
			End If

			' TextBoxBase からの派生型
			If TypeOf parent Is TextBoxBase Then
				parent.Text = String.Empty
				Exit Sub
			End If
			' ListControl からの派生型
			If TypeOf parent Is ListControl Then
				If TypeOf parent Is ListBox Then
					' ListBox からの派生型
					Dim lstbx As ListBox = DirectCast(parent, ListBox)
					lstbx.BeginUpdate()
					lstbx.DataSource = Nothing
					lstbx.Items.Clear()
					lstbx.EndUpdate()
				ElseIf TypeOf parent Is ComboBox Then
					' ComboBox からの派生型
					Dim lstbx As ComboBox = DirectCast(parent, ComboBox)
					lstbx.BeginUpdate()
					lstbx.DataSource = Nothing
					lstbx.Items.Clear()
					lstbx.EndUpdate()
				Else
					DirectCast(parent, ListControl).DataSource = Nothing
				End If
				Exit Sub
			End If
			' CheckBox の場合
			If TypeOf parent Is CheckBox Then
				DirectCast(parent, CheckBox).Checked = False
				Exit Sub
			End If
			' RadioButton の場合
			If TypeOf parent Is RadioButton Then
				DirectCast(parent, RadioButton).Checked = False
				Exit Sub
			End If
			' DateTimePicker の場合
			If TypeOf parent Is DateTimePicker Then
				Try
					' NullableDateTimePicker の場合
					parent.GetType().InvokeMember("Value", BindingFlags.SetProperty Or BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public, Nothing, parent, New Object() {Nothing})
				Catch ex As Exception
					' DateTimePicker の場合
					DirectCast(parent, DateTimePicker).Value = Date.Today
				End Try
				Exit Sub
			End If
			'Debug.Print(parent.ToString & ", " & parent.Name)
		End Sub

        ''' <summary>
        ''' 指定したコントロール内に含まれる TextBox の Text をクリアする。
        ''' </summary>
        ''' <param name="parent"></param>
        ''' <remarks>
        ''' コントロールの型が TextBoxBase からの派生型の場合は Text をクリアする
        ''' </remarks>
        Public Shared Sub ClearTextBox(ByVal parent As System.Windows.Forms.Control)
            ' parent 内のすべてのコントロールを列挙する
            For Each cControl As System.Windows.Forms.Control In parent.Controls
                ' 列挙したコントロールにコントロールが含まれている場合は再帰呼び出しする
                If cControl.HasChildren Then
                    ClearTextBox(cControl)
                End If

                ' コントロールの型が TextBoxBase からの派生型の場合は Text をクリアする
                If TypeOf cControl Is TextBoxBase Then
                    cControl.DataBindings.Clear()
                    cControl.Text = String.Empty
                End If
            Next cControl
        End Sub

        ''' <summary>
        ''' 指定したコントロール内に含まれる ComboBox の Text をクリアする。
        ''' </summary>
        ''' <param name="parent"></param>
        ''' <remarks>
        ''' コントロールの型が ListControl からの派生型の場合は SelectedIndex をクリアする
        ''' </remarks>
        Public Shared Sub ClearComboBox(ByVal parent As System.Windows.Forms.Control)
            ' parent 内のすべてのコントロールを列挙する
            For Each cControl As System.Windows.Forms.Control In parent.Controls
                ' 列挙したコントロールにコントロールが含まれている場合は再帰呼び出しする
                If cControl.HasChildren Then
                    ClearComboBox(cControl)
                End If

                ' コントロールの型が ListControl からの派生型の場合は SelectedIndex をクリアする
                If TypeOf cControl Is ListControl Then
                    cControl.DataBindings.Clear()
                    DirectCast(cControl, ListControl).SelectedIndex = -1
                End If
            Next cControl
        End Sub

        ''' <summary>
        ''' コンボボックスを構築する
        ''' </summary>
        ''' <param name="cbo">対象のコンボボックス</param>
        ''' <param name="dataSource">データソース</param>
        ''' <param name="displayMember">リストに表示する列名</param>
        ''' <param name="valueMember">値にする列名</param>
        ''' <param name="selectedIndex">デフォルトのSelectedIndex</param>
        ''' <remarks>
        ''' </remarks>
        Public Shared Sub SetComboBox(ByVal cbo As ComboBox, ByVal dataSource As Object, ByVal displayMember As String, ByVal valueMember As String, Optional ByVal selectedIndex As Integer = -1)
			cbo.BeginUpdate()
			cbo.DisplayMember = displayMember
			cbo.ValueMember = valueMember
			cbo.DataSource = dataSource
			If cbo.DropDownStyle <> ComboBoxStyle.DropDownList Then
				cbo.SelectedIndex = selectedIndex
			End If
			cbo.EndUpdate()
		End Sub

		''' <summary>
		''' メッセージの内容により表示スタイルを変更する
		''' </summary>
		''' <param name="strMsg">表示メッセージ内容</param>
		''' <param name="LngFlg">表示形態</param>
		''' <returns>メッセージに対しての応答　OK・1　はい・6　いいえ・7　ﾍﾙﾌﾟ・処理が中断する(ﾍﾙﾌﾟﾌｧｲﾙ表示)</returns>
		''' <remarks>
		''' 表示形態に設定出来る値は下記のとおりです。
		''' 
		'''     0・情報メッセージアイコン   【表示形態：OKﾎﾞﾀﾝのみ】
		'''     1・情報メッセージアイコン   【表示形態：はい・いいえﾎﾞﾀﾝのみ】
		'''     2・問い合わせアイコン       【表示形態：OKﾎﾞﾀﾝのみ】
		'''     3・問い合わせアイコン       【表示形態：はい・いいえﾎﾞﾀﾝのみ】
		'''     4・警告アイコン             【表示形態：OKﾎﾞﾀﾝのみ】
		'''     5・警告アイコン             【表示形態：はい・いいえﾎﾞﾀﾝのみ】
		'''     6・注意アイコン             【表示形態：OKﾎﾞﾀﾝのみ】
		'''     7・注意アイコン             【表示形態：はい・いいえﾎﾞﾀﾝのみ】
		'''     8・表示アイコンなし         【表示形態：OK・ﾍﾙﾌﾟﾎﾞﾀﾝのみ】
		'''     9・表示アイコンなし         【表示形態：OKﾎﾞﾀﾝのみ】
		''' </remarks>
		Public Shared Function DispMessageBox(ByVal strMsg As String, ByVal LngFlg As MessageBoxType, ByVal Systemname As String) As MsgBoxResult
			Dim style As MsgBoxStyle
			Dim buttons As MessageBoxButtons
			Dim icon As MessageBoxIcon
			Dim buttonDef As MessageBoxDefaultButton
			Dim opt As MessageBoxOptions
			Dim help As Boolean
			Dim title As String

			title = Systemname

			Select Case LngFlg
				Case MessageBoxType.InformationOK '情報メッセージ【OK】
					buttons = MessageBoxButtons.OK
					icon = MessageBoxIcon.Information
				Case MessageBoxType.InformationYesNo '情報メッセージ【はい・いいえ】
					buttons = MessageBoxButtons.YesNo
					icon = MessageBoxIcon.Information
				Case MessageBoxType.QuestionOK	'問い合わせ【OK】
					buttons = MessageBoxButtons.OK
					icon = MessageBoxIcon.Question
				Case MessageBoxType.QuestionYesNo '問い合わせ【はい・いいえ】
					buttons = MessageBoxButtons.YesNo
					icon = MessageBoxIcon.Question
					buttonDef = MessageBoxDefaultButton.Button2
				Case MessageBoxType.QuestionYesNoCancel	'問い合わせ【はい・いいえ・キャンセル】
					buttons = MessageBoxButtons.YesNoCancel
					icon = MessageBoxIcon.Question
					buttonDef = MessageBoxDefaultButton.Button2
				Case MessageBoxType.CriticalOK	'警告【OK】
					buttons = MessageBoxButtons.OK
					icon = MessageBoxIcon.Exclamation
				Case MessageBoxType.CriticalYesNO '警告【はい・いいえ】
					buttons = MessageBoxButtons.YesNo
					icon = MessageBoxIcon.Exclamation
					buttonDef = MessageBoxDefaultButton.Button2
				Case MessageBoxType.ExclamationOK '注意【OK】
					buttons = MessageBoxButtons.OK
					icon = MessageBoxIcon.Exclamation
				Case MessageBoxType.ExclamationYesNo '注意【はい・いいえ】
					buttons = MessageBoxButtons.YesNo
					icon = MessageBoxIcon.Exclamation
					buttonDef = MessageBoxDefaultButton.Button2
				Case MessageBoxType.ErrorOkHelp	'エラー情報【OK・ﾍﾙﾌﾟ】
					style = MsgBoxStyle.MsgBoxHelp
					buttons = MessageBoxButtons.OK
					icon = MessageBoxIcon.Exclamation
					help = True
					title = "例外エラー発生 " & Systemname
				Case MessageBoxType.ErrorTransactionOk	'トランザクションエラー情報【中止】
					buttons = MessageBoxButtons.OK
					icon = MessageBoxIcon.Error
					title = " トランザクションエラー発生 " & Systemname
			End Select

			Select Case MessageBox.Show(strMsg, title, buttons, icon, buttonDef, opt, help)
				Case DialogResult.Abort
					Return MsgBoxResult.Abort
				Case DialogResult.Cancel
					Return MsgBoxResult.Cancel
				Case DialogResult.Ignore
					Return MsgBoxResult.Ignore
				Case DialogResult.No
					Return MsgBoxResult.No
				Case DialogResult.None
					Return 0
				Case DialogResult.OK
					Return MsgBoxResult.Ok
				Case DialogResult.Retry
					Return MsgBoxResult.Retry
				Case DialogResult.Yes
					Return MsgBoxResult.Yes
			End Select
		End Function

		''' <summary>
		''' イメージオブジェクトをバイト配列へ変換
		''' </summary>
		''' <param name="img">イメージオブジェクト</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function CImage2ByteArray(ByVal img As System.Drawing.Image) As Byte()
			Dim imgc As ImageConverter
			imgc = New ImageConverter
			Return DirectCast(imgc.ConvertTo(img, GetType(Byte())), Byte())
		End Function

		''' <summary>
		''' バイト配列をイメージオブジェクトへ変換
		''' </summary>
		''' <param name="img">バイト配列</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function CByteArray2Image(ByVal img As Byte()) As System.Drawing.Image
			Dim imgc As ImageConverter
			imgc = New ImageConverter
			Return DirectCast(imgc.ConvertFrom(img), Image)
		End Function

		''' <summary>
		''' コントロールに対してカルチャーを設定
		''' </summary>
		''' <param name="cultureName"></param>
		''' <param name="target"></param>
		''' <remarks></remarks>
		Public Shared Sub SetCultureControls(ByVal cultureName As String, ByVal target As Control)
			CultureUtil.SetCulture(cultureName)
			LanguageSetting(target)
		End Sub

		Protected Shared Sub LanguageSetting(ByVal target As Control)
			LanguageSetting(target, "Text")
			LanguageSetting(target, "ToolTipText")

			For Each ctrl As Control In target.Controls
				LanguageSetting(ctrl)
			Next
		End Sub

		Protected Shared Sub LanguageSetting(ByVal target As Control, ByVal propName As String)
			Dim bindFlags As BindingFlags
			Dim prop As PropertyInfo

			bindFlags = BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty

			prop = target.GetType.GetProperty(propName, bindFlags)
			If prop Is Nothing Then
				Return
			End If

			Dim txt As String
			txt = My.Resources.ResourceManager.GetString(target.Name & "_" & propName)
			If txt Is Nothing Then
				Return
			End If

			target.GetType.InvokeMember(propName, bindFlags, Nothing, target, New Object() {txt})
		End Sub

		''' <summary>
		''' ユーザーコントロール内で、デザインモードかどうか判定する。
		''' </summary>
		''' <returns>True:デザインモード、False:実行モード</returns>
		''' <remarks>
		''' ユーザーコントロール内では、親フォームがデザインモードかどうか判定するのに、Me.DesignMode が使えない。（必ず False のため）<br/>
		''' よって、こちらを使用する。
		''' </remarks>
		Public Shared Function UserControlDesignMode() As Boolean
			'If System.ComponentModel.DesignerProperties.GetIsInDesignMode(Me) Then
			'	Exit Sub
			'End If
			If System.ComponentModel.LicenseManager.UsageMode = System.ComponentModel.LicenseUsageMode.Designtime Then
				Return True
			ElseIf (Process.GetCurrentProcess().ProcessName.ToUpper().Equals("DEVENV")) Then
				Return True
			End If
			Return False
		End Function

		''' -----------------------------------------------------------------------------
		''' <summary>
		''' アプリケーションでキャッチしきれていない例外をキャッチする準備
		''' </summary>
		''' <param name="listener"></param>
		''' <remarks>
		''' 発生した例外がどこでもキャッチされていないと、ユーザーには理解できないメッセージが表示されてアプリケーションが落ちてしまう。
		''' これを避ける為に、アプリケーション全体で最終的に例外をキャッチする機能を付加する。
		''' </remarks>
		''' -----------------------------------------------------------------------------
		Public Shared Sub ApplicationExceptionHandler(ByVal listener As IApplicationExceptionListener)
			appExceptionListener = listener

			' ThreadExceptionイベント・ハンドラを登録する
			AddHandler Application.ThreadException, AddressOf Application_ThreadException
			' UnhandledExceptionイベント・ハンドラを登録する
			AddHandler Thread.GetDomain().UnhandledException, AddressOf Application_UnhandledException
		End Sub

		''' -----------------------------------------------------------------------------
		''' <summary>
		''' 未処理例外をキャッチするイベント・ハンドラ（Windowsアプリケーション用）
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		''' <remarks>
		''' </remarks>
		''' -----------------------------------------------------------------------------
		Protected Shared Sub Application_ThreadException(ByVal sender As Object, ByVal e As ThreadExceptionEventArgs)
			appExceptionListener.ApplicationException(e.Exception)
		End Sub

		''' -----------------------------------------------------------------------------
		''' <summary>
		''' 未処理例外をキャッチするイベント・ハンドラ （主にコンソール・アプリケーション用） 
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		''' <remarks>
		''' </remarks>
		''' -----------------------------------------------------------------------------
		Protected Shared Sub Application_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
			Dim ex As Exception = CType(e.ExceptionObject, Exception)
			If ex Is Nothing Then
				Exit Sub
			End If

			appExceptionListener.ApplicationException(ex)
		End Sub

	End Class

End Namespace
