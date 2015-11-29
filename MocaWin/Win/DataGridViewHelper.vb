
Imports System.ComponentModel

Namespace Win

    ''' <summary>
    ''' DataGridView の操作を補助するコンポーネント
    ''' </summary>
    ''' <remarks></remarks>
    <Description("DataGridView の操作を補助するコンポーネント"), _
    ToolboxItem(True), _
    DesignTimeVisible(True)> _
    Public Class DataGridViewHelper

#Region " Declare "

        ''' <summary>
        ''' セルのコントロール種別
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum CellType As Integer
            ''' <summary>ボタン</summary>
            Button = 0
            ''' <summary>チェックボックス</summary>
            CheckBox
            ''' <summary>コンボボックス</summary>
            ComboBox
            ''' <summary>イメージ</summary>
            Image
            ''' <summary>リンク</summary>
            Link
            ''' <summary>テキストボックス</summary>
            TextBox
            ''' <summary>ボタン（無効化可能）</summary>
            DisableButton
        End Enum

#Region " Events "

        ''' <summary>
        ''' グリッドの設定変更イベント
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' このイベントでグリッドの固定列、固定行を設定しても反映されません。<br/>
        ''' これはVB.NETの初期化タイミングの問題なので回避できません。<br/>
        ''' デザイン時に設定するか、フォームロード時に設定してください。
        ''' </remarks>
        Public Event TargetGridSetting(ByVal sender As Object, ByVal e As TargetGridSettingEventArgs)

        ''' <summary>
        ''' DataColumn の列情報設定イベント
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Event DataColumnSetting(ByVal sender As Object, ByVal e As DataColumnSettingEventArgs)

#End Region

#Region " 列名編集用コード "

        ''' <summary>列の改行コード</summary>
        Public Const C_COLTITLE_CR As String = "\n"

#End Region

        ''' <summary>固定列位置</summary>
        Private _frozenIndex As Integer

        ''' <summary>固定列のセルスタイル</summary>
        Private _frozenCellStyle As DataGridViewCellStyle

        ''' <summary>制御対象となるDataGridView</summary>
        Private _grd As DataGridView

        ''' <summary>補助するグリッド</summary>
        Private WithEvents _targetGrid As DataGridView

        ''' <summary>グリッドの列定義列挙</summary>
        Private _gridColumnCaptions As [Enum]

#End Region

#Region " コンストラクタ "

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="grd">制御するDataGridView</param>
        ''' <remarks>
        ''' 指定されたグリッドを下記の状態に初期化します。<br/>
        ''' <list>
        ''' <item><description>固定列はなしとする。</description></item>
        ''' <item><description>固定列のセルスタイルはデフォルトセルスタイルとする。</description></item>
        ''' <item><description>列指定はクリアする。</description></item>
        ''' <item><description>ヘッダー列のセルの内容の表示場所は <see cref="DataGridViewContentAlignment.MiddleCenter" /> とする。</description></item>
        ''' <item><description><see cref="DataGridView.DataSource" /> 設定時の自動列作成をOFF (<c>AutoGenerateColumns</c> = False) </description></item>
        ''' </list>
        ''' </remarks>
        Public Sub New(ByVal grd As DataGridView)
            _grd = grd
            FrozenIndex = 0
            FrozenCellStyle = _grd.DefaultCellStyle.Clone()

            _grd.Columns.Clear()
            _grd.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            _grd.AutoGenerateColumns = False    ' DataSource設定時の自動列作成をOFF
        End Sub

#End Region

#Region " プロパティ "

        ''' <summary>
        ''' 操作対象となるグリッド
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Description("操作対象となるC1FlexGridコントロール")> _
        Public Property TargetGrid() As DataGridView
            Get
                Return _targetGrid
            End Get
            Set(ByVal value As DataGridView)
                _targetGrid = value
                ' セットアップ
                _setupGrid()
            End Set
        End Property

        ''' <summary>
        ''' グリッドの列定義列挙
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Browsable(False)> _
        Public Property GridColumnCaptions() As [Enum]
            Get
                Return _gridColumnCaptions
            End Get
            Set(value As [Enum])
                _gridColumnCaptions = value
                ' 列セットアップ
                _setDataTableColumns()
            End Set
        End Property

        ''' <summary>固定列位置</summary>
        Public Property FrozenIndex() As Integer
            Get
                Return _frozenIndex
            End Get
            Set(ByVal value As Integer)
                _frozenIndex = value
            End Set
        End Property

        ''' <summary>固定列のセルスタイル</summary>
        Public Property FrozenCellStyle() As DataGridViewCellStyle
            Get
                Return _frozenCellStyle
            End Get
            Set(ByVal value As DataGridViewCellStyle)
                _frozenCellStyle = value
            End Set
        End Property

        ''' <summary>
        ''' ダブルバッファー
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' 表示のちらつきを抑えたいときはTrueにする
        ''' </remarks>
        Public Property DoubleBuffered As Boolean
            Get
                Dim dgvType As Type = GetType(DataGridView)
                Dim dgvPropInfo As System.Reflection.PropertyInfo = dgvType.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
                Return DirectCast(dgvPropInfo.GetValue(_grd, Nothing), Boolean)
            End Get
            Set(value As Boolean)
                Dim dgvType As Type = GetType(DataGridView)
                Dim dgvPropInfo As System.Reflection.PropertyInfo = dgvType.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
                dgvPropInfo.SetValue(_grd, value, Nothing)
            End Set
        End Property

#End Region

#Region " Shared Method "

        ''' <summary>
        ''' 各種コードを削除して正式な列名を返す。
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CnvColCaption(ByVal val As String) As String
            Dim caption As String = val

            caption = caption.Replace(C_COLTITLE_CR, Environment.NewLine)

            Return caption
        End Function

        ''' <summary>
        ''' 表示名称を返す
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetColCaption(ByVal val As [Enum]) As String
            Dim attr As EnumGridColumnAttribute
            attr = EnumGridColumnAttribute.GetAttribute(CType(val, [Enum]))

            Return CnvColCaption(attr.Literal)
        End Function

        ''' <summary>
        ''' 読み取り専用列か判定する
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function IsColReadOnly(ByVal val As [Enum]) As Boolean
            Dim attr As EnumGridColumnAttribute
            attr = EnumGridColumnAttribute.GetAttribute(CType(val, [Enum]))

            Return attr.IsReadOnly()
        End Function

        ''' <summary>
        ''' 非表示列か判定する
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function IsColHidden(ByVal val As [Enum]) As Boolean
            Dim attr As EnumGridColumnAttribute
            attr = EnumGridColumnAttribute.GetAttribute(CType(val, [Enum]))

            Return attr.IsHidden()
        End Function

        ''' <summary>
        ''' チェック項目
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function ColValidateType(ByVal val As [Enum]) As Moca.Util.ValidateTypes
            Dim attr As EnumGridColumnAttribute
            attr = EnumGridColumnAttribute.GetAttribute(CType(val, [Enum]))

            Return attr.IsValidateType()
        End Function

#End Region

        ''' <summary>
        ''' 列の追加
        ''' </summary>
        ''' <param name="propertyName">DBの列とバインディングする名称</param>
        ''' <param name="caption">画面に表示するタイトル</param>
        ''' <param name="width">列幅</param>
        ''' <param name="align">データの表示位置</param>
        ''' <param name="format">データの表示フォーマット</param>
        ''' <param name="cellTyp">セルのコントロール種別</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AddColumn(ByVal propertyName As String, ByVal caption As String, ByVal width As Integer, ByVal align As DataGridViewContentAlignment, Optional ByVal format As String = "", Optional ByVal cellTyp As CellType = DataGridViewHelper.CellType.TextBox) As DataGridViewColumn
            Dim col As DataGridViewColumn
            Dim colIndex As Integer

            col = _makeColumn(cellTyp)
            col.HeaderText = caption
            col.DataPropertyName = propertyName
            col.Name = propertyName
            col.Width = width
            col.DefaultCellStyle.Alignment = align
            col.SortMode = DataGridViewColumnSortMode.NotSortable

            colIndex = _grd.Columns.Add(col)

            col.DefaultCellStyle.Font = _grd.DefaultCellStyle.Font

            ' 固定行の設定有無判定
            If colIndex >= FrozenIndex Then
                Return col
            End If

            ' 固定行なのでスタイルを設定する
            col.Frozen = True
            col.DefaultCellStyle = FrozenCellStyle.Clone()
            col.DefaultCellStyle.Alignment = align

            Return col
        End Function

        ''' <summary>
        ''' テキストボックス表示の列を追加
        ''' </summary>
        ''' <param name="propertyName">DBの列とバインディングする名称</param>
        ''' <param name="caption">画面に表示するタイトル</param>
        ''' <param name="width">列幅</param>
        ''' <param name="align">データの表示位置</param>
        ''' <param name="format">データの表示フォーマット</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AddTxtColumn(ByVal propertyName As String, ByVal caption As String, ByVal width As Integer, ByVal align As DataGridViewContentAlignment, Optional ByVal format As String = "") As DataGridViewColumn
            Return AddColumn(propertyName, caption, width, align, format, CellType.TextBox)
        End Function

        ''' <summary>
        ''' コンボボックス表示の列を追加
        ''' </summary>
        ''' <param name="propertyName">DBの列とバインディングする名称</param>
        ''' <param name="caption">画面に表示するタイトル</param>
        ''' <param name="width">列幅</param>
        ''' <param name="align">データの表示位置</param>
        ''' <param name="dataSource">コンボボックスへバインドするデータソース</param>
        ''' <param name="displayMember">コンボ ボックスに表示する文字列の取得先となるプロパティまたは列を指定する文字列を取得または設定します。 </param>
        ''' <param name="valueMember">ドロップダウン リストの選択項目に対応する値の取得先となる、プロパティまたは列を指定する文字列を取得または設定します。 </param>
        ''' <param name="format">データの表示フォーマット</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AddCboColumn(ByVal propertyName As String, ByVal caption As String, ByVal width As Integer, ByVal align As DataGridViewContentAlignment, ByVal dataSource As DataTable, ByVal displayMember As String, ByVal valueMember As String, Optional ByVal format As String = "") As DataGridViewColumn
            Dim col As DataGridViewColumn
            col = AddColumn(propertyName, caption, width, align, format, CellType.ComboBox)
            Return _setComboBoxItems(col, dataSource, displayMember, valueMember)
        End Function

        ''' <summary>
        ''' コンボボックスへ表示するデータをバインドする
        ''' </summary>
        ''' <param name="propertyName">バインドしたい列名称</param>
        ''' <param name="dataSource">コンボボックスへバインドするデータソース</param>
        ''' <param name="displayMember">コンボ ボックスに表示する文字列の取得先となるプロパティまたは列を指定する文字列を取得または設定します。 </param>
        ''' <param name="valueMember">ドロップダウン リストの選択項目に対応する値の取得先となる、プロパティまたは列を指定する文字列を取得または設定します。 </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SetComboBoxItems(ByVal propertyName As String, ByVal dataSource As DataTable, ByVal displayMember As String, ByVal valueMember As String) As DataGridViewComboBoxColumn
            Dim col As DataGridViewColumn
            col = _grd.Columns.Item(propertyName)
            Return _setComboBoxItems(col, dataSource, displayMember, valueMember)
        End Function

        ''' <summary>
        ''' コンボボックスへ表示するデータをバインドする
        ''' </summary>
        ''' <param name="index">列位置</param>
        ''' <param name="dataSource">コンボボックスへバインドするデータソース</param>
        ''' <param name="displayMember">コンボ ボックスに表示する文字列の取得先となるプロパティまたは列を指定する文字列を取得または設定します。 </param>
        ''' <param name="valueMember">ドロップダウン リストの選択項目に対応する値の取得先となる、プロパティまたは列を指定する文字列を取得または設定します。 </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SetComboBoxItems(ByVal index As Integer, ByVal dataSource As DataTable, ByVal displayMember As String, ByVal valueMember As String) As DataGridViewComboBoxColumn
            Dim col As DataGridViewColumn
            col = _grd.Columns.Item(index)
            Return _setComboBoxItems(col, dataSource, displayMember, valueMember)
        End Function

        ''' <summary>
        ''' コンボボックスへ表示するデータをバインドする
        ''' </summary>
        ''' <param name="col">列</param>
        ''' <param name="dataSource">コンボボックスへバインドするデータソース</param>
        ''' <param name="displayMember">コンボ ボックスに表示する文字列の取得先となるプロパティまたは列を指定する文字列を取得または設定します。 </param>
        ''' <param name="valueMember">ドロップダウン リストの選択項目に対応する値の取得先となる、プロパティまたは列を指定する文字列を取得または設定します。 </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function _setComboBoxItems(ByVal col As DataGridViewColumn, ByVal dataSource As DataTable, ByVal displayMember As String, ByVal valueMember As String) As DataGridViewComboBoxColumn
            If Not TypeOf col Is DataGridViewComboBoxColumn Then
                Throw New ArgumentException("指定された列はコンボボックススタイルになっていません。")
            End If

            Dim cbo As DataGridViewComboBoxColumn = DirectCast(col, DataGridViewComboBoxColumn)

            cbo.DataSource = dataSource
            cbo.DisplayMember = displayMember
            cbo.ValueMember = valueMember

            Return cbo
        End Function

        ''' <summary>
        ''' 指定されたセル種別でDataGridViewColumnを作成する。
        ''' </summary>
        ''' <param name="type"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function _makeColumn(ByVal type As CellType) As DataGridViewColumn
            Dim col As DataGridViewColumn
            Select Case type
                Case CellType.Button
                    col = New DataGridViewButtonColumn()
                Case CellType.DisableButton
                    col = New DataGridViewDisableButtonColumn()
                Case CellType.CheckBox
                    col = New DataGridViewCheckBoxColumn()
                Case CellType.ComboBox
                    col = New DataGridViewComboBoxColumn()
                    Dim cbo As DataGridViewComboBoxColumn = DirectCast(col, DataGridViewComboBoxColumn)
                    ' 現在のセルしかコンボボックスが表示されないようにする
                    cbo.DisplayStyleForCurrentCellOnly = True
                    ' 編集モードの時だけコンボボックスを表示する
                    cbo.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                Case CellType.Image
                    col = New DataGridViewImageColumn()
                Case CellType.Link
                    col = New DataGridViewLinkColumn()
                Case Else
                    col = New DataGridViewTextBoxColumn()
            End Select

            Return col
        End Function

        ''' <summary>
        ''' グリッドのセットアップ
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub _setupGrid()
            '' デザイン時はここまで
            'If MyBase.DesignMode Then
            '	Exit Sub
            'End If

            If _targetGrid Is Nothing Then
                Return
            End If

            ' 固定行列定義
            Dim args As TargetGridSettingEventArgs
            args = New TargetGridSettingEventArgs
            args.TargetGrid = Me.TargetGrid

            RaiseEvent TargetGridSetting(Me, args)
        End Sub

        ''' <summary>
        ''' データテーブルの列定義設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub _setDataTableColumns()
            Dim args As DataColumnSettingEventArgs

            Me.TargetGrid.Columns.Clear()

            args = New DataColumnSettingEventArgs

            For ii As Integer = 0 To [Enum].GetValues(Me.GridColumnCaptions.GetType).Length - 1
                Dim caption As String
                Dim val As [Enum]
                val = DirectCast([Enum].GetValues(Me.GridColumnCaptions.GetType).GetValue(ii), [Enum])

                Dim attr As EnumGridColumnAttribute
                attr = EnumGridColumnAttribute.GetAttribute(val)

                caption = attr.Literal
                If caption.Length = 0 Then
                    caption = val.ToString
                End If

                Dim column As DataGridViewColumn
                column = AddColumn(attr.MapPropertyName,
                                          CnvColCaption(caption),
                                          attr.Width,
                                          attr.Align,
                                          attr.Format,
                                          attr.CellType)
                column.MinimumWidth = attr.Width
                column.ReadOnly = attr.IsReadOnly
                column.DefaultCellStyle.Format = attr.Format
                column.Visible = Not attr.IsHidden

                ' 列のデフォルト値やその他設定
                args.Index = ii
                args.Column = column
                args.Attribute = attr
                RaiseEvent DataColumnSetting(Me, args)
            Next
        End Sub

    End Class

End Namespace
