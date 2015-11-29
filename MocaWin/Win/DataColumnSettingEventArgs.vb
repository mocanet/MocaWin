
Namespace Win

    ''' <summary>
    ''' DataColumn の列情報設定イベント引数
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DataColumnSettingEventArgs
        Inherits EventArgs

        ''' <summary>列位置</summary>
        Public Property Index As Integer

        ''' <summary>列情報</summary>
        Public Property Column As DataGridViewColumn

        ''' <summary>プロパティ</summary>
        Public Property Attribute As EnumGridColumnAttribute

        ''' <summary>ボタンのときは DataGridViewButtonColumn。違うときは Nothing</summary>
        Public ReadOnly Property Button As DataGridViewButtonColumn
            Get
                Return TryCast(Me.Column, DataGridViewButtonColumn)
            End Get
        End Property

        ''' <summary>テキストボックスのときは DataGridViewTextBoxColumn。違うときは Nothing</summary>
        Public ReadOnly Property TextBox As DataGridViewTextBoxColumn
            Get
                Return TryCast(Me.Column, DataGridViewTextBoxColumn)
            End Get
        End Property

        ''' <summary>コンボボックスのときは DataGridViewComboBoxColumn。違うときは Nothing</summary>
        Public ReadOnly Property ComboBox As DataGridViewComboBoxColumn
            Get
                Return TryCast(Me.Column, DataGridViewComboBoxColumn)
            End Get
        End Property

        ''' <summary>イメージのときは DataGridViewImageColumn。違うときは Nothing</summary>
        Public ReadOnly Property Image As DataGridViewImageColumn
            Get
                Return TryCast(Me.Column, DataGridViewImageColumn)
            End Get
        End Property

        ''' <summary>チェックボックスのときは DataGridViewCheckBoxColumn。違うときは Nothing</summary>
        Public ReadOnly Property CheckBox As DataGridViewCheckBoxColumn
            Get
                Return TryCast(Me.Column, DataGridViewCheckBoxColumn)
            End Get
        End Property

        ''' <summary>ボタンのときは DataGridViewDisableButtonColumn。違うときは Nothing</summary>
        Public ReadOnly Property DisableButton As DataGridViewDisableButtonColumn
            Get
                Return TryCast(Me.Column, DataGridViewDisableButtonColumn)
            End Get
        End Property

    End Class

End Namespace
