
Imports System.Reflection

Namespace Win

	''' <summary>
	''' グリッドヘッダーに使う列挙型属性
	''' </summary>
	''' <remarks></remarks>
	<AttributeUsage(AttributeTargets.Field, allowmultiple:=False)> _
	Public Class EnumGridColumnAttribute
		Inherits Attribute

#Region " Declare "

		''' <summary>ヘッダータイトル</summary>
		Private _literal As String
		''' <summary>読取専用</summary>
		Private _ro As Boolean
		''' <summary>非表示</summary>
		Private _hidden As Boolean
		''' <summary>チェック項目</summary>
		Private _validateType As Moca.Util.ValidateTypes

		''' <summary>表示するエンティティのプロパティ名</summary>
		Private _mapPropertyName As String
		''' <summary>表示幅</summary>
		Private _width As Integer
		''' <summary>表示位置</summary>
		Private _align As DataGridViewContentAlignment
		''' <summary>表示フォーマット式</summary>
		Private _format As String
		''' <summary>セルの種類</summary>
		Private _cellType As Moca.Win.DataGridViewHelper.CellType

#End Region

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="literal">表示する項目名。省略時はEnumの文字をそのまま表示する。改行は『\n』</param>
		''' <param name="mapPropertyName">表示するエンティティのプロパティ名</param>
		''' <param name="width">表示幅</param>
		''' <param name="align">表示位置</param>
		''' <param name="format">表示フォーマット式</param>
		''' <param name="cellType">セルの種類</param>
		''' <param name="ro">読込み専用（編集不可列）。省略時はFalse</param>
		''' <param name="hidden">非表示列。省略時はFalse</param>
		''' <param name="validateType">チェック項目。省略時はなし</param>
		''' <remarks></remarks>
		Public Sub New(ByVal literal As String, ByVal mapPropertyName As String, ByVal width As Integer, Optional ByVal align As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft, Optional ByVal format As String = "", Optional ByVal cellType As Moca.Win.DataGridViewHelper.CellType = Moca.Win.DataGridViewHelper.CellType.TextBox, Optional ByVal ro As Boolean = False, Optional ByVal hidden As Boolean = False, Optional ByVal validateType As Moca.Util.ValidateTypes = Moca.Util.ValidateTypes.None)
			_literal = literal
			_ro = ro
			_hidden = hidden
			_validateType = validateType

			_mapPropertyName = mapPropertyName
			_width = width
			_align = align
			_format = format
			_cellType = cellType
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="mapPropertyName">表示するエンティティのプロパティ名</param>
		''' <param name="width">表示幅</param>
		''' <param name="align">表示位置</param>
		''' <param name="format">表示フォーマット式</param>
		''' <param name="cellType">セルの種類</param>
		''' <param name="ro">読込み専用（編集不可列）。省略時はFalse</param>
		''' <param name="hidden">非表示列。省略時はFalse</param>
		''' <param name="validateType">チェック項目。省略時はなし</param>
		''' <remarks></remarks>
		Public Sub New(ByVal mapPropertyName As String, ByVal width As Integer, Optional ByVal align As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft, Optional ByVal format As String = "", Optional ByVal cellType As Moca.Win.DataGridViewHelper.CellType = Moca.Win.DataGridViewHelper.CellType.TextBox, Optional ByVal ro As Boolean = False, Optional ByVal hidden As Boolean = False, Optional ByVal validateType As Moca.Util.ValidateTypes = Moca.Util.ValidateTypes.None)
			_literal = String.Empty
			_ro = ro
			_hidden = hidden
			_validateType = validateType

			_mapPropertyName = mapPropertyName
			_width = width
			_align = align
			_format = format
			_cellType = cellType
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="ro">読込み専用（編集不可列）。省略時はFalse</param>
		''' <param name="hidden">非表示列。省略時はFalse</param>
		''' <param name="validateType">チェック項目。省略時はなし</param>
		''' <remarks></remarks>
		Public Sub New(Optional ByVal ro As Boolean = False, Optional ByVal hidden As Boolean = False, Optional ByVal validateType As Moca.Util.ValidateTypes = Moca.Util.ValidateTypes.None)
			_literal = String.Empty
			_ro = ro
			_hidden = hidden
			_validateType = validateType
		End Sub

#End Region
#Region " Property "

		''' <summary>ヘッダータイトル</summary>
		Public ReadOnly Property Literal() As String
			Get
				Return _literal
			End Get
		End Property

		''' <summary>読取専用</summary>
		Public ReadOnly Property IsReadOnly() As Boolean
			Get
				Return _ro
			End Get
		End Property

		''' <summary>非表示</summary>
		Public ReadOnly Property IsHidden() As Boolean
			Get
				Return _hidden
			End Get
		End Property

		''' <summary>チェック項目</summary>
		Public ReadOnly Property IsValidateType() As Moca.Util.ValidateTypes
			Get
				Return _validateType
			End Get
		End Property

		''' <summary>表示するエンティティのプロパティ名</summary>
		Public ReadOnly Property MapPropertyName As String
			Get
				Return _mapPropertyName
			End Get
		End Property

		''' <summary>表示幅</summary>
		Public ReadOnly Property Width As Integer
			Get
				Return _width
			End Get
		End Property

		''' <summary>表示位置</summary>
		Public ReadOnly Property Align As DataGridViewContentAlignment
			Get
				Return _align
			End Get
		End Property

		''' <summary>表示フォーマット式</summary>
		Public ReadOnly Property Format As String
			Get
				Return _format
			End Get
		End Property

		''' <summary>セルの種類</summary>
		Public ReadOnly Property CellType As Moca.Win.DataGridViewHelper.CellType
			Get
				Return _cellType
			End Get
		End Property

#End Region
#Region " Method "

		''' <summary>
		''' 属性取得
		''' </summary>
		''' <param name="value"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function GetAttribute(ByVal value As [Enum]) As EnumGridColumnAttribute
			Dim t As Type

			t = value.GetType

			Dim name As String

			name = [Enum].GetName(t, value)

			Dim fi As FieldInfo
			fi = t.GetField(name)

			Dim aryAttr() As Attribute
			aryAttr = DirectCast(fi.GetCustomAttributes(GetType(EnumGridColumnAttribute), False), EnumGridColumnAttribute())

			If aryAttr.Length = 0 Then
                Return New EnumGridColumnAttribute(CBool(value.ToString))
			End If

			Return DirectCast(aryAttr(0), EnumGridColumnAttribute)
		End Function

#End Region

	End Class

End Namespace
