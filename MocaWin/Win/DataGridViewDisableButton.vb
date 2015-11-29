﻿
Namespace Win

	''' <summary>
	''' グリッドのボタン制御処理
	''' </summary>
	''' <remarks></remarks>
	Public Class DataGridViewDisableButtonColumn
		Inherits DataGridViewButtonColumn

		Public Sub New()
			Me.CellTemplate = New DataGridViewDisableButtonCell()
		End Sub
	End Class

	Public Class DataGridViewDisableButtonCell
		Inherits DataGridViewButtonCell

		Private enabledValue As Boolean
		Public Property Enabled() As Boolean
			Get
				Return enabledValue
			End Get
			Set(ByVal value As Boolean)
				enabledValue = value
			End Set
		End Property

		' Override the Clone method so that the Enabled property is copied.
		Public Overrides Function Clone() As Object
			Dim Cell As DataGridViewDisableButtonCell = _
				CType(MyBase.Clone(), DataGridViewDisableButtonCell)
			Cell.Enabled = Me.Enabled
			Return Cell
		End Function

		' By default, enable the button cell.
		Public Sub New()
			Me.enabledValue = True
		End Sub

		Protected Overrides Sub Paint(ByVal graphics As Graphics, _
			ByVal clipBounds As Rectangle, ByVal cellBounds As Rectangle, _
			ByVal rowIndex As Integer, _
			ByVal elementState As DataGridViewElementStates, _
			ByVal value As Object, ByVal formattedValue As Object, _
			ByVal errorText As String, _
			ByVal cellStyle As DataGridViewCellStyle, _
			ByVal advancedBorderStyle As DataGridViewAdvancedBorderStyle, _
			ByVal paintParts As DataGridViewPaintParts)

			' The button cell is disabled, so paint the border,  
			' background, and disabled button for the cell.
			If Not Me.enabledValue Then

				' Draw the background of the cell, if specified.
				If (paintParts And DataGridViewPaintParts.Background) = _
					DataGridViewPaintParts.Background Then

					Dim cellBackground As New SolidBrush(cellStyle.BackColor)
					graphics.FillRectangle(cellBackground, cellBounds)
					cellBackground.Dispose()
				End If

				' Draw the cell borders, if specified.
				If (paintParts And DataGridViewPaintParts.Border) = _
					DataGridViewPaintParts.Border Then

					PaintBorder(graphics, clipBounds, cellBounds, cellStyle, _
						advancedBorderStyle)
				End If

				' Calculate the area in which to draw the button.
				Dim buttonArea As Rectangle = cellBounds
				Dim buttonAdjustment As Rectangle = _
					Me.BorderWidths(advancedBorderStyle)
				buttonArea.X += buttonAdjustment.X
				buttonArea.Y += buttonAdjustment.Y
				buttonArea.Height -= buttonAdjustment.Height
				buttonArea.Width -= buttonAdjustment.Width

				' Draw the disabled button.                
				ButtonRenderer.DrawButton(graphics, buttonArea, _
				   System.Windows.Forms.VisualStyles.PushButtonState.Disabled)

				' Draw the disabled button text. 
				If TypeOf Me.FormattedValue Is String Then
					TextRenderer.DrawText(graphics, CStr(Me.FormattedValue), _
						cellStyle.Font, buttonArea, SystemColors.GrayText)
				End If

			Else
				' The button cell is enabled, so let the base class 
				' handle the painting.
				MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, _
					elementState, value, formattedValue, errorText, _
					cellStyle, advancedBorderStyle, paintParts)
			End If
		End Sub

	End Class

End Namespace
