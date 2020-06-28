Imports System.Drawing.Drawing2D

Public Class frmMain

	Public Declare Auto Function SystemParametersInfo Lib "user32" (ByVal uAction As Integer, ByVal uParam As Integer, ByRef pvParam As Integer, ByVal fuWinIni As Integer) As Boolean
	Public Const SPI_SCREENSAVERRUNNING As Long = 97

	'Public Const intPosun As Integer = 7
	'Public Const intPenWidth As Integer = 1

	Public intPosun As Integer = Rnd() * 20
	Public intPenWidth As Integer = Rnd() * 5

	Public goGr As Graphics

	Public blnIsInitialized As Boolean = False
	Public intMaxLines As Integer = (150 - (intPosun * 2.5)) \ 2
	Public intLinesCount As Integer = 0
	Public blnQuit As Boolean = False

	Public Structure tBody
		Dim intXStart As Integer
		Dim intYStart As Integer
		Dim intXEnd As Integer
		Dim intYEnd As Integer
	End Structure

	Public pPenShow As Pen
	Public pPenMirrRT As Pen
	Public pPenMirrRB As Pen
	Public pPenMirrLB As Pen
	Public pPenClear As Pen
	Public oBSMain As tBody
	Public oBSMirr As tBody
	Public oBCMain As tBody
	Public oBCMirr As tBody
	Public oPosunShow As tBody
	Public oPosunClear As tBody

	Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Dim i As Long

		Randomize()

		Dim tmpLng As Integer
		tmpLng = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 1&, 0&, 0&)

		'Dim pRegKey As RegistryKey = Registry.CurrentUser
		'pRegKey = pRegKey.OpenSubKey("Software\Test Screen Saver", True)
		'Dim val As Object = pRegKey.GetValue("Message", "Enter Text here")
		'pRegKey.Close()

		goGr = Me.CreateGraphics()

		SetStyle(ControlStyles.UserPaint, True)
		SetStyle(ControlStyles.AllPaintingInWmPaint, True)
		SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

		SetStyle(ControlStyles.UserPaint, True)
		SetStyle(ControlStyles.AllPaintingInWmPaint, True)
		SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

		Initialize()

		Refresh()

	End Sub

	Public Function GetRNDColor() As Color
		GetRNDColor = Color.FromArgb(CInt(100 + Rnd() * 150), CInt(100 + Rnd() * 150), CInt(100 + Rnd() * 150))
	End Function

	Public Sub Initialize()

		With oBSMain

			.intXStart = Rnd() * Me.ClientSize.Width
			.intYStart = Rnd() * Me.ClientSize.Height
			.intXEnd = Rnd() * Me.ClientSize.Width
			.intYEnd = Rnd() * Me.ClientSize.Height

		End With

		oBCMain = oBSMain

		oPosunShow.intXStart = intPosun
		oPosunShow.intYStart = intPosun
		oPosunShow.intXEnd = intPosun
		oPosunShow.intYEnd = (intPosun \ 2)

		oPosunClear = oPosunShow

		pPenShow = New Pen(Color.Red)
		pPenShow.Width = intPenWidth
		pPenShow.DashStyle = DashStyle.Solid
		pPenShow.Color = GetRNDColor()

		pPenMirrRT = New Pen(Color.Red)
		pPenMirrRT.Width = intPenWidth
		pPenMirrRT.DashStyle = DashStyle.Solid
		pPenMirrRT.Color = GetRNDColor()

		pPenMirrRB = New Pen(Color.Red)
		pPenMirrRB.Width = intPenWidth
		pPenMirrRB.DashStyle = DashStyle.Solid
		pPenMirrRB.Color = GetRNDColor()

		pPenMirrLB = New Pen(Color.Red)
		pPenMirrLB.Width = intPenWidth
		pPenMirrLB.DashStyle = DashStyle.Solid
		pPenMirrLB.Color = GetRNDColor()

		pPenClear = New Pen(Color.Black)
		pPenClear.Width = intPenWidth
		pPenClear.DashStyle = DashStyle.Solid

		blnIsInitialized = True

	End Sub

	Public Sub Kaleidoskop()

		Do
			If blnQuit Then Exit Do

			Application.DoEvents()
			'-----------------------
			'mirror coordinates
			oBSMirr.intXStart = Me.ClientSize.Width - oBSMain.intXStart
			oBSMirr.intYStart = Me.ClientSize.Height - oBSMain.intYStart
			oBSMirr.intXEnd = Me.ClientSize.Width - oBSMain.intXEnd
			oBSMirr.intYEnd = Me.ClientSize.Height - oBSMain.intYEnd
			'------------
			oBCMirr.intXStart = Me.ClientSize.Width - oBCMain.intXStart
			oBCMirr.intYStart = Me.ClientSize.Height - oBCMain.intYStart
			oBCMirr.intXEnd = Me.ClientSize.Width - oBCMain.intXEnd
			oBCMirr.intYEnd = Me.ClientSize.Height - oBCMain.intYEnd
			'------------
			'animated lines for each part of screen
			goGr.DrawLine(pPenShow, oBSMain.intXStart, oBSMain.intYStart, oBSMain.intXEnd, oBSMain.intYEnd)
			goGr.DrawLine(pPenMirrRT, oBSMirr.intXStart, oBSMain.intYStart, oBSMirr.intXEnd, oBSMain.intYEnd)
			goGr.DrawLine(pPenMirrRB, oBSMain.intXStart, oBSMirr.intYStart, oBSMain.intXEnd, oBSMirr.intYEnd)
			goGr.DrawLine(pPenMirrLB, oBSMirr.intXStart, oBSMirr.intYStart, oBSMirr.intXEnd, oBSMirr.intYEnd)
			'------------------------
			oBSMain.intXStart = oBSMain.intXStart - oPosunShow.intXStart
			oBSMain.intYStart = oBSMain.intYStart + oPosunShow.intYStart
			oBSMain.intXEnd = oBSMain.intXEnd + oPosunShow.intXEnd
			oBSMain.intYEnd = oBSMain.intYEnd - oPosunShow.intYEnd
			'------bounce effect------
			If oBSMain.intXStart > Me.ClientSize.Width Then
				oPosunShow.intXStart = intPosun
				pPenShow.Color = GetRNDColor()
				pPenMirrRT.Color = GetRNDColor()
				pPenMirrRB.Color = GetRNDColor()
				pPenMirrLB.Color = GetRNDColor()
			End If
			If oBSMain.intXStart < 0 Then
				oPosunShow.intXStart = -intPosun
				pPenShow.Color = GetRNDColor()
				pPenMirrRT.Color = GetRNDColor()
				pPenMirrRB.Color = GetRNDColor()
				pPenMirrLB.Color = GetRNDColor()
			End If
			If oBSMain.intXEnd > Me.ClientSize.Width Then
				oPosunShow.intXEnd = -intPosun
				pPenShow.Color = GetRNDColor()
				pPenMirrRT.Color = GetRNDColor()
				pPenMirrRB.Color = GetRNDColor()
				pPenMirrLB.Color = GetRNDColor()
			End If
			If oBSMain.intXEnd < 0 Then
				oPosunShow.intXEnd = intPosun
				pPenShow.Color = GetRNDColor()
				pPenMirrRT.Color = GetRNDColor()
				pPenMirrRB.Color = GetRNDColor()
				pPenMirrLB.Color = GetRNDColor()
			End If
			If oBSMain.intYStart > Me.ClientSize.Height Then
				oPosunShow.intYStart = -intPosun
				pPenShow.Color = GetRNDColor()
				pPenMirrRT.Color = GetRNDColor()
				pPenMirrRB.Color = GetRNDColor()
				pPenMirrLB.Color = GetRNDColor()
			End If
			If oBSMain.intYStart < 0 Then
				oPosunShow.intYStart = intPosun
				pPenShow.Color = GetRNDColor()
				pPenMirrRT.Color = GetRNDColor()
				pPenMirrRB.Color = GetRNDColor()
				pPenMirrLB.Color = GetRNDColor()
			End If
			If oBSMain.intYEnd > Me.ClientSize.Height Then
				oPosunShow.intYEnd = intPosun
				pPenShow.Color = GetRNDColor()
				pPenMirrRT.Color = GetRNDColor()
				pPenMirrRB.Color = GetRNDColor()
				pPenMirrLB.Color = GetRNDColor()
			End If
			If oBSMain.intYEnd < 0 Then
				oPosunShow.intYEnd = -(intPosun \ 2)
				pPenShow.Color = GetRNDColor()
				pPenMirrRT.Color = GetRNDColor()
				pPenMirrRB.Color = GetRNDColor()
				pPenMirrLB.Color = GetRNDColor()
			End If
			'---white lines----
			If (intLinesCount >= intMaxLines) Then

				goGr.DrawLine(pPenClear, oBCMain.intXStart, oBCMain.intYStart, oBCMain.intXEnd, oBCMain.intYEnd)
				goGr.DrawLine(pPenClear, oBCMirr.intXStart, oBCMain.intYStart, oBCMirr.intXEnd, oBCMain.intYEnd)
				goGr.DrawLine(pPenClear, oBCMain.intXStart, oBCMirr.intYStart, oBCMain.intXEnd, oBCMirr.intYEnd)
				goGr.DrawLine(pPenClear, oBCMirr.intXStart, oBCMirr.intYStart, oBCMirr.intXEnd, oBCMirr.intYEnd)
				'--------------------------
				oBCMain.intXStart = oBCMain.intXStart - oPosunClear.intXStart
				oBCMain.intYStart = oBCMain.intYStart + oPosunClear.intYStart
				oBCMain.intXEnd = oBCMain.intXEnd + oPosunClear.intXEnd
				oBCMain.intYEnd = oBCMain.intYEnd - oPosunClear.intYEnd
				'------------
				If oBCMain.intXStart > Me.ClientSize.Width Then
					oPosunClear.intXStart = intPosun
				End If
				If oBCMain.intXStart < 0 Then
					oPosunClear.intXStart = -intPosun
				End If
				If oBCMain.intXEnd > Me.ClientSize.Width Then
					oPosunClear.intXEnd = -intPosun
				End If
				If oBCMain.intXEnd < 0 Then
					oPosunClear.intXEnd = intPosun
				End If

				If oBCMain.intYStart > Me.ClientSize.Height Then
					oPosunClear.intYStart = -intPosun
				End If
				If oBCMain.intYStart < 0 Then
					oPosunClear.intYStart = intPosun
				End If
				If oBCMain.intYEnd > Me.ClientSize.Height Then
					oPosunClear.intYEnd = intPosun
				End If
				If oBCMain.intYEnd < 0 Then
					oPosunClear.intYEnd = -(intPosun \ 2)
				End If

			Else

				If (intLinesCount < intMaxLines) Then intLinesCount = intLinesCount + 1

			End If

			'Me.Invalidate()
			Application.DoEvents()

		Loop Until blnQuit

		If blnQuit Then Me.Close()

	End Sub

	Private Sub frmMain_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

		If blnIsInitialized Then

			Call Kaleidoskop()

		End If

	End Sub

	Private Sub frmMain_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
		Static pLastPosition As Point

		If (pLastPosition.X = 0) And (pLastPosition.Y = 0) Then

			pLastPosition = e.Location
			Exit Sub

		End If

		If (Math.Abs(pLastPosition.X - e.X) > 3) Or (Math.Abs(pLastPosition.Y - e.Y) > 3) Then

			blnQuit = True

		End If

	End Sub

	Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

		blnQuit = True

	End Sub
End Class
