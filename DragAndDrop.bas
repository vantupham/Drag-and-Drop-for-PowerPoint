Attribute VB_Name = "Module1"
' ===============================
' Drag & Drop
' Written by Pham Van Tu
' Email: phamtund@gmail.com
' Mobile: 0939.725.119
' ===============================
 
Option Explicit
Public Type PointAPI
  x As Long
  y As Long
End Type
  
Public Type RECT
  lLeft As Long
  lTop As Long
  lRight As Long
  lBottom As Long
End Type
Public currSldnum As Long
Private Const SM_SCREENX = 0
Private Const SM_SCREENY = 1
Private Const msgCancel = "."
Private Const msgNoXlInstance = "."
Private Const sigProc = "Drag & Drop"
Private Const VK_SHIFT = &H10
Private Const VK_CTRL = &H11
Private Const VK_ALT = &H12
#If VBA7 Then
  Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As LongPtr) As Integer
  Public Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As LongPtr, ByVal yPoint As LongPtr) As LongPtr
  Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As LongPtr
  Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As LongPtr
  Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As LongPtr) As LongPtr
#Else
  Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
  Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
  Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
  Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
  Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
Public mPoint As PointAPI
Private ActiveShape As Shape
Private dragMode As Boolean
Private dx As Double, dy As Double
Sub DragAndDrop(oShp As Shape)
    currSldnum = SlideShowWindows(1).View.CurrentShowPosition
    If CBool(GetKeyState(VK_SHIFT) And &HF0000000) And CBool(GetKeyState(VK_ALT) And &HF0000000) Then DragCalculate oShp: Exit Sub
    dragMode = Not dragMode
    DoEvents
    'If oShp.HasTextFrame And dragMode Then oShp.TextFrame.TextRange.Copy
    dx = GetSystemMetrics(SM_SCREENX)
    dy = GetSystemMetrics(SM_SCREENY)
    Drag oShp
    ActivePresentation.Slides(currSldnum).Shapes("Timer").TextFrame.TextRange.Text = ""
    DoEvents
End Sub
Private Sub Drag(oShp As Shape)
  #If VBA7 Then
    Dim mWnd As LongPtr
  #Else
    Dim mWnd As Long
  #End If
  Dim sx As Long, sy As Long
  Dim WR As RECT
  Dim StartTime As Single
  ' Thoi gian di chuyen
  Const DropInSeconds = 3
  GetCursorPos mPoint
  mWnd = WindowFromPoint(mPoint.x, mPoint.y)
  GetWindowRect mWnd, WR
  sx = WR.lLeft
  sy = WR.lTop
  Debug.Print sx, sy
  With ActivePresentation.PageSetup
    dx = (WR.lRight - WR.lLeft) / .SlideWidth
    dy = (WR.lBottom - WR.lTop) / .SlideHeight
    Select Case True
      Case dx > dy
        sx = sx + (dx - dy) * .SlideWidth / 2
        dx = dy
      Case dy > dx
        sy = sy + (dy - dx) * .SlideHeight / 2
        dy = dx
    End Select
  End With
  StartTime = Timer
  While dragMode
    GetCursorPos mPoint
    oShp.Left = (mPoint.x - sx) / dx - oShp.Width / 2
    oShp.Top = (mPoint.y - sy) / dy - oShp.Height / 2
    ActivePresentation.Slides(currSldnum).Shapes("Timer").Left = oShp.Left - 30
    ActivePresentation.Slides(currSldnum).Shapes("Timer").Top = oShp.Top - 9
    ' Hien thi thoi gian tren shape sau khi click
    ActivePresentation.Slides(currSldnum).Shapes("Timer").TextFrame.TextRange.Text = CInt(DropInSeconds - (Timer - StartTime))
    DoEvents
    'dieu kien dung di chuyen shape
    If Timer > StartTime + DropInSeconds Then dragMode = False
  Wend
  DoEvents
End Sub
Private Sub DragCalculate(oShp As Shape)
  Dim xl As Object
  Dim FormulaArray
  If oShp.HasTextFrame Then
    Set xl = CreateObject("Excel.Application")
    If xl Is Nothing Then MsgBox msgNoXlInstance, vbCritical, "Quiz": Exit Sub
    FormulaArray = Split(oShp.TextFrame.TextRange.Text & "=", "=")
    While InStr(FormulaArray(0), ",") > 0
      FormulaArray(0) = Replace(FormulaArray(0), ",", ".")
    Wend
    If FormulaArray(0) > "" Then
      FormulaArray(1) = xl.Evaluate(FormulaArray)
      oShp.TextFrame.TextRange.Text = FormulaArray(0) & "=" & FormulaArray(1)
    End If
    xl.Quit: Set xl = Nothing
    oShp.Top = oShp.Top + 1: oShp.Top = oShp.Top - 1
  End If
  DoEvents
End Sub
