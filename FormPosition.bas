Attribute VB_Name = "FormPosition"
'FormPosition.BAS
'Created by LividCreations
'http://lividcreations.cjb.net

'NOTE:
'   Switching from "FormOnBottom" to "FormOnTop" can only be done
'   in this order:  FormOnBottom, FormNormal, FormOnTop

'EXAMPLE OF ALL:
'   FORMONBOTTOM - FormOnBottom Form1
'   FORMNORMAL   - FormNormal Form1
'   FORMONTOP    - FormOnTop Form1



Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Sub FormOnBottom(Frm As Form)
'http://lividcreations.cjb.net

    'Example - FormOnBottom Form1

Dim DeskH As Long
DeskH = GethWndByWinTitle("Program Manager")
Call SetParent(Frm.hWnd, DeskH)
End Sub


Sub FormNormal(Frm As Form)
'http://lividcreations.cjb.net

    'Example - FormNormal Form1

Dim DeskH As Long
DeskH = GethWndByWinTitle("Form1")
Call SetParent(Frm.hWnd, DeskH)
End Sub



Sub FormOnTop(Frm As Form)
'http://lividcreations.cjb.net

    'Example - FormOnTop Form1

Call SetWindowPos(Frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub


Public Function GethWndByWinTitle(winTitle As String) As Long
    Dim retval As Long
    GethWndByWinTitle = FindWindow(vbNullString, winTitle)
End Function

