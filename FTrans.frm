VERSION 5.00
Begin VB.Form FTrans 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Normalize"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   1800
      TabIndex        =   10
      Top             =   4320
      Width           =   1155
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Use &ColorKey"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3840
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3900
      Top             =   3000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Demo Mode"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   435
      Left            =   3180
      TabIndex        =   9
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Color&Key Value:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   3600
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   2355
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   4755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Important! Read the following:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Alpha Value:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   2880
      Width           =   900
   End
End
Attribute VB_Name = "FTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2001 Sveinn R. Sigurðsson
'  All Rights Reserved, http://www.svenni.com
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Const CLR_INVALID = &HFFFFFFFF

' Scrollbar as "updown" constants
Private Const MIN_VALUE = 50
Private Const MAX_VALUE = 255
Private Const LG_CHANGE = 5
Private Const SM_CHANGE = 1

Private Const WM_SETCURSOR = &H20
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_CONTEXTMENU = &H7B
Private Const WM_DISPLAYCHANGE = &H7E
Private Const WM_RBUTTONUP = &H205

' Member vars...
Private m_Trans As CTranslucentForm
Private m_Debug As Boolean

Private Sub Check1_Click()
   ' Turn off ColorKey mode
   Check2.Value = vbUnchecked
   
   ' Turn demo mode on/off
   Timer1.Enabled = (Check1.Value = vbChecked)
   Timer1.Interval = 100

   ' Enable normalize button
   Command2.Enabled = True
End Sub

Private Sub Check2_Click()
   Dim nColor As Long
   
   ' Disable demo mode, if running.
   Timer1.Enabled = False
   
   ' Set appropriate mode.
   If Check2.Value = vbChecked Then
      ' Update colorkey value.
      On Error Resume Next
         nColor = Val(Text2.Text)
      On Error GoTo 0
      m_Trans.ColorKey = nColor
   Else
      m_Trans.Mode = lwaAlpha
   End If

   ' Enable normalize button
   Command2.Enabled = True
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   m_Trans.Mode = lwaNormal
   Command2.Enabled = False
End Sub

Private Sub Form_Activate()
   ' Let user know if this probably won't work.
   If m_Trans.Supported = False Then
      MsgBox "Layered windows are only supported in " & _
         "Windows 2000.", vbExclamation, "Bummer"
   End If
End Sub

Private Sub Form_Load()
   ' Set scrollbar to act as "up-down" control.
   With VScroll1
      ' max < min, so down arrow = decrement,
      '   up arrow = INCREMENT
      .Max = MIN_VALUE
      .Min = MAX_VALUE
      .SmallChange = SM_CHANGE
      .LargeChange = LG_CHANGE
      ' start at HIGHEST value
      .Value = .Min
   End With
   
   ' Set up translucency.
   Set m_Trans = New CTranslucentForm
   m_Trans.hWnd = Me.hWnd
   Text2.Text = "&h" & Hex(m_Trans.ColorKey)
   
   ' Offer some instructions...
   Label3.Caption = "As you adjust the Alpha Value downward, " & _
      "using the scrollbar (updown) below, this form will " & _
      "become progressively more translucent. For your benefit " & _
      "the lower bound has been set to 50, to prevent the form " & _
      "from disappearing altogether. Should you lose sight of the " & _
      "anyway, press Escape to quit the demo." & _
      "If you would like to make the form disapear completely" & _
      "then simply change the value of the MIN_VALUE constant to 0." & _
      "You can also change the fading speed by increasing the " & _
      "value of SM_CHANGE and LG_CHANGE constants." & _
      "Don't forget to vote if you like this code." & _
      "For further information please contact : " & _
      "depill2000@hotmail.com"
   Text1.ToolTipText = "Use UpDown to right, or navigation keys, to adjust Alpha value."
   Text2.ToolTipText = "Right-click anywhere to grab the color under the cursor."
   Check1.ToolTipText = "Check to start auto-incrementing of Alpha value."
   Check2.ToolTipText = "Check to use ColorKey translucency."
   Me.Caption = "Translucency Demo"
   Set Me.Icon = Nothing
   
   ' Begin subclassing form and textboxes...
   Call HookWindow(Text1.hWnd, Me)
   Call HookWindow(Text2.hWnd, Me)
   Call HookWindow(Me.hWnd, Me)
   
   ' Check for debug flag
   m_Debug = CBool(InStr(Command, "/x"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call UnhookWindow(Text1.hWnd)
   Call UnhookWindow(Text2.hWnd)
   Call UnhookWindow(Me.hWnd)
End Sub

Private Sub Timer1_Timer()
   Static GoingUp As Boolean
   Dim nVal As Long
   ' Slide translucency.
   With VScroll1
      On Error Resume Next
      If GoingUp Then
         .Value = .Value + LG_CHANGE
      Else
         .Value = .Value - LG_CHANGE
      End If
      If Err.Number Then
         GoingUp = Not GoingUp
      End If
   End With
End Sub

Private Sub VScroll1_Change()
   ' Updates textbox value when scrollbar is changed.
   Text1.Text = VScroll1.Value
   If Me.Visible Then
      ' Focus shift to textbox, if not
      ' in demo mode.
      If Check1.Value = vbUnchecked Then
         Text1.SetFocus
      End If
      ' Update translucency.
      m_Trans.Alpha = VScroll1.Value
   End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   ' Change scrollbar value using up and down
   ' arrows when TextBox has the input focus.
   VScroll1.SetFocus
   Select Case KeyCode
      Case vbKeyUp
         SendKeys "{UP}"
      Case vbKeyDown
         SendKeys "{DOWN}"
      Case vbKeyHome
         SendKeys "{END}"
      Case vbKeyEnd
         SendKeys "{HOME}"
      Case vbKeyPageUp
         SendKeys "{PGUP}"
      Case vbKeyPageDown
         SendKeys "{PGDN}"
   End Select
   KeyCode = 0
End Sub

Friend Function WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
   Dim Result As Long
   
   Select Case msg
      Case WM_SETCURSOR
         If m_Debug Then
            Me.Caption = Hex(GetColor)
         End If
         Result = CallWindowProc(GetProp(hWnd, MHookMe.keyWndProc), hWnd, msg, wp, lp)
         
      Case WM_CONTEXTMENU
         Text2.Text = "&h" & Right$("000000" & Hex(GetColor()), 6)
         Result = 0
         
      Case WM_RBUTTONUP
         Text2.Text = "&h" & Right$("000000" & Hex(GetColor()), 6)
         Result = CallWindowProc(GetProp(hWnd, MHookMe.keyWndProc), hWnd, msg, wp, lp)
         
      Case WM_DISPLAYCHANGE
         ' Force refresh of layered window
         m_Trans.Mode = m_Trans.Mode
         Result = CallWindowProc(GetProp(hWnd, MHookMe.keyWndProc), hWnd, msg, wp, lp)
         
      Case Else
         ' Pass along to default window procedure.
         Result = CallWindowProc(GetProp(hWnd, MHookMe.keyWndProc), hWnd, msg, wp, lp)
         
   End Select
   
   ' Return desired result code to Windows.
   WindowProc = Result
End Function

Private Function GetColor() As Long
   Dim hWnd As Long
   Dim hDC As Long
   Dim pt As POINTAPI
   Dim nColor As Long
   
   ' Grab the color under the cursor.
   Call GetCursorPos(pt)
   hWnd = WindowFromPoint(pt.x, pt.y)
   hDC = GetDC(hWnd)
   Call ScreenToClient(hWnd, pt)
   nColor = GetPixel(hDC, pt.x, pt.y)
   If nColor = CLR_INVALID Then
      Call BitBlt(Me.hDC, 0, 0, 1, 1, hDC, pt.x, pt.y, vbSrcCopy)
      nColor = Me.Point(0, 0)
   End If
   Call ReleaseDC(hWnd, hDC)
   GetColor = nColor
End Function
