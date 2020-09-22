VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTranslucentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

'   BOOL SetLayeredWindowAttributes(
'     HWND hwnd,           // handle to the layered window
'     COLORREF crKey,      // specifies the color key
'     BYTE bAlpha,         // value for the blend function
'     DWORD dwFlags        // action
'   );
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

'   BOOL UpdateLayeredWindow(
'     HWND hwnd,             // handle to layered window
'     HDC hdcDst,            // handle to screen DC
'     POINT *pptDst,         // new screen position
'     SIZE *psize,           // new size of the layered window
'     HDC hdcSrc,            // handle to surface DC
'     POINT *pptSrc,         // layer position
'     COLORREF crKey,        // color key
'     BLENDFUNCTION *pblend, // blend function
'     DWORD dwFlags          // options
'   );
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, pptDst As POINTAPI, pSize As SIZE, ByVal hdcSrc As Long, pptSrc As POINTAPI, ByVal crKey As Long, pBlend As BLENDFUNCTION, ByVal dwFlags As Long) As Long

'   typedef struct _BLENDFUNCTION {
'     BYTE     BlendOp;
'     BYTE     BlendFlags;
'     BYTE     SourceConstantAlpha;
'     BYTE     AlphaFormat;
'   }BLENDFUNCTION, *PBLENDFUNCTION, *LPBLENDFUNCTION;
Private Type BLENDFUNCTION
   BlendOp As Byte
   BlendFlags As Byte
   SourceConstantAlpha As Byte
   AlphaFormat As Byte
End Type

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type SIZE
   cx As Long
   cy As Long
End Type

Private Const LWA_COLORKEY = &H1&
Private Const LWA_ALPHA = &H2&
Private Const LWA_OPAQUE = &HFF&

Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4

' Style setting APIs
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

' Win32 APIs to determine OS information.
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Color translation
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

' Translucency Modes
Public Enum TranslucencyModes
   lwaNormal = 0
   lwaColorKey = LWA_COLORKEY
   lwaAlpha = LWA_ALPHA
End Enum

' Trick to keep the case of our Enums
#If False Then
   Private Const lwaNormal = 0
   Private Const lwaColorKey = 1
   Private Const lwaAlpha = 2
#End If

' Member variables
Private m_Supported As Boolean
Private m_hWnd As Long
Private m_Alpha As Long
Private m_ColorKey As OLE_COLOR
Private m_Mode As TranslucencyModes

' ************************************************
'  Initialize/Terminate
' ************************************************
Private Sub Class_Initialize()
   ' Make sure we're in Windows 2000, or at
   ' least some version capable of layered
   ' windows.
   Dim os As OSVERSIONINFO
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   m_Supported = (os.dwMajorVersion >= 5)
   ' Set some default values.
   m_Alpha = LWA_OPAQUE
   m_ColorKey = vb3DFace
   m_Mode = lwaAlpha
End Sub

Private Sub Class_Terminate()
   ' Restore window to normal.
   ClearTranslucency m_hWnd
End Sub

' ************************************************
'  Public Properties
' ************************************************
Public Property Let Alpha(ByVal NewVal As Byte)
   If m_Supported Then
      ' Set window translucency, and cache value
      ' if successful.
      Select Case m_Mode
         Case lwaAlpha
            If SetLayeredWindowAttributes(m_hWnd, 0, CLng(NewVal), LWA_ALPHA) Then
               m_Alpha = NewVal
            End If
         Case lwaColorKey, lwaNormal
            m_Mode = lwaAlpha
            ToggleTranslucency m_hWnd
      End Select
   End If
End Property

Public Property Get Alpha() As Byte
   Alpha = CByte(m_Alpha)
End Property

Public Property Let ColorKey(ByVal NewVal As OLE_COLOR)
   If m_Supported Then
      ' Set window translucency, and cache value
      ' if successful.
      Select Case m_Mode
         Case lwaColorKey
            If SetLayeredWindowAttributes(m_hWnd, CheckSysColor(NewVal), 0, LWA_COLORKEY) Then
               m_ColorKey = NewVal
            End If
         Case lwaAlpha, lwaNormal
            m_Mode = lwaColorKey
            m_ColorKey = NewVal
            ToggleTranslucency m_hWnd
      End Select
   End If
End Property

Public Property Get ColorKey() As OLE_COLOR
   ColorKey = m_ColorKey
End Property

Public Property Let hWnd(ByVal NewVal As Long)
   If m_Supported Then
      ' Restore previous window to normal.
      ClearTranslucency m_hWnd
      ' Cache handle to new window, and
      ' setup for translucency.
      m_hWnd = NewVal
      SetTranslucency m_hWnd
   End If
End Property

Public Property Get hWnd() As Long
   hWnd = m_hWnd
End Property

Public Property Let Mode(ByVal NewVal As TranslucencyModes)
   Select Case NewVal
      Case lwaColorKey, lwaAlpha
         ' Toggle translucency to clear
         ' previous settings.
         m_Mode = NewVal
         Call ToggleTranslucency(m_hWnd)
      
      Case lwaNormal
         m_Mode = NewVal
         Call ClearTranslucency(m_hWnd)
         
      Case Else
         ' ignore input
   End Select
End Property

Public Property Get Mode() As TranslucencyModes
   Mode = m_Mode
End Property

' ************************************************
'  Public Properties - Read/Only
' ************************************************
Public Property Get Supported() As Boolean
   Supported = m_Supported
End Property

' ************************************************
'  Private Methods
' ************************************************
Private Function CheckSysColor(ByVal ColorRef As OLE_COLOR) As Long
   Const HighBit = &H80000000
   If ColorRef And HighBit Then
      CheckSysColor = GetSysColor(ColorRef And Not HighBit)
   Else
      CheckSysColor = ColorRef
   End If
End Function

Private Function ClearTranslucency(ByVal hWnd As Long) As Boolean
   Dim nStyle As Long
   If hWnd Then
      ' Set translucency to fully opaque.
      Call SetLayeredWindowAttributes(hWnd, 0, LWA_OPAQUE, LWA_ALPHA)
      ' Clear exstyle bit.
      nStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
      ClearTranslucency = CBool(SetWindowLong(hWnd, GWL_EXSTYLE, nStyle))
   End If
End Function

Private Function SetTranslucency(ByVal hWnd As Long) As Boolean
   Dim nStyle As Long
   If hWnd Then
      ' Set exstyle bit.
      nStyle = GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
      If SetWindowLong(hWnd, GWL_EXSTYLE, nStyle) Then
         ' Reset cached value for translucency, and
         ' corresponding window, to fully opaque.
         m_Alpha = LWA_OPAQUE
         SetTranslucency = CBool(SetLayeredWindowAttributes(hWnd, 0, m_Alpha, LWA_ALPHA))
      End If
   End If
End Function

Private Function ToggleTranslucency(ByVal hWnd As Long) As Boolean
   Dim nStyle As Long
   If hWnd Then
      ' Clear, then reset, exstyle bit.
      nStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
      If m_Mode <> lwaNormal Then
         If SetWindowLong(hWnd, GWL_EXSTYLE, nStyle) Then
            nStyle = nStyle Or WS_EX_LAYERED
            If SetWindowLong(hWnd, GWL_EXSTYLE, nStyle) Then
               Select Case m_Mode
                  Case lwaAlpha
                     ToggleTranslucency = CBool(SetLayeredWindowAttributes(hWnd, 0, m_Alpha, LWA_ALPHA))
                  Case lwaColorKey
                     ToggleTranslucency = CBool(SetLayeredWindowAttributes(hWnd, CheckSysColor(m_ColorKey), 0, LWA_COLORKEY))
               End Select
            End If
         End If
      End If
   End If
End Function

