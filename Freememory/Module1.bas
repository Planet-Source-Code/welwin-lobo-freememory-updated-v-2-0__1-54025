Attribute VB_Name = "Module1"
Option Explicit

      Type NOTIFYICONDATA
         cbSize As Long
         hwnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type

      
      Global Const NIM_ADD = &H0
      Global Const NIM_MODIFY = &H1
      Global Const NIM_DELETE = &H2
Global Const WM_MOUSEMOVE = &H200
     Global Const NIF_MESSAGE = &H1
     Global Const NIF_ICON = &H2
      Global Const NIF_TIP = &H4
      Global Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Global Const WM_LBUTTONDOWN = &H201     'Button down
      Global Const WM_LBUTTONUP = &H202       'Button up
          Global Const WM_RBUTTONDBLCLK = &H206   'Double-click
      Global Const WM_RBUTTONDOWN = &H204     'Button down
      Global Const WM_RBUTTONUP = &H205       'Button up

      
      Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

     
      Global nid As NOTIFYICONDATA







     


