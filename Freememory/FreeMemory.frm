VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "W-freeram"
   ClientHeight    =   3510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4560
   FontTransparent =   0   'False
   Icon            =   "FreeMemory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin Project1.OptmA OptmA 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Raminfo."
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4335
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   -120
         Top             =   1680
      End
      Begin VB.Label lbltotal 
         Caption         =   "&TOTAL PHYSICAL MEMORY"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "&AVAILABLE PHYSICAL MEMORY"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "&PERCENT FREE (%)"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.OptionButton optadvanced 
      Caption         =   "&Advanced"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.OptionButton optbasic 
      Caption         =   "&Basic"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   3720
      Top             =   0
   End
   Begin VB.CommandButton cmdoptimize 
      Caption         =   "&Optimize"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Label6"
      Height          =   135
      Left            =   6120
      TabIndex        =   1
      Top             =   2280
      Width           =   135
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnushow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MemoryStatus)
Private Type MemoryStatus
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Private Sub cmdoptimize_Click()
Recover_Memory
End Sub



Private Sub form_load()
OptmA.Caption = "Please select an optimization method."
On Error Resume Next
Me.Show
              Me.Refresh
              With nid
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon 'use form's icon in system tray
            .szTip = "W-Freeram" & vbNullChar 'tooltip text
        End With
        
    
    Shell_NotifyIcon NIM_ADD, nid 'add to tray

Dim ms As MemoryStatus
Call GlobalMemoryStatus(ms)
With ms
Label2.Caption = Format$(.dwTotalPhys \ 1024 \ 1024, "####.##" & "Mb")
Label3.Caption = Format$(.dwAvailPhys \ 1024 \ 1024, "####.##" & "Mb")
Label5.Caption = Format$(100 - .dwMemoryLoad, "##.00") & " %"
End With
End Sub

Sub Recover_Memory()
On Error Resume Next
    
   Timer1.Enabled = False
ReDim a(100)
Dim j As Long
OptmA.Max = 100

For j = 0 To 100
OptmA.Value = j
If optbasic.Value = True Then                 'optimization method selection
  a(j) = Space$(1000000)
 ElseIf optadvanced.Value = True Then
 a(j) = Space$(5000000)
 End If
If GetInputState <> 0 Then 'accept user input while optimizing
DoEvents
 End If
 OptmA.Caption = "[" & j / 100 * 100 & "%] Optimizing..."
 Next j
    
     For j = 0 To 100
    a(j) = vbNull
    Next j
    Timer1.Enabled = True
    OptmA.Caption = "Please select an optimization method."
    OptmA.Value = 0
 
End Sub


Private Sub mnuexit_Click()
 Shell_NotifyIcon NIM_DELETE, nid
 Unload Me: End
End Sub

Private Sub Timer1_Timer()
Dim ms As MemoryStatus
Call GlobalMemoryStatus(ms)
With ms
Label2.Caption = Format$(.dwTotalPhys / 1024 / 1024, "####.##" & "Mb")
Label3.Caption = Format$(.dwAvailPhys / 1024 / 1024, "####.##" & "Mb")
Label5.Caption = Format$(100 - .dwMemoryLoad, "##.00") & " %"
End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result, Action As Long
    
    'which display mode is being used
    
    If Me.ScaleMode = vbPixels Then
        Action = X
    Else
        Action = X / Screen.TwipsPerPixelX
    End If
    
Select Case Action

    Case WM_LBUTTONDBLCLK 'Left Button Double Click
        Me.WindowState = vbNormal 'put into taskbar
                 Me.Show
    
    Case WM_RBUTTONUP 'Right Button Up
            PopupMenu mnufile 'popup menu
    
    End Select
    
End Sub
Private Sub Form_Unload(Cancel As Integer) 'on form unload
    Cancel = 1
    Me.Hide
End Sub
Private Sub mnuShow_Click()
Me.Visible = True
Me.SetFocus
End Sub

Private Sub Timer2_Timer()
Dim ms As MemoryStatus
Call GlobalMemoryStatus(ms)
With ms
Label10.Caption = Format$(100 - .dwMemoryLoad, "##.00") & " %"
If 100 - .dwMemoryLoad < 40 Then cmdoptimize_Click
End With
End Sub

