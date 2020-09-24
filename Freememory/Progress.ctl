VERSION 5.00
Begin VB.UserControl OptmA 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   4740
   Begin VB.PictureBox MainBox 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox Progress 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   15
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   15
         Begin VB.Label Stat2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2175
            TabIndex        =   2
            Top             =   60
            Width           =   465
         End
      End
      Begin VB.Label Stat1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click the level 1,2 or 3 button below to begin optimization."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   60
         Width           =   4170
      End
   End
End
Attribute VB_Name = "OptmA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ProgVal As Integer
Private MaxNum As Long
Public Property Let Max(lngNum As Long)

    MaxNum = lngNum

End Property
Public Property Get Max() As Long

    Max = MaxNum

End Property
Public Property Let Value(IntValue As Long)

    On Error Resume Next
    If IntValue = 0 Then
        Progress.Visible = False
        Else
        Progress.Visible = True
    End If
    ProgVal = IntValue
    Progress.Width = MainBox.Width * (ProgVal / MaxNum)
    Refresh

End Property
Public Property Get Value() As Long

    ProgVal = Value

End Property
Public Property Let Caption(MyCaption As String)

    Stat1 = MyCaption
    Stat2 = MyCaption

End Property
Public Property Get Caption() As String

    Caption = Stat1

End Property

Private Sub UserControl_Initialize()

    Progress.Visible = False
    UserControl_Resize

End Sub
Private Sub UserControl_Resize()

    MainBox.Width = UserControl.Width
    MainBox.Height = UserControl.Height
    Stat1.Left = 50
    Stat1.Top = (MainBox.Height / 2) - (Stat1.Height / 2) - 30
    Stat2.Left = 50
    Stat2.Top = Stat1.Top
    Progress.Height = MainBox.Height

End Sub



