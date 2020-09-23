VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form SMSmainfrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "SMS messenger"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   ControlBox      =   0   'False
   Icon            =   "SMSmainfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "About"
      ForeColor       =   &H00A27E66&
      Height          =   1815
      Left            =   2760
      TabIndex        =   23
      Top             =   2520
      Width           =   3735
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "This Program is Programmed by Lakshmikantan G."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   720
         TabIndex        =   24
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Phone Settings"
      ForeColor       =   &H00A27E66&
      Height          =   2655
      Left            =   2640
      TabIndex        =   22
      Top             =   2040
      Width           =   3975
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "SMSmainfrm.frx":030A
         Left            =   1200
         List            =   "SMSmainfrm.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default"
         Height          =   255
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Text            =   "919844198441"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Apply"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Settings:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   440
         TabIndex        =   38
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Msg Validity:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   37
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Port No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SMSC No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   6240
   End
   Begin MSCommLib.MSComm SMS 
      Left            =   120
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      ForeColor       =   &H00A27E66&
      Height          =   1335
      Left            =   6240
      TabIndex        =   17
      Top             =   600
      Width           =   2055
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "About"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phone Settings"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sent Messages"
      ForeColor       =   &H00A27E66&
      Height          =   2415
      Left            =   1200
      TabIndex        =   14
      Top             =   3600
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid SMSgrid 
         Height          =   1935
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   8421504
         BackColorFixed  =   10649190
         ForeColorFixed  =   16777215
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   10649190
         GridColorFixed  =   4210752
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Message type"
      ForeColor       =   &H00A27E66&
      Height          =   1455
      Left            =   6240
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Blinking SMS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Flash SMS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A27E66&
      BorderStyle     =   0  'None
      FillColor       =   &H00A27E66&
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   7560
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   120
      Width           =   135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Message"
      ForeColor       =   &H00A27E66&
      Height          =   2895
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   4815
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save to SIM"
         Height          =   375
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Send SMS"
         Height          =   375
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         MaxLength       =   160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   840
         Width           =   4215
      End
      Begin VB.CheckBox DReport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delivery Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Destination No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A27E66&
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A27E66&
      Caption         =   "v1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   35
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5880
      Left            =   1080
      TabIndex        =   15
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00A27E66&
      Caption         =   " ©2002. g Inc."
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
      Height          =   255
      Left            =   6960
      TabIndex        =   20
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Line Line13 
      X1              =   0
      X2              =   8520
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line12 
      X1              =   0
      X2              =   8520
      Y1              =   6810
      Y2              =   6810
   End
   Begin VB.Line Line11 
      X1              =   8520
      X2              =   8520
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   7560
      X2              =   7680
      Y1              =   255
      Y2              =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00A27E66&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   2
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00A27E66&
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00A27E66&
      Caption         =   " SMS Messenger"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8390
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8390
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8390
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8390
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   8655
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8390
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8390
      Y1              =   6540
      Y2              =   6540
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8390
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8390
      Y1              =   6660
      Y2              =   6660
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "kanthji@rediffmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   32
      Top             =   6060
      Width           =   1695
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connecting..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   6075
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   6360
      Width           =   8775
   End
   Begin VB.Image Image1 
      Height          =   7875
      Left            =   -120
      Picture         =   "SMSmainfrm.frx":030E
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "SMSmainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
'================================================================
'***                  SMS Messenger v1.0                      ***
'================================================================
'================================================================
'*** For any Questions or Comments concerning this program    ***
'*** Homepage : http://xt.tp//www.sms.com                     ***
'*** Email    : kanthji@rediffmail.com                        ***
'*** Yahoo    : lucky_kanth                                   ***
'================================================================
'================================================================
'*** Frame Name  : SMSmainfrm.frm                             ***
'*** Date & Time : 21st October 2002, 04:30(IST)              ***
'================================================================
'================================================================


Private Sub Command1_Click()
   'Command1.Enabled = False
   Ssend = 1
   Call sendsms(Text3.Text, Text2.Text, Text1.Text)

End Sub

Private Sub Command5_Click()
   If SMS.CommPort <> Combo1.Text Then
      If SMS.PortOpen = True Then SMS.PortOpen = False
         SMS.CommPort = Combo1.Text
         SMS.InputLen = 0
         SMS.Settings = Combo2.Text
         SMS.PortOpen = True
   End If
   Frame5.Visible = False
End Sub

Private Sub Command8_Click()
   'Command8.Enabled = False
   Ssend = 2
   Call sendsms(Text3.Text, Text2.Text, Text1.Text)
End Sub

Private Sub Command2_Click()
   Frame5.Visible = True

End Sub

Private Sub Command3_Click()
   Frame6.Visible = True
End Sub

Private Sub Command4_Click()
   Frame6.Visible = False
End Sub


Private Sub Command6_Click()
   Frame5.Visible = False
End Sub



Private Sub Form_Load()
   Dim h As Integer
   SMSgrid.TextMatrix(0, 0) = "Status"
   SMSgrid.TextMatrix(0, 1) = "Destination"
   SMSgrid.TextMatrix(0, 2) = "Message"
   SMSgrid.TextMatrix(0, 3) = "Type"
   SMSgrid.ColWidth(0) = 800
   SMSgrid.ColWidth(1) = 1200
   SMSgrid.ColWidth(2) = 3900
   SMSgrid.ColWidth(3) = 800

   Frame6.Visible = False
   Frame5.Visible = False
   addcombo
   GetPrefs
   id = 0
   Text2.Text = "919844087634"
   Text1.Text = "SMS Messenger v1.0 by Lakshmikantan." & vbCrLf & ">SMSjadoo.com"
   
   Label4.Caption = Len(Text1.Text) & "/160"
   Command1.Enabled = False
   Command8.Enabled = False
   SMS.CommPort = Combo1.Text
   SMS.InputLen = 0
   SMS.Settings = Combo2.Text
   SMS.PortOpen = True
   SMS.Output = "AT+CPAS" & vbCrLf
End Sub

Private Sub Label3_Click()
   SMS.PortOpen = False
   SetPrefs
   Unload Me
End Sub

Private Sub Option1_Click()
    Label4.Caption = Len(Text1.Text) & "/160"
    Text1.MaxLength = 160
End Sub
Private Sub Option2_Click()
    Label4.Caption = Len(Text1.Text) & "/160"
    Text1.MaxLength = 160
End Sub
Private Sub Option3_Click()
    Label4.Caption = Len(Text1.Text) & "/70"
    Text1.MaxLength = 70
End Sub

Private Sub Picture1_Click()
   WindowState = 1
End Sub

Private Sub Text1_Change()
    If Option3.Value <> True Then
       B = 1
       For B = 1 To Len(Text1.Text)
           If Mid(Text1.Text, B, 1) = "|" Then D = "*"
       Next B
       If D = "*" Then
          Option3.Value = True
          Label4.Caption = Len(Text1.Text) & "/70"
          Text1.MaxLength = 70
       Else
          Label4.Caption = Len(Text1.Text) & "/160"
          Text1.MaxLength = 160
       End If
    Else
       Label4.Caption = Len(Text1.Text) & "/70"
       Text1.MaxLength = 70
    End If
End Sub

Private Sub Timer1_Timer()

    SMS.Output = "AT+CPAS" & vbCrLf
    presp = SMS.Input
    stat = findstr(presp, "+CPAS: 0")
    connected (stat)

End Sub

Private Sub connected(id)
   Select Case id
   Case 1: Label12.Caption = "Connected."
           Command1.Enabled = True
           Command8.Enabled = True
   Case Is <> 1: Label12.Caption = "Not Connected."
           Command1.Enabled = False
           Command8.Enabled = False
   End Select
End Sub
