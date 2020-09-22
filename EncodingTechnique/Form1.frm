VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encoding Techniques Simulator"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   6030
      Top             =   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Simulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   6525
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   4365
         Left            =   90
         ScaleHeight     =   4305
         ScaleWidth      =   6315
         TabIndex        =   8
         Top             =   300
         Width           =   6375
         Begin VB.Label labEnc 
            Alignment       =   2  'Center
            Caption         =   "D"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   5
            Left            =   90
            TabIndex        =   22
            Top             =   3840
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label labEnc 
            Alignment       =   2  'Center
            Caption         =   "B"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   4
            Left            =   90
            TabIndex        =   21
            Top             =   3210
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label labEnc 
            Alignment       =   2  'Center
            Caption         =   "P"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   90
            TabIndex        =   20
            Top             =   2580
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label labEnc 
            Alignment       =   2  'Center
            Caption         =   "M"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   19
            Top             =   1980
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label labEnc 
            Alignment       =   2  'Center
            Caption         =   "I"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   18
            Top             =   1320
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label labEnc 
            Alignment       =   2  'Center
            Caption         =   "L"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   17
            Top             =   690
            Width           =   225
         End
         Begin VB.Line objLine27 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   450
            Y1              =   3690
            Y2              =   4200
         End
         Begin VB.Line objLine22 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   680
            Y1              =   4170
            Y2              =   4170
         End
         Begin VB.Line objLine23 
            Index           =   0
            Visible         =   0   'False
            X1              =   660
            X2              =   660
            Y1              =   3690
            Y2              =   4170
         End
         Begin VB.Line objLine24 
            Index           =   0
            Visible         =   0   'False
            X1              =   660
            X2              =   895
            Y1              =   3690
            Y2              =   3690
         End
         Begin VB.Line objLine25 
            Index           =   0
            Visible         =   0   'False
            X1              =   660
            X2              =   895
            Y1              =   4170
            Y2              =   4170
         End
         Begin VB.Line objLine26 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   680
            Y1              =   3690
            Y2              =   3690
         End
         Begin VB.Line objLine18 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   3330
            Y2              =   3330
         End
         Begin VB.Line objLine17 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   3090
            Y2              =   3090
         End
         Begin VB.Line objLine19 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   3570
            Y2              =   3570
         End
         Begin VB.Line objLine20 
            Index           =   0
            Visible         =   0   'False
            X1              =   900
            X2              =   900
            Y1              =   3090
            Y2              =   3330
         End
         Begin VB.Line objLine21 
            Index           =   0
            Visible         =   0   'False
            X1              =   900
            X2              =   900
            Y1              =   3330
            Y2              =   3570
         End
         Begin VB.Line objLine16 
            Index           =   0
            Visible         =   0   'False
            X1              =   900
            X2              =   900
            Y1              =   2700
            Y2              =   2940
         End
         Begin VB.Line objLine15 
            Index           =   0
            Visible         =   0   'False
            X1              =   900
            X2              =   900
            Y1              =   2460
            Y2              =   2700
         End
         Begin VB.Line objLine14 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   2940
            Y2              =   2940
         End
         Begin VB.Line Line11 
            BorderStyle     =   3  'Dot
            X1              =   450
            X2              =   450
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line objLine12 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   2460
            Y2              =   2460
         End
         Begin VB.Line objLine13 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   2700
            Y2              =   2700
         End
         Begin VB.Line objLine11 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   450
            Y1              =   1830
            Y2              =   2310
         End
         Begin VB.Line objLine10 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   680
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Line objLine9 
            Index           =   0
            Visible         =   0   'False
            X1              =   660
            X2              =   895
            Y1              =   2310
            Y2              =   2310
         End
         Begin VB.Line objLine8 
            Index           =   0
            Visible         =   0   'False
            X1              =   660
            X2              =   895
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Line objLine7 
            Index           =   0
            Visible         =   0   'False
            X1              =   660
            X2              =   660
            Y1              =   1830
            Y2              =   2310
         End
         Begin VB.Line objLine6 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   680
            Y1              =   2310
            Y2              =   2310
         End
         Begin VB.Line objLine3 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line objLine5 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line objLine4 
            Index           =   0
            Visible         =   0   'False
            X1              =   900
            X2              =   900
            Y1              =   1200
            Y2              =   1680
         End
         Begin VB.Line Line26 
            BorderStyle     =   3  'Dot
            X1              =   450
            X2              =   450
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line25 
            BorderStyle     =   3  'Dot
            X1              =   900
            X2              =   900
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line24 
            BorderStyle     =   3  'Dot
            X1              =   1350
            X2              =   1350
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line23 
            BorderStyle     =   3  'Dot
            X1              =   1800
            X2              =   1800
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line22 
            BorderStyle     =   3  'Dot
            X1              =   2250
            X2              =   2250
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line21 
            BorderStyle     =   3  'Dot
            X1              =   2700
            X2              =   2700
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line20 
            BorderStyle     =   3  'Dot
            X1              =   3150
            X2              =   3150
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line19 
            BorderStyle     =   3  'Dot
            X1              =   3600
            X2              =   3600
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line18 
            BorderStyle     =   3  'Dot
            X1              =   4050
            X2              =   4050
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line17 
            BorderStyle     =   3  'Dot
            X1              =   4500
            X2              =   4500
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line16 
            BorderStyle     =   3  'Dot
            X1              =   4950
            X2              =   4950
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line15 
            BorderStyle     =   3  'Dot
            X1              =   5400
            X2              =   5400
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            X1              =   5850
            X2              =   5850
            Y1              =   2280
            Y2              =   4230
         End
         Begin VB.Line objLine1 
            Index           =   0
            Visible         =   0   'False
            X1              =   900
            X2              =   900
            Y1              =   570
            Y2              =   1050
         End
         Begin VB.Line objLine2 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   1050
            Y2              =   1050
         End
         Begin VB.Label labSeq 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   630
            TabIndex        =   9
            Top             =   90
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Line objLine 
            Index           =   0
            Visible         =   0   'False
            X1              =   450
            X2              =   900
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line10 
            BorderStyle     =   3  'Dot
            X1              =   5850
            X2              =   5850
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line14 
            BorderStyle     =   3  'Dot
            X1              =   5400
            X2              =   5400
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line13 
            BorderStyle     =   3  'Dot
            X1              =   4950
            X2              =   4950
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line9 
            BorderStyle     =   3  'Dot
            X1              =   4500
            X2              =   4500
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line6 
            BorderStyle     =   3  'Dot
            X1              =   4050
            X2              =   4050
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line8 
            BorderStyle     =   3  'Dot
            X1              =   3600
            X2              =   3600
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line7 
            BorderStyle     =   3  'Dot
            X1              =   3150
            X2              =   3150
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line5 
            BorderStyle     =   3  'Dot
            X1              =   2700
            X2              =   2700
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line4 
            BorderStyle     =   3  'Dot
            X1              =   2250
            X2              =   2250
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line3 
            BorderStyle     =   3  'Dot
            X1              =   1800
            X2              =   1800
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line2 
            BorderStyle     =   3  'Dot
            X1              =   1350
            X2              =   1350
            Y1              =   300
            Y2              =   2250
         End
         Begin VB.Line Line12 
            BorderStyle     =   3  'Dot
            X1              =   900
            X2              =   900
            Y1              =   300
            Y2              =   2250
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6465
      Left            =   6540
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command1 
         Caption         =   "Resume"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   1725
      End
      Begin VB.CommandButton butReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1725
      End
      Begin VB.CommandButton butStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   1725
      End
      Begin VB.CommandButton butExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3960
         Width           =   1725
      End
      Begin VB.CommandButton butStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encoding Techniques"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   6525
      Begin VB.CheckBox chTec 
         Caption         =   "&Differential Manchester"
         Height          =   255
         Index           =   5
         Left            =   4020
         TabIndex        =   16
         Top             =   630
         Width           =   1965
      End
      Begin VB.CheckBox chTec 
         Caption         =   "&Bipolar-AMI"
         Height          =   255
         Index           =   4
         Left            =   4020
         TabIndex        =   15
         Top             =   300
         Width           =   1305
      End
      Begin VB.CheckBox chTec 
         Caption         =   "&Pseudoternary"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   14
         Top             =   630
         Width           =   1515
      End
      Begin VB.CheckBox chTec 
         Caption         =   "&Manchester"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   13
         Top             =   300
         Width           =   1305
      End
      Begin VB.CheckBox chTec 
         Caption         =   "NRZ-&I"
         Height          =   255
         Index           =   1
         Left            =   630
         TabIndex        =   12
         Top             =   630
         Width           =   1305
      End
      Begin VB.CheckBox chTec 
         Caption         =   "NRZ-&L"
         Height          =   255
         Index           =   0
         Left            =   630
         TabIndex        =   11
         Top             =   300
         Value           =   1  'Checked
         Width           =   1305
      End
   End
   Begin VB.TextBox txtSeq 
      BeginProperty DataFormat 
         Type            =   5
         Format          =   "1111111111111111"
         HaveTrueFalseNull=   1
         TrueValue       =   "1"
         FalseValue      =   "0"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   7
      EndProperty
      Height          =   315
      Left            =   2730
      MaxLength       =   12
      TabIndex        =   1
      ToolTipText     =   "Enter a binary sequence here"
      Top             =   150
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Enter a binary sequence "
      Height          =   255
      Left            =   390
      TabIndex        =   0
      Top             =   180
      Width           =   2505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private varTec1 As String
Private varTec2 As String
Private varTec3 As String
Private varTec4 As String
Private varTec5 As String
Private varTec6 As String
Private counter As Integer
Private preState As String  ' previous state of NRZ-I
Private preStateM As String  ' previous state of Manchester
Private preStateP As String  ' previous state of Pseudoternary
Private preStateB As String 'Previous state of Bipolar-AMI
Private preStateD As String 'Previous state of Differential Manchester

Private Sub butExit_Click()
    End
End Sub

Private Sub butReset_Click()
    Dim i As Integer
    Timer1.Interval = 0
    counter = 1
    For i = 0 To 11
        'NRZ-L
        objLine(i).Visible = False
        objLine1(i).Visible = False
        objLine2(i).Visible = False
        'NRZ-I
        objLine3(i).Visible = False
        objLine4(i).Visible = False
        objLine5(i).Visible = False
        
        labSeq(i).Visible = False
        'Manchester
        objLine6(i).Visible = False
        objLine7(i).Visible = False
        objLine8(i).Visible = False
        objLine9(i).Visible = False
        objLine10(i).Visible = False
        objLine11(i).Visible = False
        'Pseudoternary
        objLine12(i).Visible = False
        objLine13(i).Visible = False
        objLine14(i).Visible = False
        objLine15(i).Visible = False
        objLine16(i).Visible = False
        'Bipolar-AMI
        objLine17(i).Visible = False
        objLine18(i).Visible = False
        objLine19(i).Visible = False
        objLine20(i).Visible = False
        objLine21(i).Visible = False
        'Differential Manchester
        objLine22(i).Visible = False
        objLine23(i).Visible = False
        objLine24(i).Visible = False
        objLine25(i).Visible = False
        objLine26(i).Visible = False
        objLine27(i).Visible = False
    Next i
    butStart.Enabled = True
End Sub

Private Sub butStart_Click()
    counter = 1
    Timer1.Interval = 500
    If Len(txtSeq.Text) <> 0 Then butStart.Enabled = False
End Sub

Private Sub butStop_Click()
    Timer1.Interval = 0
End Sub

Private Sub chTec_Click(Index As Integer)
    Select Case Index
        Case 0
            If chTec(0).Value = 1 Then
                varTec1 = "NRZ-L"
                labEnc(0).Visible = True
            Else
                varTec1 = ""
                labEnc(0).Visible = False
            End If
        Case 1
            If chTec(1).Value = 1 Then
                varTec2 = "NRZ-I"
                labEnc(1).Visible = True
            Else
                varTec2 = ""
                labEnc(1).Visible = False
            End If
        Case 2
            If chTec(2).Value = 1 Then
                varTec3 = "Manchester"
                labEnc(2).Visible = True
            Else
                varTec3 = ""
                labEnc(2).Visible = False
            End If
        Case 3
            If chTec(3).Value = 1 Then
                varTec4 = "Pseudoternary"
                labEnc(3).Visible = True
            Else
                varTec4 = ""
                labEnc(3).Visible = False
            End If
        Case 4
            If chTec(4).Value = 1 Then
                varTec5 = "Bipolar-AMI"
                labEnc(4).Visible = True
            Else
                varTec5 = ""
                labEnc(4).Visible = False
            End If
        Case 5
            If chTec(5).Value = 1 Then
                varTec6 = "Differential Manchester"
                labEnc(5).Visible = True
            Else
                varTec6 = ""
                labEnc(5).Visible = False
            End If

    End Select
End Sub

Private Sub Command1_Click()
    Timer1.Interval = 500
End Sub

Private Sub Form_Load()
    Dim i As Integer
    varTec1 = "NRZ-L"
    varTec2 = ""
    varTec3 = ""
    varTec4 = ""
    varTec5 = ""
    varTec6 = ""
    For i = 1 To 11
        'NRZ-L
        Load objLine(i)
        Load objLine1(i)
        Load objLine2(i)
        Load labSeq(i)   'sequence label
        ' NRZ-I
        Load objLine3(i)
        Load objLine4(i)
        Load objLine5(i)
        ' Manchester
        Load objLine6(i)
        Load objLine7(i)
        Load objLine8(i)
        Load objLine9(i)
        Load objLine10(i)
        Load objLine11(i)
        'Pseudoternary
        Load objLine12(i)
        Load objLine13(i)
        Load objLine14(i)
        Load objLine15(i)
        Load objLine16(i)
        'Bipolar-AMI
        Load objLine17(i)
        Load objLine18(i)
        Load objLine19(i)
        Load objLine20(i)
        Load objLine21(i)
        'Differential Manchester
        Load objLine22(i)
        Load objLine23(i)
        Load objLine24(i)
        Load objLine25(i)
        Load objLine26(i)
        Load objLine27(i)
        
        'NRZ-L
        objLine(i).X1 = objLine(i - 1).X2
        objLine(i).X2 = objLine(i - 1).X2 + 450
        objLine1(i).X1 = objLine1(i - 1).X1 + 450
        objLine1(i).X2 = objLine1(i - 1).X2 + 450
        objLine2(i).X1 = objLine2(i - 1).X2
        objLine2(i).X2 = objLine2(i - 1).X2 + 450
        
        labSeq(i).Left = labSeq(i - 1).Left + 450
        'NRZ-I
        objLine3(i).X1 = objLine3(i - 1).X2
        objLine3(i).X2 = objLine3(i - 1).X2 + 450
        objLine4(i).X1 = objLine4(i - 1).X1 + 450
        objLine4(i).X2 = objLine4(i - 1).X2 + 450
        objLine5(i).X1 = objLine5(i - 1).X2
        objLine5(i).X2 = objLine5(i - 1).X2 + 450
        'Manchester
        objLine6(i).X1 = objLine5(i - 1).X2
        objLine6(i).X2 = objLine6(i - 1).X2 + 450
        objLine7(i).X1 = objLine7(i - 1).X1 + 450
        objLine7(i).X2 = objLine7(i - 1).X2 + 450
        objLine8(i).X1 = objLine8(i - 1).X1 + 450
        objLine8(i).X2 = objLine8(i - 1).X2 + 450
        objLine9(i).X1 = objLine9(i - 1).X1 + 450
        objLine9(i).X2 = objLine9(i - 1).X2 + 450
        objLine10(i).X1 = objLine5(i - 1).X2
        objLine10(i).X2 = objLine10(i - 1).X2 + 450
        objLine11(i).X1 = objLine11(i - 1).X1 + 450
        objLine11(i).X2 = objLine11(i - 1).X2 + 450
        'Pseudoternary
        objLine12(i).X1 = objLine12(i - 1).X2
        objLine12(i).X2 = objLine12(i - 1).X2 + 450
        objLine13(i).X1 = objLine13(i - 1).X2
        objLine13(i).X2 = objLine13(i - 1).X2 + 450
        objLine14(i).X1 = objLine14(i - 1).X2
        objLine14(i).X2 = objLine14(i - 1).X2 + 450
        objLine15(i).X1 = objLine15(i - 1).X1 + 450
        objLine15(i).X2 = objLine15(i - 1).X2 + 450
        objLine16(i).X1 = objLine16(i - 1).X1 + 450
        objLine16(i).X2 = objLine16(i - 1).X2 + 450
        'Bipolar-AMI
        objLine17(i).X1 = objLine17(i - 1).X2
        objLine17(i).X2 = objLine17(i - 1).X2 + 450
        objLine18(i).X1 = objLine18(i - 1).X2
        objLine18(i).X2 = objLine18(i - 1).X2 + 450
        objLine19(i).X1 = objLine19(i - 1).X2
        objLine19(i).X2 = objLine19(i - 1).X2 + 450
        objLine20(i).X1 = objLine20(i - 1).X1 + 450
        objLine20(i).X2 = objLine20(i - 1).X2 + 450
        objLine21(i).X1 = objLine21(i - 1).X1 + 450
        objLine21(i).X2 = objLine21(i - 1).X2 + 450
        'Differential Manchester
        objLine22(i).X1 = objLine5(i - 1).X2
        objLine22(i).X2 = objLine22(i - 1).X2 + 450
        objLine23(i).X1 = objLine23(i - 1).X1 + 450
        objLine23(i).X2 = objLine23(i - 1).X2 + 450
        objLine24(i).X1 = objLine24(i - 1).X1 + 450
        objLine24(i).X2 = objLine24(i - 1).X2 + 450
        objLine25(i).X1 = objLine25(i - 1).X1 + 450
        objLine25(i).X2 = objLine25(i - 1).X2 + 450
        objLine26(i).X1 = objLine5(i - 1).X2
        objLine26(i).X2 = objLine26(i - 1).X2 + 450
        objLine27(i).X1 = objLine27(i - 1).X1 + 450
        objLine27(i).X2 = objLine27(i - 1).X2 + 450

        
    Next i

End Sub

Private Sub Timer1_Timer()
   If counter > Len(txtSeq.Text) Then
        Timer1.Interval = 0
        Exit Sub
   End If
   simulate (varTec1)
   simulate (varTec2)
   simulate (varTec3)
   simulate (varTec4)
   simulate (varTec5)
   simulate (varTec6)
   counter = counter + 1
End Sub

Private Sub simulate(tec As String)
    Select Case tec
        Case "NRZ-L"
            Call NRZLSim(counter - 1)
        Case "NRZ-I"
            Call NRZISim(counter - 1)
        Case "Manchester"
            Call ManchesterSim(counter - 1)
        Case "Pseudoternary"
            Call PseudoternarySim(counter - 1)
        Case "Bipolar-AMI"
            Call BamiSim(counter - 1)
        Case "Differential Manchester"
            Call difManchesterSim(counter - 1)
   End Select
End Sub

Private Function getInput(ind As Integer) As String
    getInput = Mid$(txtSeq.Text, ind, 1)
End Function

Private Sub txtSeq_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = 49 Or KeyAscii = 48 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub NRZLSim(ByVal ind As Integer)
    labSeq(ind).Caption = getInput(ind + 1)
    labSeq(ind).Visible = True
    If getInput(ind + 1) = "0" Then
        objLine(ind).Visible = True
    Else
        objLine2(ind).Visible = True
    End If
    If ind <> 0 Then
        If getInput(ind) <> getInput(ind + 1) Then
            objLine1(ind - 1).Visible = True
        End If
    End If
End Sub

Private Sub NRZISim(ByVal ind As Integer)
    labSeq(ind).Caption = getInput(ind + 1)
    labSeq(ind).Visible = True
    If ind = 0 Then
        objLine5(ind).Visible = True
        preState = "down"
    Else
        If getInput(ind + 1) = "1" Then
            objLine4(ind - 1).Visible = True
            If preState = "up" Then
                objLine5(ind).Visible = True
                preState = "down"
            Else
                objLine3(ind).Visible = True
                preState = "up"
            End If
        Else
            If preState = "up" Then
                objLine3(ind).Visible = True
            Else
                objLine5(ind).Visible = True
            End If
        End If
    End If
End Sub

Private Sub ManchesterSim(ByVal ind As Integer)
    labSeq(ind).Caption = getInput(ind + 1)
    labSeq(ind).Visible = True
    If ind = 0 Then
        If getInput(ind + 1) = "1" Then
            objLine6(ind).Visible = True
            objLine7(ind).Visible = True
            objLine8(ind).Visible = True
            preStateM = "1"
        Else
            objLine10(ind).Visible = True
            objLine7(ind).Visible = True
            objLine9(ind).Visible = True
            preStateM = "0"
        End If
    Else
        If getInput(ind + 1) = preStateM Then
            objLine11(ind).Visible = True
        End If
        If getInput(ind + 1) = "1" Then
            objLine6(ind).Visible = True
            objLine7(ind).Visible = True
            objLine8(ind).Visible = True
        Else
            objLine10(ind).Visible = True
            objLine7(ind).Visible = True
            objLine9(ind).Visible = True
        End If
        preStateM = getInput(ind + 1)
    End If
End Sub

Private Sub PseudoternarySim(ByVal ind As Integer)
    labSeq(ind).Caption = getInput(ind + 1)
    labSeq(ind).Visible = True
    If ind = 0 Then
        If getInput(ind + 1) = "1" Then
            objLine13(ind).Visible = True
            preStateP = "down-mid"
        Else
            objLine12(ind).Visible = True
            preStateP = "up"
        End If
    Else
        Select Case preStateP
            Case "up-mid"
                If getInput(ind + 1) = "1" Then
                    objLine13(ind).Visible = True
                Else
                    objLine16(ind - 1).Visible = True
                    objLine14(ind).Visible = True
                    preStateP = "down"
                End If
            Case "up"
                If getInput(ind + 1) = "1" Then
                    objLine15(ind - 1).Visible = True
                    objLine13(ind).Visible = True
                    preStateP = "up-mid"
                Else
                    objLine15(ind - 1).Visible = True
                    objLine16(ind - 1).Visible = True
                    objLine14(ind).Visible = True
                    preStateP = "down"
                End If
            Case "down-mid"
                If getInput(ind + 1) = "1" Then
                    objLine13(ind).Visible = True
                Else
                    objLine15(ind - 1).Visible = True
                    objLine12(ind).Visible = True
                    preStateP = "up"
                End If
            Case "down"
                If getInput(ind + 1) = "1" Then
                    objLine16(ind - 1).Visible = True
                    objLine13(ind).Visible = True
                    preStateP = "down-mid"
                Else
                    objLine15(ind - 1).Visible = True
                    objLine16(ind - 1).Visible = True
                    objLine12(ind).Visible = True
                    preStateP = "up"
                End If
        End Select
    End If
End Sub
Private Sub BamiSim(ByVal ind As Integer)
    labSeq(ind).Caption = getInput(ind + 1)
    labSeq(ind).Visible = True
    If ind = 0 Then
        If getInput(ind + 1) = "0" Then
            objLine18(ind).Visible = True
            preStateB = "down-mid"
        Else
            objLine17(ind).Visible = True
            preStateB = "up"
        End If
    Else
        Select Case preStateB
            Case "up-mid"
                If getInput(ind + 1) = "0" Then
                    objLine18(ind).Visible = True
                Else
                    objLine21(ind - 1).Visible = True
                    objLine19(ind).Visible = True
                    preStateB = "down"
                End If
            Case "up"
                If getInput(ind + 1) = "0" Then
                    objLine20(ind - 1).Visible = True
                    objLine18(ind).Visible = True
                    preStateB = "up-mid"
                Else
                    objLine20(ind - 1).Visible = True
                    objLine21(ind - 1).Visible = True
                    objLine19(ind).Visible = True
                    preStateB = "down"
                End If
            Case "down-mid"
                If getInput(ind + 1) = "0" Then
                    objLine18(ind).Visible = True
                Else
                    objLine20(ind - 1).Visible = True
                    objLine17(ind).Visible = True
                    preStateB = "up"
                End If
            Case "down"
                If getInput(ind + 1) = "0" Then
                    objLine21(ind - 1).Visible = True
                    objLine18(ind).Visible = True
                    preStateB = "down-mid"
                Else
                    objLine20(ind - 1).Visible = True
                    objLine21(ind - 1).Visible = True
                    objLine17(ind).Visible = True
                    preStateB = "up"
                End If
        End Select
    End If
End Sub

Private Sub difManchesterSim(ByVal ind As Integer)
    labSeq(ind).Caption = getInput(ind + 1)
    labSeq(ind).Visible = True
    If ind = 0 Then
        objLine26(ind).Visible = True
        objLine23(ind).Visible = True
        objLine25(ind).Visible = True
        preStateD = "up-down"
    Else
        If getInput(ind + 1) = "1" Then
            Select Case preStateD
            Case "up-down"
                objLine26(ind).Visible = True
                objLine23(ind).Visible = True
                objLine25(ind).Visible = True
            Case "down-up"
                objLine22(ind).Visible = True
                objLine23(ind).Visible = True
                objLine24(ind).Visible = True
            End Select
            objLine27(ind).Visible = True
        Else
            Select Case preStateD
            Case "up-down"
                objLine22(ind).Visible = True
                objLine23(ind).Visible = True
                objLine24(ind).Visible = True
                preStateD = "down-up"
            Case "down-up"
                objLine26(ind).Visible = True
                objLine23(ind).Visible = True
                objLine25(ind).Visible = True
                preStateD = "up-down"
            End Select
        End If
    End If
End Sub

