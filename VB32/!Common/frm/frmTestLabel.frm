VERSION 5.00
Begin VB.Form frmTestLabel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "jeffLabel Test Form"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmTestLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTestLabel.frx":27A2
   ScaleHeight     =   5880
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   180
      TabIndex        =   2
      Top             =   4980
      Width           =   1575
   End
   Begin jeffLabelTestProject.jeffLabel lblJeffLabel1 
      Height          =   5565
      Index           =   1
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9816
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      BorderStyle     =   1
      Alignment       =   1
      AutoSize        =   -1  'True
      Caption         =   $"frmTestLabel.frx":89EF
      FontName        =   "Comic Sans MS"
      FontSize        =   12
      WordWrap        =   -1  'True
      ShadowColor     =   16711935
      OffsetX         =   1
      OffsetY         =   1
   End
   Begin jeffLabelTestProject.jeffLabel lblJeffLabel1 
      Height          =   2655
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   4683
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      BorderStyle     =   1
      Alignment       =   2
      AutoSize        =   -1  'True
      Caption         =   "Jeff Label"
      FontBold        =   -1  'True
      FontName        =   "Comic Sans MS"
      FontSize        =   48
      MouseIcon       =   "frmTestLabel.frx":8C7F
      MousePointer    =   99
      WordWrap        =   -1  'True
      ShadowColor     =   33023
      OffsetX         =   -5
      OffsetY         =   -5
      URL             =   "http://members.tripod.com/thefrogprince/"
   End
End
Attribute VB_Name = "frmTestLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCommand1_Click()
    Static lClicks As Long
    Static sOrigText0 As String
    Static sOrigText1 As String
    lClicks = lClicks + 1
    Select Case True
        Case lClicks = 1
            sOrigText0 = Me.lblJeffLabel1(0).Caption
            sOrigText1 = Me.lblJeffLabel1(1).Caption
            Me.lblJeffLabel1(0).Caption = "This is a test"
            Me.lblJeffLabel1(1).Caption = "One of the benefits you will notice of using the DrawText api function to paint the label manually on the control's parent is that the text properly centers and aligns. =)  To see how VB labels perform this task, click the test button again. =)"
        Case lClicks = 2
            Me.lblJeffLabel1(0).BackStyle = elbOpaque
            Me.lblJeffLabel1(1).BackStyle = elbOpaque
            Me.lblJeffLabel1(1).Caption = "As you can see... the only thing the standard VB label control can do properly is left alignment.  It fails the test on both Center alignment and Right alignment.  A bit shocking.  Anyways... I typically paint my labels transparent anyways, so this works for me. =)  Click the Test button again to reset all of the values."
        Case lClicks = 3
            Me.lblJeffLabel1(0).BackStyle = elbTransparent
            Me.lblJeffLabel1(1).BackStyle = elbTransparent
            Me.lblJeffLabel1(0).Caption = sOrigText0
            Me.lblJeffLabel1(1).Caption = sOrigText1
            lClicks = 0
    End Select
End Sub




