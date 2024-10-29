VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmColCriteria 
   BorderStyle     =   0  'None
   Caption         =   "Transaction Summary"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1875
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   3307
      Begin VB.OptionButton Option1 
         Caption         =   "Area"
         Height          =   225
         Index           =   0
         Left            =   705
         TabIndex        =   0
         Tag             =   "et0;fb0"
         Top             =   555
         Width           =   1665
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Office Collection"
         Height          =   225
         Index           =   1
         Left            =   705
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   885
         Width           =   1665
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No Payment"
         Height          =   225
         Index           =   2
         Left            =   705
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   " Summarized By "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   315
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         Height          =   1440
         Left            =   195
         Top             =   210
         Width           =   2655
      End
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   3390
      TabIndex        =   4
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmColCriteria.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   3390
      TabIndex        =   5
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmColCriteria.frx":077A
   End
End
Attribute VB_Name = "frmColCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loSkin As clsFormSkin
Private loAppDrivr As clsAppDriver

Dim pbCancelled As Boolean
Dim pnPresentation As Integer

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set loAppDrivr = oAppDriver
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Property Get Presentation() As Integer
   Presentation = pnPresentation
End Property

Private Sub cmdButton_Click(Index As Integer)
   pbCancelled = Index = 1
   Me.Hide
End Sub

Private Sub Form_Load()
   Set loSkin = New clsFormSkin
   Set loSkin.AppDriver = loAppDrivr
   Set loSkin.Form = Me
   loSkin.ApplySkin xeFormTransDetail
   
   pbCancelled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set loSkin = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Option1_Click(Index As Integer)
   pnPresentation = Index
End Sub
