VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmDateBranchCriteria 
   BorderStyle     =   0  'None
   Caption         =   "Transaction Summary"
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2730
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame2 
      Height          =   585
      Index           =   1
      Left            =   240
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1032
      BackColor       =   12632256
      Begin VB.TextBox txtSpecify 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1065
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   120
         Width           =   3675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
         Height          =   195
         Index           =   6
         Left            =   345
         TabIndex        =   10
         Top             =   945
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   990
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1410
      Index           =   0
      Left            =   240
      Tag             =   "wt0;fb0"
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2487
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   750
         Width           =   1890
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   390
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Date"
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
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   45
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   1125
         Index           =   1
         Left            =   180
         Top             =   120
         Width           =   4845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Thru"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   465
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5535
      TabIndex        =   8
      Top             =   1815
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
      Picture         =   "frmDateBranchCriteria.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   5535
      TabIndex        =   6
      Top             =   555
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
      Picture         =   "frmDateBranchCriteria.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5535
      TabIndex        =   7
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDateBranchCriteria.frx":0EF4
   End
End
Attribute VB_Name = "frmDateBranchCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loAppDrivr As clsAppDriver
Private loSkin As clsFormSkin

Private pbCancel As Boolean
Private pnPresntn As Integer
Private psBranchCd As String

Dim lbSearch As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set loAppDrivr = oAppDriver
   psBranchCd = loAppDrivr.BranchCode
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancel
End Property

Property Get DateFrom() As Date
   DateFrom = CDate(txtDateFrom.Text)
End Property

Property Get DateThru() As Date
   DateThru = CDate(txtDateThru.Text)
End Property

Property Get BranchCd() As String
   BranchCd = psBranchCd
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer

   Select Case Index
   Case 0, 1
      pbCancel = Index = 1
      Me.Hide
   Case 2
      If lbSearch Then
         SearchBranch False
      End If
   End Select
End Sub

Private Sub Form_Load()
   Set loSkin = New clsFormSkin
   Set loSkin.AppDriver = loAppDrivr
   Set loSkin.Form = Me
   loSkin.ApplySkin xeFormTransDetail

   txtDateFrom = Format(DateAdd("m", -1, loAppDrivr.ServerDate), "MMM DD, YYYY")
   txtDateThru = Format(loAppDrivr.ServerDate, "MMM DD, YYYY")
   
   txtSpecify(0).Text = ""
   psBranchCd = ""
   
   If Not loAppDrivr.IsMainOffice Then
      SearchBranch True, loAppDrivr.BranchName
   End If
   
   txtSpecify(0).Enabled = txtSpecify(0).Text = ""
   
   pbCancel = True
   pnPresntn = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set loSkin = Nothing
End Sub

Private Sub SearchBranch(ByVal lbEqual As Boolean, Optional Value As Variant)
   Dim lsBrowse As String, lsSelected() As String, lsSQL As String
   Dim lrs As ADODB.Recordset

      lsSQL = " Select " & _
                  "  a.sBranchCd" & _
                  ", a.sBranchNm" & _
               " From Branch a" & _
               ", Branch_Others b" & _
               " Where a.sBranchCd = b.sBranchCd" & _
               " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
               " AND b.cDivision = " & strParm("0")

      If Not lbEqual Then
         If Not IsMissing(Value) Then lsSQL = lsSQL & " And a.sBranchNm LIKE " & strParm(Value & "%")
      Else
         If Not IsMissing(Value) Then lsSQL = lsSQL & " And a.sBranchNm = " & strParm(Value)
      End If
            
      Set lrs = New ADODB.Recordset
      lrs.Open lsSQL, loAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

      If lrs.EOF Then
         txtSpecify(0).Text = ""
         psBranchCd = ""
      ElseIf lrs.RecordCount = 1 Then
         txtSpecify(0).Text = lrs("sBranchNm")
         psBranchCd = lrs("sBranchCd")
      Else
         lsBrowse = KwikBrowse(loAppDrivr, lrs _
            , "sBranchCd»sBranchNm" _
            , "Code»Branch Name" _
            , "@»@", False)

      If lsBrowse <> "" Then
         lsSelected = Split(lsBrowse, "»")
         txtSpecify(0).Text = lsSelected(1)
         psBranchCd = lsSelected(0)
      Else
         If psBranchCd = "" Then txtSpecify(0).Text = ""
            txtSpecify(0).Text = txtSpecify(0).Tag
         End If
      End If
   
      txtSpecify(0).Tag = txtSpecify(0).Text
      txtSpecify(0).SelStart = 0
      txtSpecify(0).SelLength = Len(txtSpecify(0).Text)
      lrs.Close

      Set lrs = Nothing
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


Private Sub txtDateFrom_GotFocus()
   With txtDateFrom
      .Text = Format(.Text, "MM/DD/YYYY")

      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   lbSearch = False
End Sub

Private Sub txtDateFrom_Validate(Cancel As Boolean)
   With txtDateFrom
      If Not IsDate(.Text) Then
         .Text = Format(DateAdd("m", -1, loAppDrivr.ServerDate), "MMM DD, YYYY")
      Else
         .Text = Format(.Text, "MMM DD, YYYY")
      End If
   End With
End Sub

Private Sub txtDateThru_GotFocus()
   With txtDateThru
      .Text = Format(.Text, "MM/DD/YYYY")

      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   lbSearch = False
End Sub

Private Sub txtDateThru_Validate(Cancel As Boolean)
   With txtDateThru
      If Not IsDate(.Text) Then
         .Text = Format(loAppDrivr.ServerDate, "MMM DD, YYYY")
      Else
         .Text = Format(.Text, "MMM DD, YYYY")
      End If
   End With
End Sub


Private Sub txtSpecify_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 0
         SearchBranch False, txtSpecify(Index).Text
         KeyCode = 0
      End Select
   End If
End Sub
