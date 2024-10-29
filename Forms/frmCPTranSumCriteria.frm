VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPTranSumCriteria 
   BorderStyle     =   0  'None
   Caption         =   "frmCPTranSumCriteria"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1425
      Index           =   0
      Left            =   570
      Tag             =   "wt0;fb0"
      Top             =   360
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   2514
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   405
         Width           =   1845
      End
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   825
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   3
         Top             =   840
         Width           =   825
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
         Left            =   195
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   15
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   1170
         Index           =   1
         Left            =   120
         Top             =   105
         Width           =   2985
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   4035
      TabIndex        =   5
      Top             =   1635
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
      Picture         =   "frmCPTranSumCriteria.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   4035
      TabIndex        =   6
      Top             =   375
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
      Picture         =   "frmCPTranSumCriteria.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   4035
      TabIndex        =   7
      Top             =   1005
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
      Picture         =   "frmCPTranSumCriteria.frx":0EF4
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1485
      Index           =   1
      Left            =   570
      Tag             =   "wt0;fb0"
      Top             =   1800
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   2619
      Begin VB.OptionButton optSummary 
         Caption         =   "Model"
         Height          =   270
         Index           =   1
         Left            =   1725
         TabIndex        =   11
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Brand"
         Height          =   270
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   345
         Width           =   1125
      End
      Begin VB.TextBox txtSummary 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   225
         TabIndex        =   8
         Top             =   840
         Width           =   2805
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Summary"
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
         Index           =   4
         Left            =   195
         TabIndex        =   9
         Tag             =   "et0;fb0"
         Top             =   30
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   1230
         Index           =   2
         Left            =   120
         Top             =   120
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmCPTranSumCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loAppDrivr As clsAppDriver
Private loSkin As clsFormSkin

Private pbCancel As Boolean
Private pnPresntn As Integer
Private psSummary As String
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

Property Get Summary() As String
   Summary = psSummary
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer

   Select Case Index
   Case 0, 1
      pbCancel = Index = 1
      Me.Hide
   Case 2
      If lbSearch Then
         SearchSummary False
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

   txtSummary.Text = ""
  
  
   pbCancel = True
  End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set loSkin = Nothing
End Sub

Private Sub SearchSummary(ByVal lbEqual As Boolean, Optional Value As Variant)
   Dim lsBrowse As String, lsSelected() As String, lsSQL As String
   Dim lrs As ADODB.Recordset

   lsSQL = " Select " & _
               " sBrandIdx" & _
               " sBrandNme" & _
            " FROM CP_Brand" & _
            " WHERE "

   If Not lbEqual Then
      If Not IsMissing(Value) Then lsSQL = lsSQL & "  sBrandNme LIKE " & strParm(Value & "%")
   Else
      If Not IsMissing(Value) Then lsSQL = lsSQL & "  sBrandNme = " & strParm(Value)
   End If
            
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, loAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lrs.EOF Then
      txtSummary.Text = ""
      psSummary = ""
   ElseIf lrs.RecordCount = 1 Then
      txtSummary.Text = lrs("sBrandNme")
      psSummary = lrs("sBrandNme")
   Else
      lsBrowse = KwikBrowse(loAppDrivr, lrs _
                              , "sBrandIDx»sBrandNme" _
                              , "ID»Brand Name" _
                              , "@»@", False)

      If lsBrowse <> "" Then
         lsSelected = Split(lsBrowse, "»")
         txtSummary.Text = lsSelected(1)
         psSummary = lsSelected(0)
      Else
         If psSummary = "" Then txtSummary.Text = ""
         txtSummary.Text = txtSummary.Tag
      End If
   End If
   
   txtSummary.Tag = txtSummary.Text
   txtSummary.SelStart = 0
   txtSummary.SelLength = Len(txtSummary.Text)
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

Private Sub txtSummary_GotFocus()
   With txtSummary
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   lbSearch = True
End Sub

Private Sub txtSummary_KeyDown(KeyCode As Integer, Shift As Integer)
   With txtSummary
      If KeyCode = vbKeyF3 Then
         If .Text = .Tag Then Exit Sub
         SearchSummary False, .Text
         KeyCode = 0
      End If
   End With
End Sub

Private Sub txtSummary_Validate(Cancel As Boolean)
   With txtSummary
      If .Text = .Tag Then Exit Sub

      SearchSummary False, .Text
      .Tag = .Text
   End With
End Sub




