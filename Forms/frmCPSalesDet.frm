VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPSalesDet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1350
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   615
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2381
      BackColor       =   12632256
      Begin VB.TextBox txtSpecify 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   675
         Width           =   4500
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
         Left            =   285
         TabIndex        =   11
         Tag             =   "et0;fb0"
         Top             =   90
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   930
         Index           =   2
         Left            =   180
         Top             =   195
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   1
         Top             =   420
         Width           =   990
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1635
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   2040
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2884
      Begin VB.OptionButton optPresentation 
         Caption         =   "Nokia"
         Height          =   210
         Index           =   5
         Left            =   390
         TabIndex        =   16
         Tag             =   "et0;fb0"
         Top             =   1275
         Width           =   1260
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "LG"
         Height          =   210
         Index           =   4
         Left            =   375
         TabIndex        =   5
         Tag             =   "et0;fb0"
         Top             =   1065
         Width           =   1260
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Samsung"
         Height          =   210
         Index           =   3
         Left            =   375
         TabIndex        =   4
         Tag             =   "et0;fb0"
         Top             =   795
         Width           =   1260
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3030
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   510
         Width           =   1890
      End
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3030
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   840
         Width           =   1890
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "E'touch"
         Height          =   210
         Index           =   2
         Left            =   375
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   555
         Width           =   1260
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Sony Ericson"
         Height          =   210
         Index           =   1
         Left            =   375
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   315
         Value           =   -1  'True
         Width           =   1260
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
         Left            =   2130
         TabIndex        =   13
         Tag             =   "et0;fb0"
         Top             =   75
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Presentation"
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
         Index           =   2
         Left            =   270
         TabIndex        =   12
         Tag             =   "et0;fb0"
         Top             =   75
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   0
         Left            =   2205
         TabIndex        =   15
         Top             =   585
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Thru"
         Height          =   195
         Index           =   1
         Left            =   2205
         TabIndex        =   14
         Top             =   870
         Width           =   825
      End
      Begin VB.Shape Shape1 
         Height          =   1365
         Index           =   1
         Left            =   2025
         Top             =   150
         Width           =   3000
      End
      Begin VB.Shape Shape1 
         Height          =   1365
         Index           =   0
         Left            =   180
         Top             =   150
         Width           =   1680
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5535
      TabIndex        =   10
      Top             =   1860
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
      Picture         =   "frmCPSalesDet.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   5535
      TabIndex        =   8
      Top             =   600
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
      Picture         =   "frmCPSalesDet.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5535
      TabIndex        =   9
      Top             =   1230
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
      Picture         =   "frmCPSalesDet.frx":0EF4
   End
End
Attribute VB_Name = "frmCPSalesDet"
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

Property Get Presentation() As Integer
   '  0 = Summarized
   '  1 = Detailed
   Presentation = pnPresntn
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

'   optPresentation(0).Value = True
'   optPresentation(1).Value = False
   
   txtSpecify.Text = ""
   psBranchCd = ""
   pbCancel = True
   pnPresntn = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set loSkin = Nothing
End Sub

Private Sub SearchBranch(ByVal lbEqual As Boolean, Optional Value As Variant)
   Dim lsBrowse As String, lsSelected() As String, lsSQL As String
   Dim lrs As ADODB.Recordset

   lsSQL = " Select " _
               & "  sBranchCd" _
               & ", sBranchNm" _
            & " From Branch" _
            & " Where cRecdStat = " & strParm(xeRecStateActive) _
               & IIf(loAppDrivr.UserLevel > xeManager, "", " AND sBranchCd = " & strParm(loAppDrivr.BranchCode))

   If Not lbEqual Then
      If Not IsMissing(Value) Then lsSQL = lsSQL & " And sBranchNm LIKE " & strParm(Value & "%")
   Else
      If Not IsMissing(Value) Then lsSQL = lsSQL & " And sBranchNm = " & strParm(Value)
   End If
            
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, loAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lrs.EOF Then
      txtSpecify.Text = ""
      psBranchCd = ""
   ElseIf lrs.RecordCount = 1 Then
      txtSpecify.Text = lrs("sBranchNm")
      psBranchCd = lrs("sBranchCd")
   Else
      lsBrowse = KwikBrowse(loAppDrivr, lrs _
                              , "sBranchCd»sBranchNm" _
                              , "Code»Branch Name" _
                              , "@»@", False)

      If lsBrowse <> "" Then
         lsSelected = Split(lsBrowse, "»")
         txtSpecify.Text = lsSelected(1)
         psBranchCd = lsSelected(0)
      Else
         If psBranchCd = "" Then txtSpecify.Text = ""
         txtSpecify.Text = txtSpecify.Tag
      End If
   End If
   
   txtSpecify.Tag = txtSpecify.Text
   txtSpecify.SelStart = 0
   txtSpecify.SelLength = Len(txtSpecify.Text)
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

Private Sub optPresentation_Click(Index As Integer)
   pnPresntn = Index
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

Private Sub txtSpecify_GotFocus()
   With txtSpecify
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   lbSearch = True
End Sub

Private Sub txtSpecify_KeyDown(KeyCode As Integer, Shift As Integer)
   With txtSpecify
      If KeyCode = vbKeyF3 Then
         If .Text = .Tag Then Exit Sub
         SearchBranch False, .Text
         KeyCode = 0
      End If
   End With
End Sub

Private Sub txtSpecify_Validate(Cancel As Boolean)
   With txtSpecify
      If .Text = .Tag Then Exit Sub

      SearchBranch False, .Text
      .Tag = .Text
   End With
End Sub


