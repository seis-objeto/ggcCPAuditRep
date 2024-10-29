VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmTransferSummary 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1605
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   2010
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2831
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2250
         TabIndex        =   3
         Top             =   825
         Width           =   2730
      End
      Begin VB.TextBox txtBranch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2250
         TabIndex        =   2
         Top             =   465
         Width           =   2730
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   960
         Left            =   120
         Tag             =   "wt0;fb0"
         Top             =   285
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1693
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.OptionButton optStat 
            Caption         =   "ALL"
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   20
            Tag             =   "et0;fb0"
            Top             =   75
            Width           =   1200
         End
         Begin VB.OptionButton optStat 
            Caption         =   "UNPOSTED"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   17
            Tag             =   "et0;fb0"
            Top             =   765
            Width           =   1200
         End
         Begin VB.OptionButton optStat 
            Caption         =   "POSTED"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   16
            Tag             =   "et0;fb0"
            Top             =   420
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin VB.Line Line2 
         X1              =   1455
         X2              =   1455
         Y1              =   315
         Y2              =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "Category"
         Height          =   240
         Index           =   1
         Left            =   1545
         TabIndex        =   15
         Tag             =   "et0;fb0"
         Top             =   870
         Width           =   1050
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
         TabIndex        =   8
         Tag             =   "et0;fb0"
         Top             =   30
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   1320
         Index           =   2
         Left            =   60
         Top             =   135
         Width           =   4995
      End
      Begin VB.Label Label2 
         Caption         =   "Branch"
         Height          =   165
         Index           =   0
         Left            =   1545
         TabIndex        =   7
         Tag             =   "et0;fb0"
         Top             =   480
         Width           =   1020
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1440
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2540
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3450
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   390
         Width           =   1530
      End
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3450
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1530
      End
      Begin VB.OptionButton optType 
         Caption         =   "Transfer"
         Height          =   210
         Index           =   0
         Left            =   1275
         TabIndex        =   10
         Tag             =   "et0;fb0"
         Top             =   435
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optType 
         Caption         =   "Received"
         Height          =   210
         Index           =   1
         Left            =   1275
         TabIndex        =   9
         Tag             =   "et0;fb0"
         Top             =   780
         Width           =   1215
      End
      Begin xrControl.xrFrame xrFrame4 
         Height          =   840
         Left            =   135
         Tag             =   "wt0;fb0"
         Top             =   300
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1482
         BackColor       =   12632256
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.OptionButton optPresentation 
            Caption         =   "Detail"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   19
            Tag             =   "et0;fb0"
            Top             =   495
            Width           =   765
         End
         Begin VB.OptionButton optPresentation 
            Caption         =   "Summary"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   18
            Tag             =   "et0;fb0"
            Top             =   150
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Line Line1 
         X1              =   1245
         X2              =   1245
         Y1              =   285
         Y2              =   1215
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
         Left            =   2655
         TabIndex        =   11
         Tag             =   "et0;fb0"
         Top             =   60
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
         Left            =   150
         TabIndex        =   12
         Tag             =   "et0;fb0"
         Top             =   60
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   0
         Left            =   2670
         TabIndex        =   14
         Top             =   465
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   1
         Left            =   2670
         TabIndex        =   13
         Top             =   750
         Width           =   825
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Index           =   0
         Left            =   75
         Top             =   105
         Width           =   2445
      End
      Begin VB.Shape Shape1 
         Height          =   1185
         Index           =   1
         Left            =   2565
         Top             =   120
         Width           =   2490
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5535
      TabIndex        =   6
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
      Picture         =   "frmTransferSummary.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   5535
      TabIndex        =   4
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
      Picture         =   "frmTransferSummary.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5535
      TabIndex        =   5
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
      Picture         =   "frmTransferSummary.frx":0EF4
   End
End
Attribute VB_Name = "frmTransferSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loAppDrivr As clsAppDriver
Private loBranch As clsBranch
Private oSkin As clsFormSkin
Private pbCancel As Boolean
Private pnPresntn As Integer
Private psSummary As String
Private pnSummary As Integer

Private pnRcvdTyp As Integer
Private psCategID As String
Private pnStatus As Integer
Private pnType As Integer

Private p_sBranchCd As String
Private p_sBranchNm As String
Private p_sAddressx As String

Dim lbSearch As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set loAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Get BranchName() As String
   BranchName = p_sBranchNm
End Property

Property Let Branch(Value As String)
   p_sBranchCd = Value
End Property

Property Get BranchAddress() As String
   BranchAddress = p_sAddressx
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancel
End Property

Property Get DateFrom() As Date
   DateFrom = CDate(txtDateFrom)
End Property

Property Get DateThru() As Date
   DateThru = CDate(txtDateThru)
End Property

Property Get Specify() As String
   Specify = psSummary
End Property

Property Get Summary() As Integer
   Summary = pnSummary
End Property

Property Get Presentation() As Integer
   '  0 = Tranfer
   '  1 = Delivered
   Presentation = pnPresntn
End Property

Property Get Category() As String
   Category = psCategID
End Property

Property Get Status() As Integer
   Status = pnStatus
End Property

Property Let Category(lsCategory As String)
   lsCategory = psCategID
End Property

Property Get ReceivedType() As Integer
   ReceivedType = pnRcvdTyp
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer

   Select Case Index
   Case 0
      pbCancel = False
      Me.Hide
   Case 1
      pbCancel = True
      Me.Hide
   Case 2
      If lbSearch Then
         If loBranch.SearchRecord(txtBranch.Text, False) Then
            txtBranch.Text = loBranch.Master("sBranchNm")
            p_sBranchCd = loBranch.Master("sBranchCd")
            p_sBranchNm = loBranch.Master("sBranchNm")
            p_sAddressx = loBranch.Master("sAddressx")
         Else
            If p_sBranchCd <> "" Then txtBranch.Text = txtBranch.Tag
         End If
      End If
   End Select
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = loAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail
   
   Set loBranch = New clsBranch
   Set loBranch.AppDriver = loAppDrivr
   loBranch.InitRecord
   loBranch.NewRecord

   txtDateFrom = Format(Trim(Str(Month(loAppDrivr.ServerDate))) + "/1/" + Trim(Str(Year(loAppDrivr.ServerDate))), "MMM DD, YYYY")
   txtDateThru = Format(loAppDrivr.ServerDate, "MMM DD, YYYY")

   pbCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set loBranch = Nothing
End Sub

Private Sub SearchCategory(ByVal lbEqual As Boolean)
   Dim lors As ADODB.Recordset
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String
   Dim lnCtr As Integer

   With txtCategory
      lsSQL = " SELECT" & _
                  " sCategrID" & _
                  ", sCategrNm" & _
               " FROM Category " & _
               " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
               " AND sCategrNm " & _
                     IIf(lbEqual, _
                        " = " & strParm(Trim(.Text)), _
                        " LIKE " & strParm((.Text) & "%"))
                        
      Set lors = New Recordset
      lors.Open lsSQL, loAppDrivr.Connection, , , adCmdText
   
      If lors.EOF Then
         .Text = ""
         psCategID = Empty
      ElseIf lors.RecordCount = 1 Then
         .Text = lors(1)
         psCategID = lors(0)
      Else
         lsSelect = KwikBrowse(loAppDrivr, lors _
                              , "sCategrID»sCategrNm" _
                              , "Code»Category")
      
         If lsSelect <> "" Then
            lasSelect = Split(lsSelect, "»")
            .Text = lasSelect(1)
            psCategID = lasSelect(0)
         Else
            If psCategID <> "" Then .Text = .Tag
         End If
      End If
      .Tag = .Text
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

endProc:
   lors.Close
   Set lors = Nothing
   Exit Sub
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

Private Sub optStat_Click(Index As Integer)
   pnStatus = Index
End Sub

Private Sub optType_Click(Index As Integer)
  pnRcvdTyp = Index
End Sub

Private Sub txtCategory_GotFocus()
   With txtCategory
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = loAppDrivr.getColor("HT1")
   End With

   lbSearch = True
End Sub

Private Sub txtCategory_KeyDown(KeyCode As Integer, Shift As Integer)
   With txtCategory
      If KeyCode = vbKeyF3 Then
         SearchCategory False
         KeyCode = 0
      ElseIf KeyCode = vbKeyReturn Then
         If .Text <> "" Then SearchCategory False
         KeyCode = 0
      End If
   End With
End Sub

Private Sub txtCategory_LostFocus()
   With txtCategory
      .BackColor = loAppDrivr.getColor("EB")
   End With
End Sub

Private Sub txtCategory_Validate(Cancel As Boolean)
   With txtCategory
      If .Text = "" Then
         psCategID = ""
         .Tag = ""
         Exit Sub
      End If
      
      If .Text = .Tag Then Exit Sub
      SearchCategory False
      .Tag = .Text
   End With
End Sub

Private Sub txtDateFrom_GotFocus()
   With txtDateFrom
      .Text = Format(.Text, "MM/DD/YYYY")

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = loAppDrivr.getColor("HT1")
   End With

   lbSearch = False
End Sub

Private Sub txtDateFrom_LostFocus()
   With txtDateFrom
      .BackColor = loAppDrivr.getColor("EB")
   End With
End Sub

Private Sub txtDateFrom_Validate(Cancel As Boolean)
   With txtDateFrom
      If Not IsDate(.Text) Then
         .Text = Format(Trim(Str(Month(Date))) + "/1/" + Trim(Str(Year(Date))), "MMM DD, YYYY")
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
      .BackColor = loAppDrivr.getColor("HT1")
   End With

   lbSearch = False
End Sub

Private Sub txtDateThru_LostFocus()
   With txtDateThru
      .BackColor = loAppDrivr.getColor("EB")
   End With
End Sub

Private Sub txtDateThru_Validate(Cancel As Boolean)
   With txtDateThru
      If Not IsDate(.Text) Then
         .Text = Format(Date, "MMM DD, YYYY")
      Else
         .Text = Format(.Text, "MMM DD, YYYY")
      End If
   End With
End Sub

Private Sub txtBranch_GotFocus()
   With txtBranch
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = loAppDrivr.getColor("HT1")
   End With

   lbSearch = True
End Sub

Private Sub txtBranch_KeyDown(KeyCode As Integer, Shift As Integer)
   With txtBranch
      If KeyCode = vbKeyF3 Then
         If loBranch.SearchRecord(.Text, False) Then
            .Text = loBranch.Master("sBranchNm")
            p_sBranchCd = loBranch.Master("sBranchCd")
            p_sBranchNm = loBranch.Master("sBranchNm")
            p_sAddressx = loBranch.Master("sAddressx")
         Else
            .Text = ""
            If p_sBranchCd <> "" Then .Text = .Tag
         End If
         KeyCode = 0
         
         .SelStart = 0
         .SelLength = Len(.Text)
      ElseIf KeyCode = vbKeyReturn Then
         If .Text <> "" Then
            If loBranch.SearchRecord(.Text, False) Then
               .Text = loBranch.Master("sBranchNm")
               p_sBranchCd = loBranch.Master("sBranchCd")
               p_sBranchNm = loBranch.Master("sBranchNm")
               p_sAddressx = loBranch.Master("sAddressx")
            Else
               .Text = ""
               If p_sBranchCd <> "" Then .Text = .Tag
            End If
         End If
         
         KeyCode = 0
         
         .SelStart = 0
         .SelLength = Len(.Text)
      End If
   End With
End Sub

Private Sub txtBranch_LostFocus()
   With txtBranch
      .BackColor = loAppDrivr.getColor("EB")
   End With
End Sub

Private Sub txtBranch_Validate(Cancel As Boolean)
   With txtBranch
      If .Text = "" Then
         .Tag = ""
         p_sBranchCd = ""
         p_sBranchNm = ""
         p_sAddressx = ""
         Exit Sub
      End If
   
      If .Text = .Tag Then Exit Sub
      If loBranch.SearchRecord(.Text, False) Then
         .Text = loBranch.Master("sBranchNm")
         p_sBranchCd = loBranch.Master("sBranchCd")
         p_sBranchNm = loBranch.Master("sBranchNm")
         p_sAddressx = loBranch.Master("sAddressx")
      Else
         If p_sBranchCd <> "" Then .Text = .Tag
      End If
      
      .Tag = .Text
   End With
End Sub


