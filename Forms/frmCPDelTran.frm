VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPDelTran 
   BorderStyle     =   0  'None
   Caption         =   "Transaction Summary"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2265
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1905
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3995
      Begin xrControl.xrFrame xrFrame5 
         Height          =   315
         Left            =   255
         Tag             =   "wt0;fb0"
         Top             =   690
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   556
         BackColor       =   12632256
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.OptionButton optSerialized 
            Caption         =   "w/o Serial"
            Height          =   195
            Index           =   2
            Left            =   2340
            TabIndex        =   16
            Tag             =   "et0;fb0"
            Top             =   60
            Width           =   1110
         End
         Begin VB.OptionButton optSerialized 
            Caption         =   "w/Serial"
            Height          =   195
            Index           =   1
            Left            =   1305
            TabIndex        =   15
            Tag             =   "et0;fb0"
            Top             =   60
            Width           =   960
         End
         Begin VB.OptionButton optSerialized 
            Caption         =   "All"
            Height          =   195
            Index           =   0
            Left            =   705
            TabIndex        =   14
            Tag             =   "et0;fb0"
            Top             =   60
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.Label Label2 
            Caption         =   "Type"
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   13
            Tag             =   "et0;fb0"
            Top             =   45
            Width           =   675
         End
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Refer No."
         Height          =   210
         Index           =   3
         Left            =   3945
         TabIndex        =   22
         Tag             =   "et0;fb0"
         Top             =   1485
         Width           =   990
      End
      Begin VB.TextBox txtSummary 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   23
         Top             =   1695
         Width           =   4740
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Brand"
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Tag             =   "et0;fb0"
         Top             =   1470
         Width           =   720
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Model"
         Height          =   210
         Index           =   1
         Left            =   1380
         TabIndex        =   20
         Tag             =   "et0;fb0"
         Top             =   1470
         Width           =   735
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Supplier"
         Height          =   210
         Index           =   2
         Left            =   2595
         TabIndex        =   21
         Tag             =   "et0;fb0"
         Top             =   1470
         Width           =   870
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         TabIndex        =   18
         Top             =   1035
         Width           =   3930
      End
      Begin VB.TextBox txtBranch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         TabIndex        =   12
         Top             =   300
         Width           =   4170
      End
      Begin VB.Shape Shape2 
         Height          =   780
         Left            =   195
         Top             =   645
         Width           =   4740
      End
      Begin VB.Label Label2 
         Caption         =   "Category"
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   17
         Tag             =   "et0;fb0"
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Branch"
         Height          =   165
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Tag             =   "et0;fb0"
         Top             =   330
         Width           =   1020
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
         TabIndex        =   10
         Tag             =   "et0;fb0"
         Top             =   75
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   1950
         Index           =   2
         Left            =   75
         Top             =   180
         Width           =   4995
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1335
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2355
      Begin xrControl.xrFrame xrFrame4 
         Height          =   450
         Left            =   90
         Tag             =   "wt0;fb0"
         Top             =   735
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   794
         BackColor       =   12632256
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.OptionButton optType 
            Caption         =   "Load"
            Height          =   195
            Index           =   1
            Left            =   1290
            TabIndex        =   4
            Tag             =   "et0;fb0"
            Top             =   150
            Width           =   660
         End
         Begin VB.OptionButton optType 
            Caption         =   "Cellphone"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   3
            Tag             =   "et0;fb0"
            Top             =   150
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.Line Line1 
            X1              =   15
            X2              =   2025
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Delivered"
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   525
         Width           =   1950
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Transfers"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   315
         Value           =   -1  'True
         Width           =   1950
      End
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3105
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   705
         Width           =   1890
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3105
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   375
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
         Left            =   2220
         TabIndex        =   5
         Tag             =   "et0;fb0"
         Top             =   120
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   990
         Index           =   1
         Left            =   2190
         Top             =   210
         Width           =   2880
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
         TabIndex        =   0
         Tag             =   "et0;fb0"
         Top             =   120
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         Height          =   990
         Index           =   0
         Left            =   75
         Top             =   210
         Width           =   2070
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   8
         Top             =   735
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Top             =   450
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5535
      TabIndex        =   26
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
      Picture         =   "frmCPDelTran.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   5535
      TabIndex        =   24
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
      Picture         =   "frmCPDelTran.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5535
      TabIndex        =   25
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
      Picture         =   "frmCPDelTran.frx":0EF4
   End
End
Attribute VB_Name = "frmCPDelTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum xeSeachCriteria
   xeBranch = 1
   xeCategory = 2
   xeOthers = 3
End Enum

Public Enum xeSerialType
   xeAll = 0
   xeWithSerial = 1
   xeWithOutSerial = 2
End Enum

Private loAppDrivr As clsAppDriver
Private loBranch As clsBranch
Private oSkin As clsFormSkin
Private pbCancel As Boolean
Private pnSummary As Integer
Private pnPresntn As Integer
Private psSummary As String
Private pnSerialx As xeSerialType
Private pnRcvdTyp As Integer
Private psCategID As String

Private p_sBranchCd As String
Private p_sBranchNm As String
Private p_sAddressx As String

Dim lbSearch As Boolean
Dim lnSummary As xeSeachCriteria

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
   '  0 = Date
   '  1 = Brand
   '  2 = Model
   Summary = pnSummary
End Property

Property Get Presentation() As Integer
   '  0 = Tranfer
   '  1 = Delivered
   Presentation = pnPresntn
End Property

Property Get Serialized() As xeSerialType
   Serialized = pnSerialx
End Property

Property Let Serialized(lnSerialized As xeSerialType)
   lnSerialized = pnSerialx
End Property

Property Get Category() As String
   Category = psCategID
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
         Select Case lnSummary
         Case xeBranch
            If loBranch.SearchRecord(txtBranch.Text, False) Then
               txtBranch.Text = loBranch.Master("sBranchNm")
               p_sBranchCd = loBranch.Master("sBranchCd")
               p_sBranchNm = loBranch.Master("sBranchNm")
               p_sAddressx = loBranch.Master("sAddressx")
            Else
               If p_sBranchCd <> "" Then txtBranch.Text = txtBranch.Tag
            End If
         Case xeCategory
            SearchCategory False
         Case xeOthers
            SearchSummary False
         End Select
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

   optPresentation(0).Value = True
   optPresentation(1).Value = False

   optSummary(0).Value = True
   optSummary(1).Value = False
   optSummary(2).Value = False
   optType(0).Value = True
   optType(1).Value = False
   optSerialized(0).Value = True

   psSummary = ""
   pnSummary = 0
   pnPresntn = 0
   pnSerialx = 0
   psCategID = ""
   pnRcvdTyp = 0
   
   pbCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set loBranch = Nothing
End Sub

Private Sub SearchCategory(ByVal lbEqual As Boolean)
   Dim loRS As ADODB.Recordset
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
                        
   ' she 2014-07-19 select all category only
'      "SELECT" & _
'                  "  a.sCategrID" & _
'                  ", a.sCategrNm" & _
'               " FROM Category a" & _
'                  ", CP_Inventory b" & _
'               " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
'                  " AND a.sCategrNm" & _
'                     IIf(lbEqual, _
'                        " = " & strParm(Trim(.Text)), _
'                        " LIKE " & strParm((.Text) & "%")) & _
'                  " AND a.sCategrID = b.sCategID1" & _
'                  IIf(pnSerialx = xeAll, "", " AND b.cHsSerial = " & CDbl(pnSerialx)) & _
'               " ORDER BY a.sCategrNm"
'
      Set loRS = New Recordset
      loRS.Open lsSQL, loAppDrivr.Connection, , , adCmdText
   
      If loRS.EOF Then
         .Text = ""
         psCategID = Empty
      ElseIf loRS.RecordCount = 1 Then
         .Text = loRS(1)
         psCategID = loRS(0)
      Else
         lsSelect = KwikBrowse(loAppDrivr, loRS _
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
   loRS.Close
   Set loRS = Nothing
   Exit Sub
End Sub

Private Sub SearchSummary(ByVal lbEqual As Boolean)
   Dim loRS As ADODB.Recordset
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String
   Dim lnCtr As Integer

   With txtSummary
      Select Case pnSummary
      Case 0   ' Brand
         lsSQL = "SELECT" & _
                     "  sBrandIDx" & _
                     ", sBrandNme" & _
                  " FROM Brand" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND sBrandNme" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(.Text)), _
                           " LIKE " & strParm((.Text) & "%")) & _
                  " ORDER BY sBrandNme"
      Case 1   ' Model
         lsSQL = "SELECT" & _
                     "  sModelIDx" & _
                     ", sModelNme" & _
                  " FROM MC_Model" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND sModelNme" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(.Text)), _
                           " LIKE " & strParm((.Text) & "%")) & _
                  " ORDER BY sModelNme"
      Case 2   ' Supplier
         lsSQL = "SELECT" & _
                     "  b.sClientID" & _
                     ", b.sCompnyNm" & _
                  " FROM CP_Supplier a" & _
                     ", Client_Master b" & _
                  " WHERE b.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND a.sBranchCd = " & strParm(IIf(p_sBranchCd = "", loAppDrivr.BranchCode, p_sBranchCd)) & _
                     " AND b.sCompnyNm" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(.Text)), _
                           " LIKE " & strParm((.Text) & "%")) & _
                  " GROUP BY b.sClientID" & _
                  " ORDER BY b.sCompnyNm"
      Case 3   ' Reference No
         psSummary = .Text
         GoTo endProc
      End Select

      Set loRS = New Recordset
      loRS.Open lsSQL, loAppDrivr.Connection, , , adCmdText
   
      If loRS.EOF Then
         .Text = ""
         psSummary = Empty
      ElseIf loRS.RecordCount = 1 Then
         .Text = loRS(1)
         psSummary = loRS(0)
      Else
         Select Case pnSummary
         Case 0
            lsSelect = KwikBrowse(loAppDrivr, loRS _
                                 , "sBrandIDx»sBrandNme" _
                                 , "Code»Brand Name")
         Case 1
            lsSelect = KwikBrowse(loAppDrivr, loRS _
                                 , "sModelIDx»sModelNme" _
                                 , "Code»Model Name")
         Case 2
            lsSelect = KwikBrowse(loAppDrivr, loRS _
                                 , "sClientID»sCompnyNm" _
                                 , "Code»Company Name")
         End Select

         If lsSelect <> "" Then
            lasSelect = Split(lsSelect, "»")
            .Text = lasSelect(1)
            psSummary = lasSelect(0)
         Else
            If psSummary <> "" Then .Text = .Tag
         End If
      End If

endProc:
   Set loRS = Nothing
   .Tag = .Text
   .SelStart = 0
   .SelLength = Len(.Text)
   Exit Sub
   End With
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

Private Sub optSerial_Click(Index As Integer)
   pnSerialx = Index
End Sub

Private Sub optSerialized_Click(Index As Integer)
   pnSerialx = Index
End Sub

Private Sub optSummary_Click(Index As Integer)
   pnSummary = Index
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
   lnSummary = xeCategory
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

Private Sub txtSummary_GotFocus()
   With txtSummary
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = loAppDrivr.getColor("HT1")
   End With
   
   lbSearch = True
   lnSummary = xeOthers
End Sub

Private Sub txtSummary_KeyDown(KeyCode As Integer, Shift As Integer)
   With txtSummary
      If KeyCode = vbKeyF3 Then
         SearchSummary False
         KeyCode = 0
      ElseIf KeyCode = vbKeyReturn Then
         If .Text <> "" Then SearchSummary False
         KeyCode = 0
      End If
   End With
End Sub

Private Sub txtSummary_LostFocus()
   With txtSummary
      .BackColor = loAppDrivr.getColor("EB")
   End With
End Sub

Private Sub txtSummary_Validate(Cancel As Boolean)
   With txtSummary
      If .Text = "" Then
         psSummary = ""
         .Tag = ""
         Exit Sub
      End If
      
      If .Text = .Tag Then Exit Sub
      SearchSummary False
      .Tag = .Text
   End With
End Sub

Private Sub txtBranch_GotFocus()
   With txtBranch
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = loAppDrivr.getColor("HT1")
   End With

   lbSearch = True
   lnSummary = xeBranch
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
