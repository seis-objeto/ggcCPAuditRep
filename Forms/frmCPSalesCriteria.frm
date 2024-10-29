VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPSalesCriteria 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1320
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   2328
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2970
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   330
         Width           =   1890
      End
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2970
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   690
         Width           =   1890
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Detailed"
         Height          =   210
         Index           =   1
         Left            =   330
         TabIndex        =   10
         Tag             =   "et0;fb0"
         Top             =   765
         Width           =   1260
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Summarized"
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Tag             =   "et0;fb0"
         Top             =   450
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   0
         Left            =   2100
         TabIndex        =   24
         Top             =   390
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   23
         Top             =   705
         Width           =   825
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
         Left            =   180
         TabIndex        =   12
         Tag             =   "et0;fb0"
         Top             =   75
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         Height          =   1005
         Index           =   0
         Left            =   90
         Top             =   150
         Width           =   1830
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
         Left            =   2115
         TabIndex        =   11
         Tag             =   "et0;fb0"
         Top             =   90
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   1020
         Index           =   1
         Left            =   1995
         Top             =   150
         Width           =   3045
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3030
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1920
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   5345
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1530
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   2460
         Width           =   3315
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1530
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2085
         Width           =   3315
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1530
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1710
         Width           =   3315
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1530
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1335
         Width           =   3315
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1530
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   3315
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1530
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   600
         Width           =   3315
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1530
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   3315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Financer"
         Height          =   195
         Index           =   4
         Left            =   285
         TabIndex        =   26
         Top             =   2535
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   12
         Left            =   285
         TabIndex        =   20
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   11
         Left            =   285
         TabIndex        =   19
         Top             =   1725
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Category"
         Height          =   195
         Index           =   10
         Left            =   285
         TabIndex        =   18
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   15
         Top             =   975
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   6
         Left            =   270
         TabIndex        =   14
         Top             =   645
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   13
         Top             =   255
         Width           =   915
      End
      Begin VB.Shape Shape1 
         Height          =   2775
         Index           =   3
         Left            =   105
         Top             =   105
         Width           =   4920
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5490
      TabIndex        =   21
      Top             =   1845
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
      Picture         =   "frmCPSalesCriteria.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   5490
      TabIndex        =   8
      Top             =   585
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
      Picture         =   "frmCPSalesCriteria.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5490
      TabIndex        =   22
      Top             =   1215
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
      Picture         =   "frmCPSalesCriteria.frx":0EF4
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
      Height          =   195
      Index           =   9
      Left            =   2175
      TabIndex        =   17
      Top             =   4080
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
      Height          =   195
      Index           =   8
      Left            =   2220
      TabIndex        =   16
      Top             =   3690
      Width           =   510
   End
End
Attribute VB_Name = "frmCPSalesCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'she 2016-03-17 04:57 pm

Private loAppDrivr As clsAppDriver
Private oSkin As clsFormSkin
Private pbCancel As Boolean

Private pnIndex As Integer 'txtsearch index
Private pnPresentation As Integer 'presentation index

Private psAreaDesc As String
Private psBranch As String
Private psCategory1 As String
Private psCategory2 As String
Private psBrand As String
Private psModel As String
Private psFinancer As String

Dim lbSearch As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set loAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = psBranch
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

Property Get Presentation() As Integer
   Presentation = pnPresentation
End Property

Property Get Area() As String
   Area = psAreaDesc
End Property

Property Get Financer() As String
   Financer = psFinancer
End Property

Property Get Category() As String
   Category = psCategory1
End Property

Property Get SubCategory() As String
   SubCategory = psCategory2
End Property

Property Get Brand() As String
   Brand = psBrand
End Property

Property Get Model() As String
   Model = psModel
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
      End If
   End Select
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = loAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail


   txtDateFrom = Format(Trim(Str(Month(loAppDrivr.ServerDate))) + "/1/" + Trim(Str(Year(loAppDrivr.ServerDate))), "MMM DD, YYYY")
   txtDateThru = Format(loAppDrivr.ServerDate, "MMM DD, YYYY")

   InitForm
   
   If loAppDrivr.IsMainOffice = True Or loAppDrivr.IsWarehouse = True Then
      txtSearch(0).Enabled = True
      txtSearch(1).Enabled = True
   Else
      txtSearch(0).Enabled = False
      txtSearch(1).Enabled = False
      txtSearch(1).Text = loAppDrivr.BranchName
      psBranch = loAppDrivr.BranchCode
   End If
   

   pbCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub SearchSummary(ByVal lbEqual As Boolean)
   Dim lors As ADODB.Recordset
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String
   Dim lnCtr As Integer

   With txtSearch
      Select Case pnIndex
      Case 0   ' Area
         lsSQL = "SELECT" & _
                     "  sAreaCode" & _
                     ", sAreaDesc" & _
                  " FROM Branch_Area" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND sAreaDesc" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(txtSearch(pnIndex).Text)), _
                           " LIKE " & strParm((txtSearch(pnIndex).Text) & "%")) & _
                  " ORDER BY sAreaDesc"
      Case 1   ' Branch
         lsSQL = "SELECT" & _
                     "  sBranchCd" & _
                     ", sBranchNm" & _
                  " FROM Branch" & _
                  " WHERE cAutomate = " & strParm(xeRecStateActive) & _
                     " AND sBranchNm" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(txtSearch(pnIndex).Text)), _
                           " LIKE " & strParm((txtSearch(pnIndex).Text) & "%")) & _
                  " ORDER BY sBranchNm"
      Case 2   ' Category
         lsSQL = "SELECT" & _
                     "  sCategrID" & _
                     ", sCategrNm" & _
                  " FROM Category" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND cLevelxxx = '1' " & _
                     " AND sCategrNm" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(txtSearch(pnIndex).Text)), _
                           " LIKE " & strParm((txtSearch(pnIndex).Text) & "%")) & _
                  " ORDER BY sCategrNm"
      Case 3   ' Sub Category
         lsSQL = "SELECT" & _
                     "  sCategrID" & _
                     ", sCategrNm" & _
                  " FROM Category" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND cLevelxxx = '2' " & _
                     " AND sCategrNm" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(txtSearch(pnIndex).Text)), _
                           " LIKE " & strParm((txtSearch(pnIndex).Text) & "%")) & _
                  " ORDER BY sCategrNm"
      Case 4 ' Brand
         lsSQL = "SELECT" & _
                     "  sBrandIDx" & _
                     ", sBrandNme" & _
                  " FROM CP_Brand" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND sBrandNme" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(txtSearch(pnIndex).Text)), _
                           " LIKE " & strParm((txtSearch(pnIndex).Text) & "%")) & _
                  " ORDER BY sBrandNme"
      Case 5 ' model
         lsSQL = "SELECT" & _
                     "  a.sModelIDx" & _
                     ", a.sModelNme" & _
                     ", b.sBrandNme" & _
                  " FROM CP_Model a" & _
                        " LEFT JOIN CP_Brand b" & _
                           " ON a.sBrandIDx = b.sBrandIDx" & _
                  " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND a.sModelNme" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(txtSearch(pnIndex).Text)), _
                           " LIKE " & strParm((txtSearch(pnIndex).Text) & "%")) & _
                  " ORDER BY a.sModelNme"
      Case 6 'Financer
         lsSQL = "SELECT" & _
                     "  a.sClientID" & _
                     ", b.sCompnyNm" & _
                  " FROM CP_SO_Finance a" & _
                        " LEFT JOIN Client_Master b" & _
                           " ON a.sClientID = b.sClientID" & _
                  " WHERE b.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND b.sCompnyNm" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(txtSearch(pnIndex).Text)), _
                           " LIKE " & strParm((txtSearch(pnIndex).Text) & "%")) & _
                  " GROUP BY a.sClientID" & _
                  " ORDER BY b.sCompnyNm"
      End Select
      Debug.Print lsSQL
      Set lors = New Recordset
      lors.Open lsSQL, loAppDrivr.Connection, , , adCmdText

      If Not lors.EOF Then
         Select Case pnIndex
         Case 0 'Area
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sAreaCode»sAreaDesc" _
                                 , "Code»Area Desc")
            If lsSelect <> "" Then
               lasSelect = Split(lsSelect, "»")
               txtSearch(pnIndex).Text = lasSelect(1)
               psAreaDesc = lasSelect(0)
            Else
               txtSearch(pnIndex).Text = ""
               psAreaDesc = ""
            End If
         Case 1 'Branch
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sBranchCd»sBranchNm" _
                                 , "Code»Branch")
            If lsSelect <> "" Then
               lasSelect = Split(lsSelect, "»")
               txtSearch(pnIndex).Text = lasSelect(1)
               psBranch = lasSelect(0)
            Else
               txtSearch(pnIndex).Text = ""
               psBranch = ""
            End If
         Case 2 'Category 1
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sCategrID»sCategrNm" _
                                 , "ID»Categ Name")
            If lsSelect <> "" Then
               lasSelect = Split(lsSelect, "»")
               txtSearch(pnIndex).Text = lasSelect(1)
               psCategory1 = lasSelect(0)
            Else
               txtSearch(pnIndex).Text = ""
               psCategory1 = ""
            End If
         Case 3 'Category 2
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sCategrID»sCategrNm" _
                                 , "ID»Categ Name")
            If lsSelect <> "" Then
               lasSelect = Split(lsSelect, "»")
               txtSearch(pnIndex).Text = lasSelect(1)
               psCategory2 = lasSelect(0)
            Else
               txtSearch(pnIndex).Text = ""
               psCategory2 = ""
            End If
         Case 4 'Brand
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sBrandIDx»sBrandNme" _
                                 , "ID»Brand Name")
            If lsSelect <> "" Then
               lasSelect = Split(lsSelect, "»")
               txtSearch(pnIndex).Text = lasSelect(1)
               psBrand = lasSelect(0)
            Else
               txtSearch(pnIndex).Text = ""
               psBrand = ""
            End If
         Case 5 'Model
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sModelIDx»sModelNme»sBrandNme" _
                                 , "ID»Model Name»Brand")
            If lsSelect <> "" Then
               lasSelect = Split(lsSelect, "»")
               txtSearch(pnIndex).Text = lasSelect(1)
               psModel = lasSelect(0)
            Else
               txtSearch(pnIndex).Text = ""
               psModel = ""
            End If
         Case 6 'Financer
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sClientID»sCompnyNm" _
                                 , "ID»ClientName")
            If lsSelect <> "" Then
               lasSelect = Split(lsSelect, "»")
               txtSearch(pnIndex).Text = lasSelect(1)
               psFinancer = lasSelect(0)
            Else
               txtSearch(pnIndex).Text = ""
               psFinancer = ""
            End If
         End Select

      End If
   End With

endProc:
   Set lors = Nothing
   txtSearch(pnIndex).Tag = txtSearch(pnIndex).Text
   txtSearch(pnIndex).SelStart = 0
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
   pnPresentation = Index
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

Private Sub InitForm()

   txtSearch(0).Text = ""
   txtSearch(1).Text = ""
   txtSearch(2).Text = ""
   txtSearch(3).Text = ""
   txtSearch(4).Text = ""
   txtSearch(5).Text = ""
   txtSearch(6).Text = ""
   
End Sub

Private Sub txtSearch_Click(Index As Integer)
   pnIndex = Index
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch
      txtSearch(Index).BackColor = &HC0FFFF
      pnIndex = Index
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      SearchSummary False
   End If
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch
      txtSearch(Index).BackColor = &H80000005
      pnIndex = Index
   End With
End Sub
