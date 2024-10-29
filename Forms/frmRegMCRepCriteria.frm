VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMCRepCriteria 
   BorderStyle     =   0  'None
   Caption         =   "Transaction Summary"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1410
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1935
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2487
      BackColor       =   12632256
      Begin VB.TextBox txtSummary 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   735
         Width           =   4500
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Date"
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Tag             =   "et0;fb0"
         Top             =   420
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Brand"
         Height          =   210
         Index           =   1
         Left            =   1455
         TabIndex        =   10
         Tag             =   "et0;fb0"
         Top             =   405
         Width           =   1035
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Model"
         Height          =   210
         Index           =   2
         Left            =   2610
         TabIndex        =   11
         Tag             =   "et0;fb0"
         Top             =   405
         Width           =   825
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
         TabIndex        =   8
         Tag             =   "et0;fb0"
         Top             =   90
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   1005
         Index           =   2
         Left            =   180
         Top             =   195
         Width           =   4815
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1350
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2381
      Begin VB.OptionButton optPresentation 
         Caption         =   "Detailed"
         Height          =   210
         Index           =   1
         Left            =   375
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   750
         Width           =   1260
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Summarized"
         Height          =   210
         Index           =   0
         Left            =   375
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   450
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   705
         Width           =   1890
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3000
         TabIndex        =   5
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
         Left            =   2115
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   120
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   915
         Index           =   1
         Left            =   2010
         Top             =   210
         Width           =   2985
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
         TabIndex        =   0
         Tag             =   "et0;fb0"
         Top             =   120
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         Height          =   915
         Index           =   0
         Left            =   180
         Top             =   210
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   1
         Left            =   2175
         TabIndex        =   6
         Top             =   735
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   0
         Left            =   2175
         TabIndex        =   4
         Top             =   450
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5535
      TabIndex        =   15
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
      Picture         =   "frmRegMCRepCriteria.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   5535
      TabIndex        =   13
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
      Picture         =   "frmRegMCRepCriteria.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5535
      TabIndex        =   14
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
      Picture         =   "frmRegMCRepCriteria.frx":0EF4
   End
End
Attribute VB_Name = "frmMCRepCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loAppDrivr As clsAppDriver
Private oSkin As clsFormSkin
Private pbCancel As Boolean
Private pnSummary As Integer
Private pnPresntn As Integer
Private psSummary As String

Dim lbSearch As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set loAppDrivr = oAppDriver
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
   '  0 = Summarized
   '  1 = Detailed
   Presentation = pnPresntn
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
         SearchSummary False
      End If
   End Select
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = loAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail

   txtDateFrom = Format(Trim(Str(Month(Date))) + "/1/" + Trim(Str(Year(Date))), "MMM DD, YYYY")
   txtDateThru = Format(Date, "MMM DD, YYYY")

   optPresentation(0).Value = True
   optPresentation(1).Value = False

   optSummary(0).Value = True
   optSummary(1).Value = False
   optSummary(2).Value = False

   psSummary = ""
   pnSummary = 0
   pnPresntn = 0
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

   With txtSummary
      Select Case pnSummary
      Case 0   ' Date
         GoTo endProc
      Case 1   ' Brand
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
      Case 2   ' Model
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
      End Select

      Set lors = New Recordset
      
      lors.Open lsSQL, loAppDrivr.Connection, , , adCmdText

      If lors.EOF Then
         .Text = ""
         psSummary = Empty
      ElseIf lors.RecordCount = 1 Then
         .Text = lors(1)
         psSummary = lors(0)
      Else
         Select Case pnSummary
         Case 1
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sBrandIDx»sBrandNme" _
                                 , "Code»Brand Name")
         Case 2
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sModelIDx»sModelNme" _
                                 , "Code»Model Name")
         End Select

         If lsSelect <> "" Then
            lasSelect = Split(lsSelect, "»")
            .Text = lasSelect(1)
            psSummary = lasSelect(0)
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

Private Sub optSummary_Click(Index As Integer)
   pnSummary = Index
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
   End With

   lbSearch = False
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
   End With
   lbSearch = True
End Sub

Private Sub txtSummary_KeyDown(KeyCode As Integer, Shift As Integer)
   With txtSummary
      If KeyCode = vbKeyF3 Then
         SearchSummary False
         KeyCode = 0
      End If
   End With
End Sub

Private Sub txtSummary_Validate(Cancel As Boolean)
   With txtSummary
      If .Text = .Tag Then Exit Sub

      SearchSummary False
   End With
End Sub
