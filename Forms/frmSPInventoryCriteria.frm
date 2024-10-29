VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSPInventoryCriteria 
   BorderStyle     =   0  'None
   Caption         =   "Transaction Summary"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1275
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1500
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2249
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.OptionButton optSummary 
         Caption         =   "Model"
         Height          =   210
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Tag             =   "et0;fb0"
         Top             =   450
         Width           =   825
      End
      Begin VB.TextBox txtSummary 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   705
         Width           =   4560
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Category"
         Height          =   210
         Index           =   0
         Left            =   525
         TabIndex        =   4
         Tag             =   "et0;fb0"
         Top             =   450
         Width           =   975
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Brand"
         Height          =   210
         Index           =   2
         Left            =   3705
         TabIndex        =   6
         Tag             =   "et0;fb0"
         Top             =   450
         Value           =   -1  'True
         Width           =   1035
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
         Left            =   270
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   120
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   885
         Index           =   1
         Left            =   180
         Top             =   195
         Width           =   4800
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   915
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1614
      Begin VB.OptionButton optPresentation 
         Caption         =   "All"
         Height          =   210
         Index           =   2
         Left            =   3630
         TabIndex        =   11
         Tag             =   "et0;fb0"
         Top             =   390
         Width           =   1260
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Replacement"
         Height          =   210
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   390
         Width           =   1260
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Genuine"
         Height          =   210
         Index           =   1
         Left            =   525
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   390
         Value           =   -1  'True
         Width           =   1260
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
         Height          =   540
         Index           =   0
         Left            =   180
         Top             =   195
         Width           =   4800
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5535
      TabIndex        =   10
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
      Picture         =   "frmSPInventoryCriteria.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   5535
      TabIndex        =   8
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
      Picture         =   "frmSPInventoryCriteria.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5535
      TabIndex        =   9
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
      Picture         =   "frmSPInventoryCriteria.frx":0EF4
   End
End
Attribute VB_Name = "frmSPInventoryCriteria"
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

Dim pbSearch As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set loAppDrivr = oAppDriver
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancel
End Property

Property Get Specify() As String
   Specify = psSummary
End Property

Property Get Summary() As Integer
   '  0 = Category
   '  1 = Model
   '  2 = Brand

   Summary = pnSummary
End Property

Property Get Presentation() As Integer
   '  0 = Replacement
   '  1 = Genuine
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
      If pbSearch Then
         SearchSummary False
      End If
   End Select
End Sub

Private Sub Form_Activate()
   pbCancel = True
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = loAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail

   optPresentation(0).Value = False
   optPresentation(1).Value = True

   optSummary(0).Value = False
   optSummary(1).Value = False
   optSummary(2).Value = True

   psSummary = ""
   pnSummary = 2
   pnPresntn = 1
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
      Case 0   ' Category
         lsSQL = "SELECT" & _
                     "  sCategrID" & _
                     ", sCategrNm" & _
                  " FROM Category" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND sCategrNm" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(.Text)), _
                           " LIKE " & strParm((.Text) & "%")) & _
                  " ORDER BY sCategrNm"
      Case 1   ' Model
         lsSQL = "SELECT" & _
                     "  sModelIDx" & _
                     ", sModelNme" & _
                  " FROM SP_Model" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND sModelNme" & _
                        IIf(lbEqual, _
                           " = " & strParm(Trim(.Text)), _
                           " LIKE " & strParm((.Text) & "%")) & _
                  " ORDER BY sModelNme"
      Case 2   ' Brand
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
         Case 0
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sCategrID»sCategrNm" _
                                 , "Code»Category Name")
         Case 1
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sModelIDx»sModelNme" _
                                 , "Code»Model Name")
         Case 2
            lsSelect = KwikBrowse(loAppDrivr, lors _
                                 , "sBrandIDx»sBrandNme" _
                                 , "Code»Brand Name")
         
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

Private Sub txtSummary_GotFocus()
   With txtSummary
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   pbSearch = True
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
