VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPRepCriteria 
   BorderStyle     =   0  'None
   Caption         =   "CP Inventory Summary"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "et0;eb0;et0;bc2"
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
      Picture         =   "frmCPRepCriteria.frx":0000
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
      Picture         =   "frmCPRepCriteria.frx":077A
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
      Picture         =   "frmCPRepCriteria.frx":0EF4
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1350
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1935
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2381
      BackColor       =   12632256
      Begin VB.TextBox txtSpecify 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   885
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   690
         Width           =   3990
      End
      Begin VB.TextBox txtSpecify 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   885
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Width           =   3990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   9
         Top             =   405
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Thru"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   11
         Top             =   690
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Filter"
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
         Height          =   930
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
      Top             =   540
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2381
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3030
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   360
         Width           =   1890
      End
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3030
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   690
         Width           =   1890
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Description"
         Height          =   210
         Index           =   1
         Left            =   375
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   735
         Width           =   1260
      End
      Begin VB.OptionButton optPresentation 
         Caption         =   "Barcode"
         Height          =   210
         Index           =   0
         Left            =   375
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   435
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
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   105
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
         TabIndex        =   0
         Tag             =   "et0;fb0"
         Top             =   105
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   0
         Left            =   2205
         TabIndex        =   4
         Top             =   435
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Thru"
         Height          =   195
         Index           =   1
         Left            =   2205
         TabIndex        =   6
         Top             =   720
         Width           =   825
      End
      Begin VB.Shape Shape1 
         Height          =   915
         Index           =   1
         Left            =   2025
         Top             =   195
         Width           =   3000
      End
      Begin VB.Shape Shape1 
         Height          =   915
         Index           =   0
         Left            =   180
         Top             =   195
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmCPRepCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loAppDrivr As clsAppDriver
Private loSkin As clsFormSkin

Private pbCancel As Boolean
Private pnPresntn As Integer
Private psSpecifyFr As String
Private psSpecifyTo As String

Dim lbSearch As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set loAppDrivr = oAppDriver
   psSpecifyFr = loAppDrivr.BranchCode
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

Property Get SpecifyFrom() As String
   SpecifyFrom = psSpecifyFr
End Property

Property Get SpecifyThru() As String
   SpecifyThru = psSpecifyTo
End Property

Property Get Presentation() As Integer
   '  0 = Bar Code
   '  1 = Description
   Presentation = pnPresntn
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer

   Select Case Index
   Case 0, 1
      If StrComp(psSpecifyFr, psSpecifyTo, vbTextCompare) > 1 Then
         txtSpecify(1).Text = txtSpecify(0).Text
         psSpecifyTo = psSpecifyFr
      End If
      
      pbCancel = Index = 1
      Me.Hide
   Case 2
      If lbSearch Then
         SearchSpecifyFrom False
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

   optPresentation(0).Value = True
   optPresentation(1).Value = False
   
   txtSpecify(0).Text = ""
   txtSpecify(1).Text = ""
   psSpecifyFr = ""
   psSpecifyTo = ""
   pbCancel = True
   pnPresntn = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set loSkin = Nothing
End Sub

Private Sub SearchSpecifyFrom(ByVal lbEqual As Boolean, Optional Value As Variant)
   Dim lsBrowse As String
   Dim lasSelected() As String
   Dim lsSQL As String
   Dim loRS As Recordset

   lsSQL = "SELECT" & _
               "  sBarrCode" & _
               ", sDescript" & _
            " FROM CP_Inventory"

   If Not IsMissing(Value) Then
      lsSQL = lsSQL & " WHERE " & _
               IIf(Presentation = 0, " sBarrCode", " sDescript") & _
               IIf(lbEqual, " = " & strParm(Value), " LIKE " & strParm(Value & "%"))
   End If
            
   Set loRS = New Recordset
   loRS.Open lsSQL, loAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If loRS.EOF Then
      txtSpecify(0).Text = ""
      psSpecifyFr = ""
   ElseIf loRS.RecordCount = 1 Then
      If Presentation = 0 Then
         txtSpecify(0).Text = loRS("sBarrCode")
         psSpecifyFr = loRS("sBarrCode")
      Else
         txtSpecify(0).Text = loRS("sDescript")
         psSpecifyFr = loRS("sDescript")
      End If
   Else
      lsBrowse = KwikBrowse(loAppDrivr, loRS _
                              , "sBarrCode»sDescript" _
                              , "Bar Code»Description" _
                              , "@»@", False)

      If lsBrowse <> "" Then
         lasSelected = Split(lsBrowse, "»")
         If Presentation = 0 Then
            txtSpecify(0).Text = lasSelected(0)
            psSpecifyFr = lasSelected(0)
         Else
            txtSpecify(0).Text = lasSelected(1)
            psSpecifyFr = lasSelected(1)
         End If
      Else
         If psSpecifyFr = "" Then
            txtSpecify(0).Text = ""
         Else
            txtSpecify(0).Text = txtSpecify(0).Tag
         End If
      End If
   End If
   
   txtSpecify(0).Tag = txtSpecify(0).Text
   txtSpecify(0).SelStart = 0
   txtSpecify(0).SelLength = Len(txtSpecify(0).Text)
   loRS.Close

   Set loRS = Nothing
End Sub


Private Sub SearchSpecifyThru(ByVal lbEqual As Boolean, Optional Value As Variant)
   Dim lsBrowse As String
   Dim lasSelected() As String
   Dim lsSQL As String
   Dim loRS As Recordset

   lsSQL = "SELECT" & _
               "  sBarrCode" & _
               ", sDescript" & _
            " FROM CP_Inventory"
      
   If Not IsMissing(Value) Then
      lsSQL = lsSQL & " WHERE " & _
               IIf(Presentation = 0, " sBarrCode", " sDescript") & _
               IIf(lbEqual, " = " & strParm(Value), " LIKE " & strParm(Value & "%"))
   End If
            
   Set loRS = New Recordset
   loRS.Open lsSQL, loAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If loRS.EOF Then
      txtSpecify(1).Text = ""
      psSpecifyTo = ""
   ElseIf loRS.RecordCount = 1 Then
      If Presentation = 0 Then
         txtSpecify(1).Text = loRS("sBarrCode")
         psSpecifyTo = loRS("sBarrCode")
      Else
         txtSpecify(1).Text = loRS("sDescript")
         psSpecifyTo = loRS("sDescript")
      End If
   Else
      lsBrowse = KwikBrowse(loAppDrivr, loRS _
                              , "sBarrCode»sDescript" _
                              , "Bar Code»Description" _
                              , "@»@", False)

      If lsBrowse <> "" Then
         lasSelected = Split(lsBrowse, "»")
         If Presentation = 0 Then
            txtSpecify(1).Text = lasSelected(0)
            psSpecifyTo = lasSelected(0)
         Else
            txtSpecify(1).Text = lasSelected(1)
            psSpecifyTo = lasSelected(1)
         End If
      Else
         If psSpecifyTo = "" Then
            txtSpecify(1).Text = ""
         Else
            txtSpecify(1).Text = txtSpecify(1).Tag
         End If
      End If
   End If
   
   If StrComp(psSpecifyFr, psSpecifyTo, vbTextCompare) > 0 Then
      txtSpecify(1).Text = txtSpecify(0).Text
      psSpecifyTo = psSpecifyFr
   End If
   
   txtSpecify(1).Tag = txtSpecify(1).Text
   txtSpecify(1).SelStart = 0
   txtSpecify(1).SelLength = Len(txtSpecify(1).Text)
   loRS.Close

   Set loRS = Nothing
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

Private Sub txtSpecify_GotFocus(Index As Integer)
   With txtSpecify(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   lbSearch = True
End Sub

Private Sub txtSpecify_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With txtSpecify(Index)
      If KeyCode = vbKeyF3 Then
         If .Text = .Tag Then Exit Sub
         If Index = 0 Then
            Call SearchSpecifyFrom(False, .Text)
         Else
            Call SearchSpecifyThru(False, .Text)
         End If
         KeyCode = 0
      End If
   End With
End Sub

Private Sub txtSpecify_Validate(Index As Integer, Cancel As Boolean)
   With txtSpecify(Index)
      If .Text = .Tag Then Exit Sub

      If Index = 0 Then
         Call SearchSpecifyFrom(False, .Text)
      Else
         Call SearchSpecifyThru(False, .Text)
      End If
      .Tag = .Text
   End With
End Sub


