VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmReports 
   BorderStyle     =   0  'None
   Caption         =   "Report"
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5745
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   10134
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5775
         Left            =   -15
         TabIndex        =   0
         Top             =   -15
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   10186
         _Version        =   393216
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   7350
      TabIndex        =   2
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Close"
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
      Picture         =   "frmReports.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   7350
      TabIndex        =   1
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&View"
      AccessKey       =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmReports.frx":077A
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_oSkin As clsFormSkin

Private p_oAppDrivr As clsAppDriver
Private p_oReport As Recordset

Private pbLoaded As Boolean
Private pbPreview As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Preview() As Boolean
   Preview = pbPreview
End Property

Property Get ReportID() As String
   ReportID = p_oReport.Fields("sReportID")
End Property

Property Get ReportName() As String
   ReportName = p_oReport.Fields("sReportNm")
End Property

Property Get ReportLibrary() As String
   ReportLibrary = p_oReport.Fields("sRepLibxx")
End Property

Property Get ReportClass() As String
   ReportClass = p_oReport.Fields("sRepClass")
End Property

Property Get LogReport() As Boolean
   LogReport = p_oReport.Fields("cLogRepxx") = xeYes
End Property

Property Get SaveReport() As Boolean
   SaveReport = p_oReport.Fields("cSaveRepx") = xeYes
End Property

Private Sub cmdButton_Click(Index As Integer)
   With MSFlexGrid1
      If .RowSel > 0 Then
         p_oReport.Move .RowSel - 1, adBookmarkFirst
      End If
   End With
   
   pbPreview = Index = 0
   Me.Hide
End Sub

Private Sub Form_Activate()
   If pbLoaded = False Then LoadList
End Sub

Private Sub Form_Load()
   Dim lsProcName As String
   
   lsProcName = "SearchTransaction"
   ''On Error GoTo errProc
   
   Set p_oSkin = New clsFormSkin
   Set p_oSkin.AppDriver = p_oAppDrivr
   Set p_oSkin.Form = Me
   p_oSkin.ApplySkin xeFormTransDetail
   
   With MSFlexGrid1
      .Cols = 2
      .Rows = 2
      
      .ColWidth(0) = 350
      .ColWidth(1) = .Width - 500
      
      .TextMatrix(0, 1) = "Report Name"
   End With
      
   Exit Sub
   
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
   Set p_oReport = Nothing
End Sub

Private Sub LoadList()
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   lsProcName = "LoadList"
   ''On Error GoTo errProc
   
   lsSQL = "SELECT" & _
               "  sReportID" & _
               ", sReportNm" & _
               ", sRepLibxx" & _
               ", sRepClass" & _
               ", cLogRepxx" & _
               ", cSaveRepx" & _
            " FROM xxxReportMaster" & _
            " WHERE sProdctID LIKE " & strParm("%" & p_oAppDrivr.ProductID & "%") & _
               " AND nUserRght & " & p_oAppDrivr.UserLevel & " > 0" & _
               " AND sRepLibxx = " & strParm("ggcCPAuditRep") & _
            " ORDER BY sReportNm"
            
   Debug.Print lsSQL
   Set p_oReport = New Recordset
   p_oReport.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   With MSFlexGrid1
      .Rows = p_oReport.RecordCount + 1
      
      lnCtr = 1
      Do While p_oReport.EOF = False
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = p_oReport("sReportNm")
                  
         lnCtr = lnCtr + 1
         p_oReport.MoveNext
      Loop
      
      .Row = 1
      .Col = 1
      .ColSel = 1
   End With
   
endPorc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "frmReports", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

