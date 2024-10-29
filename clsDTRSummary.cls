VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDTRSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'  sheryl [ 2015-03-11 02:04 pm ]

Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmDTRRepCriteria
Private p_oReport As Report

Private psReportID As String
Private psReportNm As String
Private psReportFl As String
Private psReportHd As String
Private psSQL As String

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Source() As Report
   Set Source = p_oReport
End Property

Sub CloseReport()
   p_oProgress.CloseProgress
   Set p_oStandard = Nothing
End Sub

Function InitReport(ByVal ReportID As String, ByVal ReportName As String) As Boolean
   Dim lsProcName As String

   lsProcName = "ProcessReport"
   ''On Error GoTo errProc

   psReportID = ReportID
   psReportNm = ReportName

   Set p_oStandard = New frmDTRRepCriteria
   Set p_oStandard.AppDriver = p_oAppDrivr

   InitReport = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & ReportID _
                       & ", " & ReportName & " )"
End Function

Function ProcessReport() As Boolean
   Dim lors As Recordset, loTemp As Recordset
   Dim lsProcName As String
   Dim lnEntryNo As Integer

   lsProcName = "ProcessReport"
   ''On Error GoTo errProc

   With p_oStandard
      .Caption = "DTR SUMMARY"
      .Show 1

      If .Cancelled Then
      MsgBox "Report Generation was Cancelled", vbInformation, "Notice"
      GoTo endProc
      End If
   
      lnEntryNo = 1
      With p_oProgress
         .InitProgress "Processing...", 5, 3
         .PrimaryRemarks = "Processing Report"
         .MoveProgress "Setting Retrieval Info..."
   
         psSQL = "SELECT * FROM xxxReportDetail" & _
                     " WHERE sReportID = " & strParm(psReportID) & _
                        " and nEntryNox = " & lnEntryNo
   
         .MoveProgress "Retriving Report Specification..."
   
         p_oRawSource.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
   
         .MoveProgress "Processing Report Specification..."
      End With

      If p_oRawSource.EOF Then
      CloseReport
      MsgBox "Unable to Retrieve Report Info..." & psSQL, vbCritical, "Warning"
      GoTo endProc
      End If
   
      psReportFl = p_oRawSource("sFileName")
      psReportHd = p_oRawSource("sReportHd")
      p_oRawSource.Close
   
      If prcDTRSummary = False Then GoTo endProc
   End With

   ProcessReport = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcDTRSummary() As Boolean
   Dim lsProcName As String
   Dim lsBranch As String
   Dim lsDateFr As String
   Dim lsDateTh As String
   

   lsProcName = "prcDTRSummary"
   ''On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
      
   If p_oAppDrivr.UserLevel = xeAudit Or p_oAppDrivr.UserLevel = xeEngineer Then
      lsBranch = IIf(.Branch = "", "", " AND a.sBranchCd = " & strParm(.Branch))
   Else
      lsBranch = " AND sBranchCd = " & strParm(p_oAppDrivr.BranchCode)
   End If
   
   lsDateFr = Format(.DateFrom, "YYYYMMDD")
   lsDateTh = Format(.DateThru, "YYYYMMDD")
   
      psSQL = "SELECT " & _
               " a.sBranchCd " & _
               ", b.sBranchNm" & _
               ", DATE_FORMAT(a.sTranDate, '%Y-%m-%d') `sTranDate`  " & _
               ", a.nTotalSle" & _
               ", a.nAdvPaymt" & _
               ", a.nCleardCk" & _
               ", a.nDiffCash" & _
               ", a.nDiffChk" & _
               ", a.nDepositd" & _
               ", a.nActChDep" & _
               ", a.nActCkDep" & _
               ", a.nCrdtCard" & _
               ", a.cPostedxx" & _
            " FROM DTR_Summary a" & _
            ", Branch b " & _
            " WHERE a.sBranchCd = b.sBranchCd" & _
            " AND a.sTrandate BETWEEN " & strParm(lsDateFr) & " AND " & strParm(lsDateTh) & _
            lsBranch & _
            " ORDER BY a.sBranchCd, a.sTranDate"
   End With
   Debug.Print psSQL
   moveRepProgress "Retrieving Records..."
   p_oRawSource.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
   
   moveRepProgress "Processing Records..."
   If p_oRawSource.EOF Then
      CloseReport
         MsgBox "No Record has Pass the Given Report Critera!!!", vbInformation, "Notice"
         GoTo endProc
   End If
   
   ' open the report source recordset...

   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 150
      .Fields.Append "sField02", adVarChar, 50
      .Fields.Append "sField03", adVarChar, 50
      .Fields.Append "lField01", adCurrency
      .Fields.Append "lField02", adCurrency
      .Fields.Append "lField03", adCurrency
      .Fields.Append "lField04", adCurrency
      .Fields.Append "lField05", adCurrency
      .Fields.Append "lField06", adCurrency
      .Fields.Append "lField07", adCurrency
      .Fields.Append "lField08", adCurrency
      .Fields.Append "lField09", adCurrency
            
      .Open
   End With
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sBranchNm") & " » " & p_oRawSource("sTranDate")) = False Then
            CloseReport
               MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
                  GoTo endProc
         End If
         
         
         
      .AddNew
      .Fields("sField01").Value = p_oRawSource("sBranchNm")
      .Fields("sField02").Value = p_oRawSource("sTranDate")
      .Fields("lField01").Value = p_oRawSource("nAdvPaymt")
      .Fields("lField02").Value = p_oRawSource("nCleardCk")
      .Fields("lField03").Value = p_oRawSource("nDiffCash")
      .Fields("lField04").Value = p_oRawSource("nDiffChk")
      .Fields("lField05").Value = p_oRawSource("nTotalSle")
      .Fields("lField06").Value = p_oRawSource("nDepositd")
      .Fields("lField07").Value = IFNull(p_oRawSource("nActChDep"), 0#)
      .Fields("lField08").Value = IFNull(p_oRawSource("nActCkDep"), 0#)
      .Fields("lField09").Value = IFNull(p_oRawSource("nCrdtCard"), 0#)
      Select Case p_oRawSource("cPostedxx")
      Case 0
         .Fields("sField03").Value = "OPEN"
      Case 1
         .Fields("sField03").Value = "CONFIRMED"
      Case 2
         .Fields("sField03").Value = "POSTED"
      Case 4
         .Fields("sField03").Value = "VERIFIED"
      End Select
      
      p_oRawSource.MoveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
   End With
   
   If prcReportObject = False Then GoTo endProc
   
   setRepProgress 1
   moveRepProgress "Loading Report..."
  
   prcDTRSummary = True
   
endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

   
Private Sub setRepProgress(ByVal lnMaxValue As Long, Optional lvRemarks As Variant)
   With p_oProgress
      .SecMaxValue = lnMaxValue
      If Not IsMissing(lvRemarks) Then
         .PrimaryRemarks = lvRemarks
      End If
   End With
End Sub

Private Function moveRepProgress(ByVal lsSecRemarks As String, Optional lvPriRemarks As Variant) As Boolean
   With p_oProgress
      If Not IsMissing(lvPriRemarks) Then
         moveRepProgress = .MoveProgress(lsSecRemarks, lvPriRemarks)
      Else
         moveRepProgress = .MoveProgress(lsSecRemarks)
      End If
   End With
End Function

Private Sub Class_Initialize()
   Set p_oProgress = New clsSpeedometer
   Set p_oRepSource = New Recordset
   Set p_oRawSource = New Recordset
End Sub

Private Sub Class_Terminate()
   Set p_oProgress = Nothing
   Set p_oRepSource = Nothing
   Set p_oReport = Nothing
   Set p_oRawSource = Nothing
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "clsDTRSummary", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub


Private Function prcReportObject() As Boolean
   Dim loRepApp As New Application
   Dim lsProcName As String
   Dim lsSubHead As String
   Dim lsAddHead As String
   
   lsProcName = "prcReportObject"
   ''On Error GoTo errProc
   
   setRepProgress 2
   moveRepProgress "Initializing Report File..."
   
   Set loRepApp = New Application
   Set p_oReport = loRepApp.OpenReport(p_oAppDrivr.AppPath & _
                        "\Reports\Definitions\" & psReportFl & ".rpt")
   Set loRepApp = Nothing
   
   lsAddHead = ""
   moveRepProgress "Assigning Record to Report..."
   
   With p_oStandard
      lsSubHead = "As of " & Format(.DateFrom, "MMMM DD, YYYY") & _
                     " to " & Format(.DateThru, "MMMM DD, YYYY")
   End With
   
   With p_oReport
      .DiscardSavedData
      .FieldMappingType = crAutoFieldMapping
      .Database.SetDataSource p_oRepSource
   End With
   
   With p_oAppDrivr
      p_oReport.Sections("RH").ReportObjects("txtCompany").SetText .ClientName
      p_oReport.Sections("RH").ReportObjects("txtAddress").SetText .Address & _
                  ", " & .TownCity & ", " & .Province & " " & .ZippCode
                  
      p_oReport.Sections("PH").ReportObjects("txtHeading1").SetText psReportHd & " " & lsAddHead
      p_oReport.Sections("PH").ReportObjects("txtHeading2").SetText lsSubHead
               
      p_oReport.Sections("PF").ReportObjects("txtRptUser").SetText .UserName
   End With
   
   prcReportObject = True
   
endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function








