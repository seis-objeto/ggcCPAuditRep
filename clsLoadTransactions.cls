VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoadTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'she 2016-04-22
'  sheryl 2015-09-03 1:47 pm

Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmCPLoadCriteria
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

   Set p_oStandard = New frmCPLoadCriteria
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
      .Caption = "Load Transactions"
      .Show 1

      If .Cancelled Then
      MsgBox "Report Generation was Cancelled", vbInformation, "Notice"
      GoTo endProc
      End If
      
      lnEntryNo = .Presentation + 1
      
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
   
      If prcLoadTrans = False Then GoTo endProc
   End With

   ProcessReport = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcLoadTrans() As Boolean
   Dim lsProcName As String
   
   lsProcName = "prcLoadTrans"
   'On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
      
      psSQL = "SELECT a.dTransact " & _
                  ", b.sBarrCode" & _
                  ", a.sSourceCd" & _
                  ", a.sSourceNo" & _
                  ", a.nAmtInxxx" & _
                  ", a.nAmtOutxx" & _
                  ", a.nAmtOnHnd" & _
                  ", d.nTranAmtx" & _
                  ", e.nAmtPaidx" & _
                  ", a.sStockIDx" & _
               " FROM CP_Load_Matrix_Ledger a" & _
                        " LEFT JOIN CP_SO_Eload d ON a.sStockIDx = d.sStockIDx" & _
                           " AND a.sSourceNo = d.sTransNox AND a.sSourceCd = 'CPSl' " & _
                        " LEFT JOIN CP_SO_Load_Wallet e ON a.sStockIDx = d.sStockIDx" & _
                           " AND a.sSourceNo = e.sTransNox AND a.sSourceCd = 'CPSl' " & _
                  ", CP_Load_Matrix b" & _
                  ", Branch c" & _
               " WHERE a.sBranchCd = c.sBranchCd" & _
               " AND a.sStockIDx = b.sStockIDx" & _
               " AND a.sStockIDx = " & strParm(.Specify) & _
               " AND a.sBranchcd = " & strParm(p_oAppDrivr.BranchCode) & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
               " ORDER BY a.dTransact, a.nLedgerNo"
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
      .Fields.Append "sField01", adVarChar, 250
      .Fields.Append "sField02", adVarChar, 200
      .Fields.Append "sField03", adVarChar, 200
      .Fields.Append "sField04", adVarChar, 200
      .Fields.Append "sField05", adVarChar, 200
      .Fields.Append "lField01", adInteger
      .Fields.Append "lField02", adCurrency
      .Fields.Append "lField03", adCurrency
      .Fields.Append "lField04", adCurrency
            
      .Open
   End With
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sSourceNo") & " » " & p_oRawSource("sSourceCd")) = False Then
            CloseReport
               MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
                  GoTo endProc
         End If
      
      Dim lorec As Recordset
      Dim lssSQL As String
      Dim lsMatrix As String
      Dim lsCode As String
      
      ' this case is for preparation that in case the auditor need the other info.
      'she
      Select Case p_oRawSource("sSourceCd")
      Case "CPDA"
         Set lorec = New Recordset
         lssSQL = "SELECT a.sTransNox" & _
                     ", 'Load Receiving' sMatrixNm" & _
                  " FROM CP_Load_Receiving_Master a" & _
                  ", CP_Load_Receiving_Detail b" & _
                  " WHERE a.sTransNox = " & strParm(p_oRawSource("sSourceNo")) & _
                  " AND b.sStockIDx = " & strParm(p_oRawSource("sStockIDx")) & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled)
         lorec.Open lssSQL, p_oAppDrivr.Connection, , , adCmdText
         If Not lorec.EOF Then
            lsMatrix = lorec("sMatrixNm")
         Else
            lsMatrix = ""
         End If
      Case "CPDv"
         Set lorec = New Recordset
         lssSQL = "SELECT a.sTransNox" & _
                  ", 'Load Transfer' sMatrixNm" & _
                  " FROM CP_Load_Transfer_Master a" & _
                  ", CP_Load_Transfer_Detail b" & _
                  " WHERE a.sTransNox = " & strParm(p_oRawSource("sSourceNo")) & _
                  " AND b.sStockIDx = " & strParm(p_oRawSource("sStockIDx")) & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled)
         lorec.Open lssSQL, p_oAppDrivr.Connection, , , adCmdText
         
         If Not lorec.EOF Then
            lsMatrix = lorec("sMatrixNm")
         Else
            lsMatrix = ""
         End If
      Case "CPCv"
         Set lorec = New Recordset
         lssSQL = "SELECT a.sTransNox" & _
                  ", 'Load Convert Transfer' sMatrixNm" & _
                  " FROM CP_Convert_Wallet a" & _
                  " WHERE a.sTransNox = " & strParm(p_oRawSource("sSourceNo")) & _
                  " AND a.sStockIdx = " & strParm(p_oRawSource("sStockIDx")) & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled)
         lorec.Open lssSQL, p_oAppDrivr.Connection, , , adCmdText
         Debug.Print lssSQL
         If Not lorec.EOF Then
            lsMatrix = lorec("sMatrixNm")
         Else
            lsMatrix = ""
         End If
      Case "CPDl"
         Set lorec = New Recordset
         lssSQL = "SELECT a.sTransNox" & _
                  ", 'Load Received' sMatrixNm" & _
                  " FROM CP_Load_Transfer_Master a" & _
                  ", CP_Load_Transfer_Detail b" & _
                  " WHERE a.sTransNox = " & strParm(p_oRawSource("sSourceNo")) & _
                  " AND b.sStockIDx = " & strParm(p_oRawSource("sStockIDx")) & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled)
         lorec.Open lssSQL, p_oAppDrivr.Connection, , , adCmdText
         
         If Not lorec.EOF Then
            lsMatrix = lorec("sMatrixNm")
         Else
            lsMatrix = ""
         End If
      Case "CPCr"
         Set lorec = New Recordset
         lssSQL = "SELECT a.sTransNox" & _
                  ", 'Load Convert Received' sMatrixNm" & _
                  " FROM CP_Convert_Wallet a" & _
                  " WHERE a.sTransNox = " & strParm(p_oRawSource("sSourceNo")) & _
                  " AND a.sStockCon = " & strParm(p_oRawSource("sStockIDx")) & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled)
         lorec.Open lssSQL, p_oAppDrivr.Connection, , , adCmdText
         Debug.Print lssSQL
         If Not lorec.EOF Then
            lsMatrix = lorec("sMatrixNm")
         Else
            lsMatrix = ""
         End If
      Case "CPSl"
         Set lorec = New Recordset
         lssSQL = "SELECT a.sTransNox" & _
                  ", b.sMatrixNm" & _
                  " FROM CP_SO_Eload a" & _
                  ", ELoad_Matrix b" & _
                  " WHERE a.sTransNox = " & strParm(p_oRawSource("sSourceNo")) & _
                  " AND a.sMatrixID = b.sMatrixID" & _
                  " AND a.sStockIDx = " & strParm(p_oRawSource("sStockIDx")) & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled)
         lorec.Open lssSQL, p_oAppDrivr.Connection, , , adCmdText
         
         If Not lorec.EOF Then
            lsMatrix = lorec("sMatrixNm")
         Else
            lsMatrix = ""
         End If
      Case "CPCm", "CPDm"
         Set lorec = New Recordset
         lssSQL = "SELECT sTransNox" & _
                  ", 'Load Adj' sMatrixNm" & _
                  " FROM CP_Load_Adjustment a" & _
                  " WHERE sTransNox = " & strParm(p_oRawSource("sSourceNo")) & _
                  " AND sStockIDx = " & strParm(p_oRawSource("sStockIDx")) & _
                  " AND cTranStat <> " & strParm(xeStateCancelled)
         lorec.Open lssSQL, p_oAppDrivr.Connection, , , adCmdText
         
         If Not lorec.EOF Then
            lsMatrix = lorec("sMatrixNm")
         Else
            lsMatrix = ""
         End If
      End Select
      
      .AddNew
      .Fields("sField01").Value = Format(p_oRawSource("dTransact"), "MM/DD/YYYY")
'      .Fields("sField02").Value = p_oRawSource("sBarrCode")
      .Fields("sField03").Value = lsMatrix
      .Fields("sField04").Value = p_oRawSource("sSourceCd")
      .Fields("sField05").Value = p_oRawSource("sSourceNo")
      .Fields("lField01").Value = p_oRawSource("nAmtInxxx")
      .Fields("lField02").Value = p_oRawSource("nAmtOutxx")
      .Fields("lField03").Value = p_oRawSource("nAmtOnHnd")
      .Fields("lField04").Value = IFNull(p_oRawSource("nTranAmtx"), IFNull(p_oRawSource("nAmtPaidx"), 0))

      p_oRawSource.MoveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
   End With
   
   If prcReportObject = False Then GoTo endProc
   
   setRepProgress 1
   moveRepProgress "Loading Report..."
  
   prcLoadTrans = True
   
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
      .xLogError Err.Number, Err.Description, "clsLoadTransactions", lsProcName, Erl
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

