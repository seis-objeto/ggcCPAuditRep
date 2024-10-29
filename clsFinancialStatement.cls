VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinancialStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================================
'  sheryl [ 2015-05-09 04:46 pm ]
'     Start creating this object.

Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmFinancialStatemnt
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

   Set p_oStandard = New frmFinancialStatemnt
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
   .Caption = "Finance Statement"
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

   If .Summary = 0 Then
      If prcPurchase = False Then GoTo endProc
   Else
      If prcReturn = False Then GoTo endProc
   End If
   
   End With

   ProcessReport = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcPurchase() As Boolean
   Dim lsProcName As String
   
   lsProcName = "prcPurchase"
   ''On Error GoTo errProc
   
   With p_oStandard
   setRepProgress 5
   moveRepProgress "Setting Report Retrieval Info..."
      
   If .Category = 0 Then 'units
      psSQL = "SELECT a.sTransNox" & _
                  ", a.dTransact" & _
                  ", d.dTransact xReceived" & _
                  ", e.sTransNox sPOTransx" & _
                  ", j.sBarrCode" & _
                  ", k.sBrandNme" & _
                  ", j.sDescript" & _
                  ", c.sSerialNo" & _
                  ", f.nUnitPrce" & _
                  ", Case a.cTranStat" & _
                     " WHEN '0' THEN 'OPEN' " & _
                     " WHEN '1' THEN 'CLOSED' " & _
                     " WHEN '2' THEN 'POSTED' " & _
                  " End cTranStat" & _
               " FROM CP_Transfer_Master a" & _
               ", CP_Transfer_Detail b" & _
               ", CP_Inventory_Serial c" & _
                     " LEFT JOIN CP_Inventory_Serial_Ledger d" & _
                        " ON c.sSerialID = d.sSerialID AND d.sSourceCd = 'CPDA'" & _
                     " LEFT JOIN CP_PO_Receiving_Serial e" & _
                        " ON c.sSerialID = e.sSerialID" & _
                     " LEFT JOIN CP_PO_Receiving_Detail f " & _
                        " ON e.sTransNox = f.sTransNox AND e.nEntryNox = f.nEntryNox" & _
               ", CP_Inventory j" & _
               ", CP_Brand k"
               
      psSQL = psSQL & _
               " Where a.sTransNox = b.sTransNox" & _
               " AND b.sSerialID = c.sSerialID" & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
               " AND c.sStockIDx = j.sStockIDx" & _
               " AND j.sBrandIDx = k.sBrandIDx" & _
               " AND a.sDestinat = " & strParm(.BranchCd) & _
               " AND a.cTranStat <> '3' "
   Else 'accessories
      psSQL = "SELECT * FROM " & _
                  "(SELECT a.sTransNox" & _
                     ", a.dTransact" & _
                     ", b.sTransNox sPOTransx" & _
                     ", b.dTransact xReceived" & _
                     ", a.sStockIDx" & _
                     ", a.sBarrcode" & _
                     ", a.sDescript" & _
                     ", a.nQuantity" & _
                     ", b.nUnitPrce" & _
                     ", CASE a.cTranStat" & _
                        " WHEN '0' THEN 'OPEN'" & _
                        " WHEN '1' THEN 'CLOSED' " & _
                        " WHEN '2' THEN 'POSTED' " & _
                     " END cTranStat" & _
                  " FROM " & _
                        "(SELECT a.sTransNox" & _
                           ", a.dTransact" & _
                           ", b.sStockIDx" & _
                           ", e.sBarrcode" & _
                           ", e.sDescript" & _
                           ", b.nQuantity" & _
                           ", a.cTranStat"
      psSQL = psSQL & _
                        " FROM CP_Transfer_Master a" & _
                        ", CP_Transfer_Detail b" & _
                        ", CP_Inventory e" & _
                        " WHERE a.sTransNox = b.sTransNox " & _
                        " AND b.sStockIDx = e.sStockIDx" & _
                        " AND ( b.sSerialID = '' OR b.sSerialID IS NULL )" & _
                        " AND e.sCategID1 <> 'C001006'" & _
                        " AND a.sDestinat = " & strParm(.BranchCd) & _
                        " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & ") a " & _
                     " LEFT JOIN " & _
                        " ( SELECT a.sTransNox, a.dTransact, b.sStockIDx, b.nUnitPrce" & _
                        " FROM CP_PO_Receiving_Master a, CP_PO_Receiving_Detail b" & _
                        " WHERE a.sTransNox = b.sTransNox" & _
                        " AND a.cTranStat <> '3'" & _
                        " AND b.nUnitPrce > 0 ) b" & _
                           " ON a.sStockIDx = b.sStockIDx AND a.dTransact >= b.dTransact " & _
                     " ORDER BY a.sTransNox, a.sStockIDx, b.dTransact DESC ) xSource" & _
               " GROUP BY sTransNox, sStockIDx"
   End If
   
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
   
'   openSource
   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 50
      .Fields.Append "sField02", adVarChar, 50
      .Fields.Append "sField03", adVarChar, 20
      .Fields.Append "sField04", adVarChar, 64
      .Fields.Append "sField05", adVarChar, 64
      .Fields.Append "sField06", adVarChar, 64
      .Fields.Append "sField07", adVarChar, 64
      .Fields.Append "nField01", adInteger
      .Fields.Append "lField01", adCurrency
      
      .Open
   End With
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sBarrCode") & " » " & p_oRawSource("sDescript")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If
         
         .AddNew
         If p_oStandard.Category = 0 Then
            .Fields("sField01").Value = p_oRawSource("sTransNox")
            .Fields("sField02").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
            .Fields("sField03").Value = IFNull(Right(p_oRawSource("sPOTransx"), 10), "")
            .Fields("sField04").Value = p_oRawSource("sBarrCode")
            .Fields("sField05").Value = p_oRawSource("sSerialNo")
            .Fields("sField06").Value = p_oRawSource("sBrandNme")
            .Fields("sField07").Value = p_oRawSource("cTranStat")
            .Fields("nField01").Value = 1
            .Fields("lField01").Value = IFNull(p_oRawSource("nUnitPrce"), 0#)
         Else
            .Fields("sField01").Value = p_oRawSource("sTransNox")
            .Fields("sField02").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
            .Fields("sField03").Value = IFNull(Right(p_oRawSource("sPOTransx"), 10), "")
            .Fields("sField04").Value = p_oRawSource("sBarrCode")
            .Fields("sField05").Value = ""
            .Fields("sField06").Value = p_oRawSource("sDescript")
            .Fields("sField07").Value = p_oRawSource("cTranStat")
            .Fields("nField01").Value = p_oRawSource("nQuantity")
            .Fields("lField01").Value = IFNull(p_oRawSource("nUnitPrce"), 0#)
         End If
         p_oRawSource.MoveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
   End With
   
   If prcReportObject = False Then GoTo endProc
   
   setRepProgress 1
   moveRepProgress "Loading Report..."
  
   prcPurchase = True
   
endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcReturn() As Boolean
   Dim lsProcName As String
   
   lsProcName = "prcReturn"
   ''On Error GoTo errProc
   
   With p_oStandard
   setRepProgress 5
   moveRepProgress "Setting Report Retrieval Info..."
      
   If .Category = 0 Then 'units
      psSQL = "SELECT a.sTransNox" & _
                  ", a.dTransact" & _
                  ", f.sTransNox sPOTransx" & _
                  ", j.sBarrCode" & _
                  ", j.sDescript" & _
                  ", a.dReceived" & _
                  ", c.sSerialNo" & _
                  ", b.nQuantity" & _
                  ", f.nUnitPrce" & _
                  ", k.sBrandNme" & _
                  ", CASE a.cTranStat" & _
                     " WHEN '0' THEN 'OPEN' " & _
                     " WHEN '1' THEN 'CLOSED' " & _
                     " WHEN '2' TEHN 'POSTED' " & _
                  " END cTranStat" & _
               " FROM CP_Transfer_Master a" & _
               ", CP_Transfer_Detail b" & _
               ", CP_Inventory_Serial c" & _
                     " LEFT JOIN CP_Inventory_Serial_Ledger d" & _
                        " ON c.sSerialID = d.sSerialID AND d.sSourceCd = 'CPDA' " & _
                     " LEFT JOIN CP_PO_Receiving_Serial e" & _
                        " ON d.sSourceNo = e.sTransNox AND d.sSerialID = e.sSerialID" & _
                     " LEFT JOIN CP_PO_Receiving_Detail f" & _
                        " ON e.sTransNox = f.sTransNox AND e.nEntryNox = f.nEntryNox" & _
               ", CP_Inventory j"
                              
      psSQL = psSQL & _
               ", CP_Brand k" & _
               " WHERE a.sTransNox = b.sTransNox" & _
               " AND b.sSerialID = c.sSerialID" & _
               " AND c.sStockIDx = j.sStockIDx" & _
               " AND j.sBrandIDx = k.sBrandIDx" & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
               " AND a.sTransNox LIKE " & strParm(.BranchCd, "%") & _
               " AND a.cTranStat <> '3'"
   Else 'accessories
      psSQL = "SELECT a.sTransNox" & _
                  ", a.dTransact" & _
                  ", a.dReceived" & _
                  ", c.sStockIDx" & _
                  ", c.sBarrCode" & _
                  ", c.sDescript" & _
                  ", b.nQuantity" & _
                  ", b.nUnitPrce" & _
                  ", CASE a.cTranStat" & _
                     " WHEN '0' THEN 'OPEN' " & _
                     " WHEN '1' THEN 'CLOSED' " & _
                     " WHEN '2' TEHN 'POSTED' " & _
                  " END cTranStat" & _
               " FROM CP_Transfer_Master a " & _
               ", CP_Transfer_Detail b" & _
               ", CP_Inventory c" & _
               " WHERE a.sTransNox = b.sTransNox" & _
               " AND b.sStockIDx = c.sStockIDx" & _
               " AND c.sCategID1 <> 'C001006'" & _
               " AND ( b.sSerialID IS NULL OR b.sSerialID = '' )" & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
               " AND a.sTransNox LIKE " & strParm(.BranchCd + "%") & _
               " AND a.cTranStat <> '3'"
   End If
   
   End With
   moveRepProgress "Retrieving Records..."
   p_oRawSource.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
   
   moveRepProgress "Processing Records..."
   If p_oRawSource.EOF Then
      CloseReport
      MsgBox "No Record has Pass the Given Report Critera!!!", vbInformation, "Notice"
      GoTo endProc
   End If
   
'   openSource
   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 50
      .Fields.Append "sField02", adVarChar, 50
      .Fields.Append "sField03", adVarChar, 20
      .Fields.Append "sField04", adVarChar, 64
      .Fields.Append "sField05", adVarChar, 64
      .Fields.Append "sField06", adVarChar, 64
      .Fields.Append "sField07", adVarChar, 64
      .Fields.Append "nField01", adInteger
      .Fields.Append "lField01", adCurrency
      .Open
   End With
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sBarrCode") & " » " & p_oRawSource("sDescript")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If
         
         .AddNew
         If p_oStandard.Category = 0 Then
            .Fields("sField01").Value = p_oRawSource("sTransNox")
            .Fields("sField02").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
            .Fields("sField03").Value = IFNull(Right(p_oRawSource("sPOTransx"), 10), "")
            .Fields("sField04").Value = p_oRawSource("sBarrCode")
            .Fields("sField05").Value = p_oRawSource("sSerialNo")
            .Fields("sField06").Value = p_oRawSource("sBrandNme")
            .Fields("sField07").Value = p_oRawSource("cTranStat")
            .Fields("nField01").Value = 1
            .Fields("lField01").Value = p_oRawSource("nUnitPrce")
         Else
            .Fields("sField01").Value = p_oRawSource("sTransNox")
            .Fields("sField02").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
            .Fields("sField03").Value = ""
            .Fields("sField04").Value = p_oRawSource("sBarrCode")
            .Fields("sField05").Value = ""
            .Fields("sField06").Value = p_oRawSource("sDescript")
            .Fields("sField07").Value = p_oRawSource("cTranStat")
            .Fields("nField01").Value = p_oRawSource("nQuantity")
            .Fields("lField01").Value = p_oRawSource("nUnitPrce")
         End If
         p_oRawSource.MoveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
   End With
   
   If prcReportObject = False Then GoTo endProc
   
   setRepProgress 1
   moveRepProgress "Loading Report..."
  
   prcReturn = True
   
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
      .xLogError Err.Number, Err.Description, "clsUnreleaseGAway", lsProcName, Erl
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
   
   With p_oStandard
      Select Case .Summary
      Case 0
         If .Category = 0 Then
            lsAddHead = " - " & .optSummary(0).Caption & "(Units)"
         Else
            lsAddHead = " - " & .optSummary(0).Caption & "(Accessories)"
         End If
      Case 1
         If .Category = 0 Then
            lsAddHead = " - " & .optSummary(1).Caption & "(Units)"
         Else
            lsAddHead = " - " & .optSummary(1).Caption & "(Accessories)"
         End If
      End Select
   End With
   lsAddHead = p_oStandard.optSummary(0).Caption & "-" & p_oStandard.optCategory(0).Caption
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




