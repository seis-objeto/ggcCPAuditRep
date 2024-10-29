VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPTransferSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================================
'  sheryl [ 06/28/2014 09:50 am ]

Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmTransferSummary
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
   '''On Error GoTo errProc

   psReportID = ReportID
   psReportNm = ReportName

   Set p_oStandard = New frmTransferSummary
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
   '''On Error GoTo errProc

   With p_oStandard
      .Caption = "CP Transfer Summary"
      
      If p_oAppDrivr.BranchCode = "C001" Or p_oAppDrivr.BranchCode = "C0W1" Or p_oAppDrivr.BranchCode = "C0W2" Then
         .txtBranch = ""
      Else
         .txtBranch = p_oAppDrivr.BranchName
      End If
      
      .Show 1

      If .Cancelled Then
      MsgBox "Report Generation was Cancelled", vbInformation, "Notice"
      GoTo endProc
      End If
      If .Presentation = 0 Then
         lnEntryNo = 1
      Else
         lnEntryNo = 2
      End If
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
   
      If prcCPTransferSum = False Then GoTo endProc
   End With

   ProcessReport = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcCPTransferSum() As Boolean
   Dim lsProcName As String
   Dim lsStatus As String
   Dim lsBranch As String

   lsProcName = "prcCPTransferSum"
   'On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
   
   If .ReceivedType = 0 Then 'Transfer
      If p_oAppDrivr.BranchCode = "C001" Or p_oAppDrivr.BranchCode = "C0W1" Or p_oAppDrivr.BranchCode = "C0W2" Then
         lsBranch = IIf(.Branch = "", "", " AND a.sTransNox LIKE " & strParm(.Branch + "%"))
      Else
         lsBranch = " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode + "%")
      End If
   Else
      If p_oAppDrivr.BranchCode = "C001" Or p_oAppDrivr.BranchCode = "C0W1" Or p_oAppDrivr.BranchCode = "C0W2" Then
         lsBranch = IIf(.Branch = "", "", " AND a.sDestinat = " & strParm(.Branch))
      Else
         lsBranch = " AND a.sDestinat LIKE " & strParm(p_oAppDrivr.BranchCode)
      End If
   End If
   
   If .Status = 0 Then
      lsStatus = " AND a.cTranStat = '2' "
   ElseIf .Status = 2 Then
      lsStatus = " AND a.cTranStat IN('0','1','2') "
   Else
      lsStatus = " AND a.cTranStat IN('0','1') "
   End If

  psSQL = "SELECT" & _
               " l.sAreaDesc" & _
               ", a.sTransNox " & _
               ", e.sBranchNm " & _
               ", a.dTransact " & _
               ", a.dReceived " & _
               ", c.sBarrCode " & _
               ", i.sBrandNme " & _
               ", h.sModelNme " & _
               ", h.sModelCde " & _
               ", j.sColorNme " & _
               ", b.nQuantity " & _
               ", d.sCategrNm " & _
               ", f.sBranchNm `sDestinat`" & _
               ", CASE a.cTranStat " & _
                  " WHEN '0' THEN 'OPEN' " & _
                  " WHEN '1' THEN 'CLOSED' " & _
                  " WHEN '2' THEN 'POSTED' " & _
                  " WHEN '3' THEN 'CANCELLED' " & _
               " END cTranStat " & _
               ", a.sRemarksX " & _
               ", g.cUnitType" & _
               ", g.sSerialNo"
      psSQL = psSQL & _
            " FROM CP_Transfer_Master a" & _
                  " LEFT JOIN Branch f ON a.sDestinat  = f.sBranchCd" & _
            ", CP_Transfer_Detail b" & _
                  " LEFT JOIN CP_Inventory_Serial g ON b.sSerialID = g.sSerialID" & _
            ", CP_Inventory c" & _
                  " LEFT JOIN CP_Model h ON c.sModelIDx = h.sModelIDx" & _
                  " LEFT JOIN CP_Brand i ON c.sBrandIDx = i.sBrandIDx" & _
                  " LEFT JOIN Color j ON c.sColorIDx = j.sColorIDx" & _
            ", Category d" & _
            ", Branch e" & _
                  " LEFT JOIN Branch_Others k ON e.sBranchCd = k.sBranchCd" & _
                  " LEFT JOIN Branch_Area l ON k.sAreaCode = l.sAreaCode" & _
            " WHERE a.sTransNox = b.sTransNox" & _
            " AND b.sStockIDx = c.sStockIDx" & _
            " AND c.sCategID1 = d.sCategrID" & _
            " AND LEFT(a.sTransNox,4) = e.sBranchCd" & _
            " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
         IIf(.Category = "", "", " AND c.sCategID1 = " & strParm(.Category)) & _
            " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
            lsStatus & _
            lsBranch & _
            " ORDER BY l.sAreaDesc, e.sBranchCd, a.dTransact, c.sCategID1, i.sBrandIDx, h.sModelIDx"
   'IIf(.Presentation = 0, " GROUP BY b.sTransNox", " GROUP BY b.sTransNox, b.sStockIDx")
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
      .Fields.Append "sField01", adVarChar, 200
      .Fields.Append "sField02", adVarChar, 200
      .Fields.Append "sField03", adVarChar, 200
      .Fields.Append "sField04", adVarChar, 500
      .Fields.Append "sField05", adVarChar, 200
      .Fields.Append "sField06", adVarChar, 200
      .Fields.Append "sField07", adVarChar, 200
      .Fields.Append "sField08", adVarChar, 150
      .Fields.Append "sField09", adVarChar, 150
      .Fields.Append "sField10", adVarChar, 150
      .Fields.Append "sField11", adVarChar, 150
      .Fields.Append "sField12", adVarChar, 150
      .Fields.Append "sField13", adVarChar, 150
      .Fields.Append "sField14", adVarChar, 150
      .Fields.Append "sField15", adVarChar, 200
      .Fields.Append "nField01", adInteger
      
      .Open
   End With
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sBranchNm") & " » " & p_oRawSource("sTransNox")) = False Then
            CloseReport
               MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
                  GoTo endProc
         End If
         
      .AddNew
      
      If p_oStandard.Presentation = 0 Then
         .Fields("sField01").Value = p_oRawSource("sAreaDesc")
         .Fields("sField02").Value = p_oRawSource("sBranchNm")
         .Fields("sField14").Value = p_oRawSource("sCategrNm")
         .Fields("sField07").Value = p_oRawSource("sBarrCode")
         .Fields("sField08").Value = IFNull(p_oRawSource("sBrandNme"), "")
         .Fields("sField09").Value = IFNull(p_oRawSource("sModelCde"), "")
         .Fields("sField10").Value = IFNull(p_oRawSource("sModelNme"), "")
         .Fields("sField11").Value = IFNull(p_oRawSource("sColorNme"), "")
         .Fields("nField01").Value = p_oRawSource("nQuantity")
      Else
         .Fields("sField01").Value = IFNull(p_oRawSource("sAreaDesc"), "")
         .Fields("sField02").Value = IFNull(p_oRawSource("sBranchNm"), "")
         .Fields("sField03").Value = IFNull(p_oRawSource("sTransNox"), "")
         .Fields("sField04").Value = IFNull(p_oRawSource("dTransact"), "")
         .Fields("sField05").Value = IFNull(p_oRawSource("dReceived"), "")
         .Fields("sField06").Value = IFNull(p_oRawSource("sDestinat"), "")
         .Fields("sField07").Value = IFNull(p_oRawSource("sBarrCode"), "")
         .Fields("sField08").Value = IFNull(p_oRawSource("sBrandNme"), "")
         .Fields("sField09").Value = IFNull(p_oRawSource("sModelCde"), "")
         .Fields("sField10").Value = IFNull(p_oRawSource("sModelNme"), "")
         .Fields("sField11").Value = IFNull(p_oRawSource("sColorNme"), "")
         .Fields("nField01").Value = p_oRawSource("nQuantity")
         .Fields("sField12").Value = p_oRawSource("cTranStat")
         Select Case p_oRawSource("cUnitType")
         Case 0, 5
            .Fields("sField13").Value = "Demo"
         Case 1
            .Fields("sField13").Value = "Regular"
         Case 2
            .Fields("sField13").Value = "Supplier"
         Case 3
            .Fields("sField13").Value = "Live"
         Case 4
            .Fields("sField13").Value = "Service"
         Case Else
            .Fields("sField13").Value = "UNKNOWN"
         End Select
         .Fields("sField14").Value = p_oRawSource("sCategrNm")
         .Fields("sField15").Value = IFNull(p_oRawSource("sSerialNo"), "")
      End If
      p_oRawSource.MoveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
   End With
   
   If prcReportObject = False Then GoTo endProc
   
   setRepProgress 1
   moveRepProgress "Loading Report..."
  
   prcCPTransferSum = True
   
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
      .xLogError Err.Number, Err.Description, "clsTransferSummary", lsProcName, Erl
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
   '''On Error GoTo errProc
   
   setRepProgress 2
   moveRepProgress "Initializing Report File..."
   
   Set loRepApp = New Application
   Set p_oReport = loRepApp.OpenReport(p_oAppDrivr.AppPath & _
                        "\Reports\Definitions\" & psReportFl & ".rpt")
   Set loRepApp = Nothing
   
   If p_oStandard.ReceivedType = 0 Then
      lsAddHead = "TRANSFER TO"
   Else
      lsAddHead = "RECEIVED FR"
   End If
   
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










