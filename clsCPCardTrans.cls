VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPCardTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
'' Guanzon Software Engineering Group
'' Guanzon Group of Companies
'' Perez Blvd., Dagupan City
''
''     CP Cards Transaction Report
''
'' Copyright 2008 and Beyond
'' All Rights Reserved
'' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
'' €  may be reproduced or transmitted in any form or  €€                                   €
'' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
'' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
'' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
'' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
'' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
''
'' ==========================================================================================
''  Jheff [ 01/22/2009 02:20 pm ]
''     Start coding this object...
''€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
'Option Explicit
'
'Private p_oAppDrivr As clsAppDriver
'Private p_oRepSource As Recordset
'Private p_oRawSource As Recordset
'Private p_oProgress As clsSpeedometer
'Private p_oStandard As frmDateCriteria
'Private p_oReport As Report
'
'Private psReportID As String
'Private psReportNm As String
'Private psReportFl As String
'Private psReportHd As String
'Private psSQL As String
'
'Property Set AppDriver(oAppDriver As clsAppDriver)
'   Set p_oAppDrivr = oAppDriver
'End Property
'
'Property Get Source() As Report
'   Set Source = p_oReport
'End Property
'
'Sub CloseReport()
'   p_oProgress.CloseProgress
'   Set p_oStandard = Nothing
'End Sub
'
'Function InitReport(ByVal ReportID As String, ByVal ReportName As String) As Boolean
'   Dim lsProcName As String
'
'   lsProcName = "ProcessReport"
'   ''On Error GoTo errProc
'
'   psReportID = ReportID
'   psReportNm = ReportName
'
'   Set p_oStandard = New frmDateCriteria
'   Set p_oStandard.AppDriver = p_oAppDrivr
'
'   InitReport = True
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsProcName & " (" _
'                 & "  " & ReportID _
'                 & ", " & ReportName _
'                 & " )"
'End Function
'
'Function ProcessReport() As Boolean
'   Dim lors As Recordset, loTemp As Recordset
'   Dim lsProcName As String
'   Dim lnEntryNo As Integer
'
'   lsProcName = "ProcessReport"
'   ''On Error GoTo errProc
'
'   With p_oStandard
'      .txtField(0).Text = Format(p_oAppDrivr.ServerDate, "MMM DD, YYYY")
'      .Caption = "Load Detail Criteria"
'      .Show 1
'
'      If .Cancelled Then
'         MsgBox "Report Generation was Cancelled", vbInformation, "Notice"
'         GoTo endProc
'      End If
'
'      lnEntryNo = 1
'      With p_oProgress
'         .InitProgress "Processing...", 5, 3
'         .PrimaryRemarks = "Processing Report"
'         .MoveProgress "Setting Retrieval Info..."
'
'         psSQL = "SELECT * FROM xxxReportDetail" & _
'            " WHERE sReportID = " & strParm(psReportID) & _
'               " and nEntryNox = " & lnEntryNo
'
'         .MoveProgress "Retriving Report Specification..."
'
'         p_oRawSource.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
'
'         .MoveProgress "Processing Report Specification..."
'      End With
'
'      If p_oRawSource.EOF Then
'         CloseReport
'         MsgBox "Unable to Retrieve Report Info..." & psSQL, vbCritical, "Warning"
'         GoTo endProc
'      End If
'
'      psReportFl = p_oRawSource("sFileName")
'      psReportHd = p_oRawSource("sReportHd")
'      p_oRawSource.Close
'
'      If prcLoad = False Then GoTo endProc
'   End With
'
'   ProcessReport = True
'
'endProc:
'   Exit Function
'errProc:
'   CloseReport
'   ShowError lsProcName & "( " & " )"
'End Function
'
'Private Function prcReportObject() As Boolean
'   Dim loRepApp As New Application
'   Dim lsProcName As String
'   Dim lsSubHead As String
'
'   lsProcName = "prcReportObject"
'   ''On Error GoTo errProc
'
'   setRepProgress 2
'   moveRepProgress "Initializing Report File..."
'
'   Set loRepApp = New Application
'   Set p_oReport = loRepApp.OpenReport(p_oAppDrivr.AppPath & _
'                        "\Reports\Definitions\" & psReportFl & ".rpt")
'   Set loRepApp = Nothing
'
'   moveRepProgress "Assigning Record to Report..."
'   With p_oStandard
'        lsSubHead = "As of " & Format(.DateFrom, "MMMM DD, YYYY") & _
'                     " to " & Format(.DateThru, "MMMM DD, YYYY")
'   End With
'
'   With p_oReport
'      .DiscardSavedData
'      .FieldMappingType = crAutoFieldMapping
'      .Database.SetDataSource p_oRepSource
'   End With
'
'   With p_oAppDrivr
'      p_oReport.Sections("RH").ReportObjects("txtCompany").SetText .ClientName
'      p_oReport.Sections("RH").ReportObjects("txtAddress").SetText .Address & _
'                  ", " & .TownCity & ", " & .Province & " " & .ZippCode
'
'      p_oReport.Sections("PH").ReportObjects("txtHeading1").SetText psReportHd
'      p_oReport.Sections("PH").ReportObjects("txtHeading2").SetText lsSubHead
'
'      p_oReport.Sections("PF").ReportObjects("txtRptUser").SetText .UserName
'   End With
'
'   prcReportObject = True
'
'endProc:
'   Exit Function
'errProc:
'   CloseReport
'   ShowError lsProcName & "( " & " )"
'End Function
'
'Private Function prcLoad() As Boolean
'   Dim lsProcName As String
'   Dim lsDescript As String
'
'   lsProcName = "prcLoad"
'   ''On Error GoTo errProc
'
'   With p_oStandard
'      setRepProgress 4
'      moveRepProgress "Setting Report Retrieval Info..."
'
'      psSQL = "SELECT" & _
'                  "  a.sTransNox" & _
'                  ", a.dTransact" & _
'                  ", a.sPhoneNum" & _
'                  ", a.nTranAmtx" & _
'                  ", b.sBarrCode" & _
'                  ", b.sDescript" & _
'                  ", b.sStockIDx" & _
'               " FROM CP_SO_Eload a" & _
'                  ", CP_Load_Matrix b" & _
'               " WHERE a.sStockIDx = b.sStockIDx" & _
'                  " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
'                  " AND a.dTransact Between " & dateParm(.DateFrom) & _
'                     " AND " & dateParm(.DateThru & " 23:59:59") & _
'                  " AND a.cTranStat <> " & strParm(xeStateCancelled)
'
'      psSQL = psSQL & " UNION SELECT" & _
'                  "  a.sTransNox" & _
'                  ", a.dTransact" & _
'                  ", a.sPhoneNum" & _
'                  ", a.nTranAmtx" & _
'                  ", b.sBarrCode" & _
'                  ", b.sDescript" & _
'                  ", b.sStockIDx" & _
'               " FROM CP_SO_Load_Wallet a" & _
'                  ", CP_Load_Matrix b" & _
'               " WHERE a.sStockIDx = b.sStockIDx" & _
'                  " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
'                  " AND a.dTransact Between " & dateParm(.DateFrom) & _
'                     " AND " & dateParm(.DateThru & " 23:59:59") & _
'                  " AND a.cTranStat <> " & strParm(xeStateCancelled)
'   End With
'
'   moveRepProgress "Retrieving Records..."
'
'   p_oRawSource.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
'
'   moveRepProgress "Processing Records..."
'   If p_oRawSource.EOF Then GoTo endProc
'
'   ' open the report source recordset...
'   openSource
'   With p_oRepSource
'      ' after retrieving record process report
'      setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."
'
'      Do While p_oRawSource.EOF = False
'         If moveRepProgress(p_oRawSource("sBarrCode") & " » " & p_oRawSource("sDescript")) = False Then
'            CloseReport
'            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
'            GoTo endProc
'         End If
'
'         .AddNew
'         .Fields("sField01") = p_oRawSource("sStockIDx")
'         .Fields("sField02") = p_oRawSource("sBarrCode") & " " & p_oRawSource("sDescript")
'         .Fields("sField03") = p_oRawSource("sPhoneNum")
'         .Fields("lField01") = p_oRawSource("nTranAmtx")
'
'         p_oRawSource.MoveNext
'      Loop
'      .Sort = "sField01,sField02,lField01"
'   End With
'
'   prcLoad = True
'
'endProc:
'   p_oRawSource.Close
'   Exit Function
'errProc:
'   CloseReport
'   ShowError lsProcName & "( " & " )"
'End Function
'
'Private Sub openSource()
'   Dim lsProcName As String
'
'   lsProcName = "openSource"
'   ''On Error GoTo errProc
'
'   With p_oRepSource
'      .Fields.Append "sField01", adVarChar, 150
'      .Fields.Append "sField02", adVarChar, 100
'      .Fields.Append "sField03", adVarChar, 20
'      .Fields.Append "lField01", adCurrency
'      .Open
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   CloseReport
'   ShowError lsProcName & "( " & " )"
'End Sub
'
'Private Sub setRepProgress(ByVal lnMaxValue As Long, Optional lvRemarks As Variant)
'   With p_oProgress
'      .SecMaxValue = lnMaxValue
'      If Not IsMissing(lvRemarks) Then
'         .PrimaryRemarks = lvRemarks
'      End If
'   End With
'End Sub
'
'Function moveRepProgress(ByVal lsSecRemarks As String, Optional lvPriRemarks As Variant) As Boolean
'   With p_oProgress
'      If Not IsMissing(lvPriRemarks) Then
'         moveRepProgress = .MoveProgress(lsSecRemarks, lvPriRemarks)
'      Else
'         moveRepProgress = .MoveProgress(lsSecRemarks)
'      End If
'   End With
'End Function
'
'Private Sub Class_Initialize()
'   Set p_oProgress = New clsSpeedometer
'   Set p_oRepSource = New Recordset
'   Set p_oRawSource = New Recordset
'End Sub
'
'Private Sub Class_Terminate()
'   Set p_oProgress = Nothing
'   Set p_oRepSource = Nothing
'   Set p_oReport = Nothing
'   Set p_oRawSource = Nothing
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String)
'   With p_oAppDrivr
'      .xLogError Err.Number, Err.Description, "clsSPReceived", lsProcName, Erl
'   End With
'   With Err
'      .Raise .Number, .Source, .Description
'   End With
'End Sub
Private Sub Class_Initialize()

End Sub
