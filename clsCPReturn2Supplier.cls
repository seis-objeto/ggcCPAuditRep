VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPReturn2Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'  sheryl 2017-01-10 4:01 pm
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmCPSalesCriteria
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

   Set p_oStandard = New frmCPSalesCriteria
   Set p_oStandard.AppDriver = p_oAppDrivr

   InitReport = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & ReportID _
                       & ", " & ReportName & " )"
End Function

Function ProcessReport() As Boolean
   Dim loRS As Recordset, loTemp As Recordset
   Dim lsProcName As String
   Dim lnEntryNo As Integer

   lsProcName = "ProcessReport"
   '''On Error GoTo errProc

   With p_oStandard
      .Caption = "CP Return Report"
      .txtSearch(0).Enabled = False
      .txtSearch(2).Enabled = False
      .txtSearch(3).Enabled = False
      .txtSearch(4).Enabled = False
      .txtSearch(5).Enabled = False
      .txtSearch(5).Enabled = False
      .txtSearch(6).Enabled = False
      If p_oAppDrivr.IsMainOffice = False Then
         .txtSearch(1).Enabled = False
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
   
      If .Presentation = 0 Then
         If prcSummary = False Then GoTo endProc
      Else
         If prcDetail = False Then GoTo endProc
      End If
   End With

   ProcessReport = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcSummary() As Boolean
   Dim lsProcName As String

   lsProcName = "prcSummary"
   '''On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
      
      psSQL = "SELECT " & _
                  " e.sAreaDesc" & _
                  ", c.sBranchNm" & _
                  ", a.dTransact" & _
                  ", a.dTransact dReferDte" & _
                  ", a.sReferNox" & _
                  ", b.sCompnyNm" & _
                  ", a.nTranTotl, a.cTranStat" & _
                  ", 'CPPP' `sSourceCd`" & _
                  ", a.sTransNox" & _
                  ", a.cTranStat" & _
                  ", a.sRemarksx" & _
               " FROM CP_PO_Replacement_Master a" & _
               ", Client_Master b" & _
               ", Branch c" & _
               ", Branch_Others d" & _
               ", Branch_Area e" & _
               " WHERE a.sClientID = b.sClientID" & _
               " AND LEFT(a.sTransNox,4) = c.sBranchCd" & _
               " AND c.sBranchCd = d.sBranchCd" & _
               " AND d.sAreaCode = e.sAreaCode" & _
               " AND a.cTranStat <> " & xeStateCancelled & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
            IIf(.Branch = "", "", " AND c.sBranchCd = " & strParm(.Branch)) & _
               " AND a.dTransact BETWEEN " & dateParm(p_oStandard.DateFrom) & " AND " & dateParm(p_oStandard.DateThru)
               
      psSQL = psSQL & " UNION SELECT " & _
                  " e.sAreaDesc" & _
                  ", c.sBranchNm" & _
                  ", a.dTransact" & _
                  ", a.dTransact dReferDte" & _
                  ", a.sReferNox" & _
                  ", b.sCompnyNm, a.nTranTotl" & _
                  ", a.cTranStat" & _
                  ", 'CPPR' `sSourceCd` , a.sTransNox" & _
                   ", a.cTranStat" & _
                   ", a.sRemarksx" & _
               " FROM CP_PO_Return_Master a" & _
               ", Client_Master b" & _
               ", Branch c" & _
               ", Branch_Others d" & _
               ", Branch_Area e" & _
               " WHERE a.sClientID = b.sClientID" & _
               " AND LEFT(a.sTransNox,4) = c.sBranchCd" & _
               " AND c.sBranchCd = d.sBranchCd" & _
               " AND d.sAreaCode = e.sAreaCode" & _
               " AND a.cTranStat <> " & xeStateCancelled & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
            IIf(.Branch = "", "", " AND c.sBranchCd = " & strParm(.Branch)) & _
               " AND a.dTransact BETWEEN " & dateParm(p_oStandard.DateFrom) & " AND " & dateParm(p_oStandard.DateThru) & _
               " ORDER BY sBranchNm, dTransact"
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
      
    openSource
    
    setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."
      
   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sAreaDesc") & " » " & p_oRawSource("sBranchNm")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If
   
         With p_oRepSource
             .AddNew
            .Fields("sField01").Value = p_oRawSource("sAreaDesc")
            .Fields("sField02").Value = p_oRawSource("sBranchNm")
            .Fields("sField03").Value = IFNull(p_oRawSource("sCompnyNm"), "")
            .Fields("sField04").Value = p_oRawSource("sTransNox")
            .Fields("sField05").Value = p_oRawSource("sSourceCd")
            .Fields("sField06").Value = IFNull(p_oRawSource("sReferNox"), "")
            .Fields("sField07").Value = Format(p_oRawSource("dReferDte"), "YYYY-MM-DD")
            .Fields("sField08").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
            .Fields("lField01").Value = p_oRawSource("nTranTotl")
            Select Case p_oRawSource("cTranStat")
            Case 0
               .Fields("sField09").Value = "OPEN"
            Case 1
               .Fields("sField09").Value = "CLOSED"
            Case 2
               .Fields("sField09").Value = "POSTED"
            Case Else
               .Fields("sField09").Value = "UNKNOWN"
            End Select
            .Fields("sField12").Value = IFNull(p_oRawSource("sRemarksx"), "")
         End With
         
            p_oRawSource.MoveNext
         Loop Until p_oRawSource.EOF
         
   End With
   
         If prcReportObject = False Then GoTo endProc
   
         setRepProgress 1
         moveRepProgress "Loading Report..."
  
         prcSummary = True
   
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
      .xLogError Err.Number, Err.Description, "clsCPReturn2Supplier", lsProcName, Erl
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
      
      p_oReport.Sections("PH").ReportObjects("txtCompany").SetText .ClientName
      p_oReport.Sections("PH").ReportObjects("txtAddress").SetText .Address & _
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


Private Sub openSource()
   Dim lsProcName As String

   lsProcName = "openSource"
   ''On Error GoTo errProc

   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 90
      .Fields.Append "sField02", adVarChar, 100
      .Fields.Append "sField03", adVarChar, 90
      .Fields.Append "sField04", adVarChar, 100
      .Fields.Append "sField05", adVarChar, 100
      .Fields.Append "sField06", adVarChar, 100
      .Fields.Append "sField07", adVarChar, 100
      .Fields.Append "sField08", adVarChar, 100
      .Fields.Append "sField09", adVarChar, 100
      .Fields.Append "sField10", adVarChar, 100
      .Fields.Append "sField11", adVarChar, 100
      .Fields.Append "sField12", adVarChar, 254
      .Fields.Append "nField01", adInteger, 4
      .Fields.Append "lField01", adDouble
      .Open
   End With

endProc:
   Exit Sub
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Sub

Private Function prcDetail() As Boolean
   Dim lsProcName As String

   lsProcName = "prcDetail"
   '''On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
      
      psSQL = "SELECT " & _
                  " e.sAreaDesc" & _
                  ", c.sBranchNm" & _
                  ", a.dTransact" & _
                  ", a.dTransact dReferDte" & _
                  ", a.sReferNox" & _
                  ", b.sCompnyNm" & _
                  ", a.nTranTotl, a.cTranStat" & _
                  ", 'CPPP' `sSourceCd`" & _
                  ", a.sTransNox" & _
                  ", a.cTranStat" & _
                  ", g.sSerialNo" & _
                  ", h.sBarrCode" & _
                  ", i.sModelNme" & _
                  ", a.sRemarksx" & _
               " FROM CP_PO_Replacement_Master a" & _
               ", Client_Master b" & _
               ", Branch c" & _
               ", Branch_Others d" & _
               ", Branch_Area e" & _
               ", CP_PO_Replacement_Detail f" & _
                  " LEFT JOIN CP_Inventory_Serial g ON f.sSerialID = g.sSerialID" & _
                  " LEFT JOIN CP_Inventory h ON g.sStockIDx = h.sStockIDx" & _
                  " LEFT JOIN CP_Model i ON h.sModelIDx = i.sModelIDx"

      psSQL = psSQL & _
               " WHERE a.sClientID = b.sClientID" & _
               " AND LEFT(a.sTransNox,4) = c.sBranchCd" & _
               " AND c.sBranchCd = d.sBranchCd" & _
               " AND d.sAreaCode = e.sAreaCode" & _
               " AND a.sTransNox = f.sTransNox" & _
               " AND a.cTranStat <> " & xeStateCancelled & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
            IIf(.Branch = "", "", " AND c.sBranchCd = " & strParm(.Branch)) & _
               " AND a.dTransact BETWEEN " & dateParm(p_oStandard.DateFrom) & " AND " & dateParm(p_oStandard.DateThru)
               
      psSQL = psSQL & " UNION SELECT " & _
                  " e.sAreaDesc" & _
                  ", c.sBranchNm" & _
                  ", a.dTransact" & _
                  ", a.dTransact dReferDte" & _
                  ", a.sReferNox" & _
                  ", b.sCompnyNm, a.nTranTotl" & _
                  ", a.cTranStat" & _
                  ", 'CPPR' `sSourceCd` , a.sTransNox" & _
                   ", a.cTranStat" & _
                   ", g.sSerialNo" & _
                  ", h.sBarrCode" & _
                  ", i.sModelNme" & _
                   ", a.sRemarksx" & _
               " FROM CP_PO_Return_Master a" & _
               ", Client_Master b" & _
               ", Branch c" & _
               ", Branch_Others d" & _
               ", Branch_Area e" & _
               ", CP_PO_Return_Detail f" & _
                  " LEFT JOIN CP_Inventory_Serial g ON f.sSerialID = g.sSerialID" & _
                  " LEFT JOIN CP_Inventory h ON g.sStockIDx = h.sStockIDx" & _
                  " LEFT JOIN CP_Model i ON h.sModelIDx = i.sModelIDx"
                  
      psSQL = psSQL & _
               " WHERE a.sClientID = b.sClientID" & _
               " AND LEFT(a.sTransNox,4) = c.sBranchCd" & _
               " AND c.sBranchCd = d.sBranchCd" & _
               " AND d.sAreaCode = e.sAreaCode" & _
               " AND a.sTransNox = f.sTransNox" & _
               " AND a.cTranStat <> " & xeStateCancelled & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
            IIf(.Branch = "", "", " AND c.sBranchCd = " & strParm(.Branch)) & _
               " AND a.dTransact BETWEEN " & dateParm(p_oStandard.DateFrom) & " AND " & dateParm(p_oStandard.DateThru) & _
               " ORDER BY sBranchNm, dTransact"
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
      
    openSource
    
    setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."
      
   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sAreaDesc") & " » " & p_oRawSource("sBranchNm")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If
   
         With p_oRepSource
             .AddNew
            .Fields("sField01").Value = p_oRawSource("sAreaDesc")
            .Fields("sField02").Value = p_oRawSource("sBranchNm")
            .Fields("sField03").Value = IFNull(p_oRawSource("sCompnyNm"), "")
            .Fields("sField04").Value = p_oRawSource("sTransNox")
            .Fields("sField05").Value = p_oRawSource("sSourceCd")
            .Fields("sField06").Value = IFNull(p_oRawSource("sReferNox"), "")
            .Fields("sField07").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
            .Fields("sField08").Value = IFNull(p_oRawSource("sModelNme"), "")
            .Fields("lField01").Value = p_oRawSource("nTranTotl")
            Select Case p_oRawSource("cTranStat")
            Case 0
               .Fields("sField09").Value = "OPEN"
            Case 1
               .Fields("sField09").Value = "CLOSED"
            Case 2
               .Fields("sField09").Value = "POSTED"
            Case Else
               .Fields("sField09").Value = "UNKNOWN"
            End Select
            .Fields("sField10").Value = IFNull(p_oRawSource("sBarrCode"), "")
            .Fields("sField11").Value = IFNull(p_oRawSource("sSerialNo"), "")
            .Fields("sField12").Value = IFNull(p_oRawSource("sRemarksx"), "")
         End With
         
            p_oRawSource.MoveNext
         Loop Until p_oRawSource.EOF
         
   End With
   
         If prcReportObject = False Then GoTo endProc
   
         setRepProgress 1
         moveRepProgress "Loading Report..."
  
         prcDetail = True
   
endProc:
      Exit Function
errProc:
      CloseReport
      ShowError lsProcName & "( " & " )"
End Function

