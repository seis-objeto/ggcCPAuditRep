VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPInOutSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2004 and beyond
' All Rights Reserved
'
'     Main object for CP Transfer-Received Report Object
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Jheff [ 09/26/2009 10:50 am]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmCPDelTran
Private p_oReport As Report

Private psReportID As String
Private psReportNm As String
Private psReportFl As String
Private psReportHd As String
Private psSQL As String

Dim psQtyInxxx As String
Dim psQtyOutxx As String
Dim psQtyOnHnd As String

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
   Dim lsOldProc As String
   
   lsOldProc = "ProcessReport"
   ''On Error GoTo errProc
   
   psReportID = ReportID
   psReportNm = ReportName
   
   Set p_oStandard = New frmCPDelTran
   Set p_oStandard.AppDriver = p_oAppDrivr
   
   InitReport = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & ReportID _
                       & ", " & ReportName & " )"
End Function

Function ProcessReport() As Boolean
   Dim lors As Recordset, loTemp As Recordset
   Dim lsOldProc As String
   Dim lnEntryNo As Integer
   
   lsOldProc = "ProcessReport"
   ''On Error GoTo errProc
   
   With p_oStandard
      .Branch = p_oAppDrivr.BranchCode
      .txtBranch.Text = p_oAppDrivr.BranchName
      .Caption = "Transfer/Received Summary Criteria"
      .Show 1
   
      If .Cancelled Then
         MsgBox "Report Generation was Cancelled"
         GoTo endProc
      End If
      lnEntryNo = .ReceivedType + 1

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
      
      If .ReceivedType = 0 Then   ' Transfer Presentation
         If prcCPSummary = False Then GoTo endProc
      Else
'         If prcLoadSummary = False Then GoTo endProc
      End If
   End With
   
   ProcessReport = True
   
endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsOldProc & "( " & " )"
End Function

Private Function prcReportObject() As Boolean
   Dim loRepApp As New Application
   Dim lsOldProc As String
   Dim lsSubHead As String
   Dim lsAddHead As String
   
   lsOldProc = "prcBranchCollection"
   ''On Error GoTo errProc
   
   setRepProgress 2
   moveRepProgress "Initializing Report File..."
   
   Set loRepApp = New Application
   Set p_oReport = loRepApp.OpenReport(p_oAppDrivr.AppPath & _
                        "\Reports\Definitions\" & psReportFl & ".rpt")
   Set loRepApp = Nothing
   
   moveRepProgress "Assigning Record to Report..."
   With p_oStandard
      lsAddHead = IIf(.Presentation = 0, " - Transfer", " - Received")
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
                  
      p_oReport.Sections("PH").ReportObjects("txtHeading1").SetText psReportHd & lsAddHead
      p_oReport.Sections("PH").ReportObjects("txtHeading2").SetText lsSubHead
               
      p_oReport.Sections("PF").ReportObjects("txtRptUser").SetText .UserName
   End With
   
   prcReportObject = True
   
endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub openSource()
   Dim lsOldProc As String
   
   lsOldProc = "openSource"
   ''On Error GoTo errProc
   
   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 20
      .Fields.Append "sField02", adVarChar, 20
      .Fields.Append "sField03", adVarChar, 50
      .Fields.Append "sField04", adVarChar, 200
      .Fields.Append "sField05", adVarChar, 20
      .Fields.Append "nField01", adInteger, 10
      .Open
   End With

endProc:
   Exit Sub
errProc:
   CloseReport
   ShowError lsOldProc & "( " & " )"
End Sub

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

Private Function prcCPSummary() As Boolean
   Dim lsOldProc As String
   Dim lsSQLCon As String
   Dim lsBranchCd As String
   Dim lsDescript As String
   Dim lsSourcexx As String
   
   lsOldProc = "prcCPSummary"
   ''On Error GoTo errProc
      
   p_oProgress.CloseProgress
   With p_oStandard
      lsBranchCd = IIf(.Branch = "", p_oAppDrivr.BranchCode, .Branch)
      
      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      psSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.dTransact" & _
                  ", g.sBranchNm" & _
                  ", c.sBarrCode" & _
                  ", c.sDescript" & _
                  ", d.sBrandNme" & _
                  ", e.sModelNme" & _
                  ", f.sColorNme" & _
                  ", COUNT(b.nQuantity) nQuantity" & _
               " FROM CP_Transfer_Master a" & _
                  ", CP_Transfer_Detail b" & _
                  ", CP_Inventory c" & _
                     " LEFT JOIN CP_Brand d" & _
                        " ON c.sBrandIDx = d.sBrandIDx" & _
                     " LEFT JOIN CP_Model e" & _
                        " ON c.sModelIDx = e.sModelIDx" & _
                     " LEFT JOIN Color f" & _
                        " ON c.sColorIDx = f.sColorIDx" & _
                  ", Branch g"
      psSQL = psSQL & _
               " WHERE a.sTransNox = b.sTransNox" & _
                  " AND b.sStockIDx = c.sStockIDx" & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                  " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & _
                     " AND " & dateParm(.DateThru & " 23:59:59") & _
                  IIf(.Presentation = 0, " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(lsBranchCd) & _
                     " AND a.sDestinat = g.sBranchCd", " AND a.sDestinat = " & strParm(lsBranchCd) & _
                     " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = g.sBranchCd") & _
               " GROUP BY b.sStockIDx" & _
               " ORDER BY" & _
                  "  a.dTransact" & _
                  ", g.sBranchNm" & _
                  ", a.sTransNox" & _
                  ", c.sBarrCode"
   End With
   
   moveRepProgress "Retrieving Records..."
   p_oRawSource.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
   Debug.Print psSQL

   moveRepProgress "Processing Records..."
   If p_oRawSource.EOF Then GoTo endProc

   ' open the report source recordset...
   openSource
   With p_oRepSource
      ' after retrieving record process report
      setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."
      
      lsSourcexx = ""
      Do While p_oRawSource.EOF = False
         If moveRepProgress(p_oRawSource("sTransNox") & " » " & p_oRawSource("sBranchNm")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If
            
         .AddNew
         .Fields("sField01") = Format(p_oRawSource("dTransact"), "MMM-DD-YYYY")
         If lsSourcexx <> p_oRawSource("sTransNox") & "»" & p_oRawSource("sBranchNm") Then
            .Fields("sField02") = p_oRawSource("sTransNox")
            .Fields("sField03") = p_oRawSource("sBranchNm")
         End If
         lsDescript = IFNull(p_oRawSource("sBrandNme"), "") & " " & _
                        IFNull(p_oRawSource("sModelNme"), "") & " " & _
                        IFNull(p_oRawSource("sColorNme"), "")
         .Fields("sField04") = IIf(lsDescript = "", p_oRawSource("sDescript"), lsDescript)
         .Fields("sField05") = p_oRawSource("sBarrCode")
         .Fields("nField01") = p_oRawSource("nQuantity")
         
         lsSourcexx = p_oRawSource("sTransNox") & "»" & p_oRawSource("sBranchNm")
         p_oRawSource.MoveNext
      Loop
'      .Sort = "sField01,sField03,sField02,sField05"
   End With
   
   ' after transferring record to source, release raw source...
   p_oRawSource.Close

   If prcReportObject = False Then GoTo endProc

   setRepProgress 1
   moveRepProgress "Loading Report..."

   prcCPSummary = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "clsCPInOutSum", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
