VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPAdjustmentSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2004 and beyond
' All Rights Reserved
'
'     Main object for CP Adjustment Report Object
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
'  Sheryl [ 04/27/2010 02:30 pm]
'     continued
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPAdjustmentSummary"

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmBranchCriteria
Private p_oReport As Report

Private p_sFileName As String

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
   psReportNm = psReportNm

   Set p_oStandard = New frmBranchCriteria
   Set p_oStandard.AppDriver = p_oAppDrivr

   InitReport = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & ReportID _
                             & ", " & ReportName & " )"
End Function

Function ProcessReport() As Boolean
   Dim lsProcName As String
   Dim lnEntryNo As Integer

   lsProcName = "ProcessReport"
   ''On Error GoTo errProc

   With p_oStandard
      .Caption = "CP Inventory Adjustment"
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

      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      ' set the retrieval info for active accounts...
      Dim lsStatus As String
      
      Select Case .Presentation
      Case 0
         lsStatus = ""
      Case 1
         lsStatus = " AND a.cTranStat IN('0','1') "
      Case 2
         lsStatus = " AND a.cTranStat IN('2') "
      End Select
      psSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.dTransact" & _
                  ", a.sDocNmbrx" & _
                  ", b.sBarrCode" & _
                  ", a.sRemarksx" & _
                  ", b.sDescript" & _
                  ", a.nQtyInxxx" & _
                  ", a.nQtyOutxx" & _
                  ", Case a.cTranStat" & _
                     " WHEN '0' THEN 'OPEN' " & _
                     " WHEN '1' THEN 'CLOSED' " & _
                     " WHEN '2' THEN 'POSTED' " & _
                  " END cTranStat" & _
               " FROM CP_Inventory_Adjustment a" & _
                  ", CP_Inventory b" & _
               " WHERE a.sStockIDx = b.sStockIDx" & _
                  " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
                  " AND a.dTransact BETWEEN " & dateParm(p_oStandard.DateFrom) & _
                     " AND " & dateParm(p_oStandard.DateThru) & _
               " AND a.cTranStat <> 3 " & _
               lsStatus & _
               " ORDER BY dTransact" & _
                  ", sTransNox"

      If prcDetailPresentation() = False Then GoTo endProc
   End With

   ProcessReport = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcReportObject() As Boolean
   Dim loRepApp As New Application
   Dim lsProcName As String
   Dim lsSubHead As String
   Dim lsAddHead As String

   lsProcName = "prcBranchCollection"
   ''On Error GoTo errProc

   setRepProgress 2
   moveRepProgress "Initializing Report File..."

   Set loRepApp = New Application
   Set p_oReport = loRepApp.OpenReport(p_oAppDrivr.AppPath & _
                        "\Reports\Definitions\" & psReportFl & ".rpt")
   Set loRepApp = Nothing

   moveRepProgress "Assigning Record to Report..."

   With p_oStandard
      lsSubHead = lsSubHead & " As of " & Format(.DateFrom, "MMMM DD, YYYY") & _
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
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcDetailPresentation() As Boolean
   Dim lsProcName As String
   Dim lsTransNox As String

   lsProcName = "prcDetailPresentation"
   ''On Error GoTo errProc

   ' start retrieving
   moveRepProgress "Retrieving Records..."
   
   p_oRawSource.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
   moveRepProgress "Processing Records..."
   If p_oRawSource.EOF Then
      CloseReport
      MsgBox "No Record has Pass the Given Report Critera!!!", vbInformation, "Notice"
      GoTo endProc
   End If

'   ' open the report source recordset...
   openSource

   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRawSource
      Do
         If moveRepProgress(.Fields("sTransNox") & " » " & .Fields("sRemarksx")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If


         With p_oRepSource
            .AddNew

            .Fields("sField01").Value = p_oRawSource("sTransNox")
            .Fields("sField02").Value = Format(p_oRawSource("dTransact"), "mmm-dd-yy")
            .Fields("sField03").Value = p_oRawSource("sDocNmbrx")
            .Fields("sField04").Value = p_oRawSource("sBarrCode")
            .Fields("sField05").Value = p_oRawSource("sRemarksx")
            .Fields("sField06").Value = p_oRawSource("sDescript")
            .Fields("sField07").Value = p_oRawSource("cTranStat")
            .Fields("nField01").Value = p_oRawSource("nQtyInxxx")
            .Fields("nField02").Value = p_oRawSource("nQtyOutxx")
         End With

         .MoveNext
      Loop Until .EOF
   End With

   ' after transferring record to source, release raw source...
   p_oRawSource.Close

   If prcReportObject = False Then GoTo endProc

   setRepProgress 1
   moveRepProgress "Loading Report..."

   prcDetailPresentation = True

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
      .Fields.Append "sField02", adVarChar, 20
      .Fields.Append "sField03", adVarChar, 20
      .Fields.Append "sField04", adVarChar, 20
      .Fields.Append "sField05", adVarChar, 100
      .Fields.Append "sField06", adVarChar, 100
      .Fields.Append "sField07", adVarChar, 50
      .Fields.Append "nField01", adInteger, 4
      .Fields.Append "nField02", adInteger, 4
      .Open
   End With

endProc:
   Exit Sub
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub setRepProgress( _
            ByVal lnMaxValue As Long, _
            Optional lvRemarks As Variant _
            )
   With p_oProgress
      .SecMaxValue = lnMaxValue
      If Not IsMissing(lvRemarks) Then
         .PrimaryRemarks = lvRemarks
      End If
   End With
End Sub

Private Function moveRepProgress( _
                 ByVal lsSecRemarks As String, _
                 Optional lvPriRemarks As Variant _
                 ) As Boolean
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
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub






