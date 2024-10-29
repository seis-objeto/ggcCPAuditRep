VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Inventory Summary Report
'
' Copyright 2009 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Jheff [ 01/15/2009 04:01 pm ]
'     Start coding this object...
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPInventory"

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
      .Caption = "CP Inventory Summary Criteria"
      .optPresentation(0).Caption = "By Location"
      .optPresentation(1).Caption = "By BarCode"
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

      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      ' set the retrieval info for active accounts...
      psSQL = "SELECT" & _
                  "  a.sStockIDx" & _
                  ", a.sBarrCode" & _
                  ", a.sDescript" & _
                  ", e.sModelNme" & _
                  ", b.nQtyOnHnd" & _
                  ", a.nSelPrice" & _
                  ", CONCAT(c.sSectnNme, '-', d.sBinNamex) xLocation" & _
               " FROM CP_Inventory a" & _
                     " LEFT JOIN CP_Model e" & _
                        " ON a.sModelIDx = e.sModelIDx" & _
                  ", CP_Inventory_Master b" & _
                     " LEFT JOIN Section c" & _
                        " ON b.sSectnIDx = c.sSectnIDx" & _
                     " LEFT JOIN Bin d" & _
                        " ON b.sLevelIDx = d.sBinIDxxx" & _
               " WHERE a.sStockIDx = b.sStockIDx" & _
                  " AND b.sBranchCd = " & strParm(.BranchCd) & _
                  " AND b.nQtyOnHnd > 0" & _
                  " AND a.sCategID1 NOT IN('C001001','C001006')" & _
               " ORDER BY" & IIf(.Presentation = 0, _
                  " xLocation, a.sBarrCode", " a.sBarrCode")
      
      If .Presentation = 1 Then
         If prcSummaryPresentation() = False Then GoTo endProc
      Else
         If prcDetailPresentation() = False Then GoTo endProc
      End If
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

Private Function prcSummaryPresentation() As Boolean
   Dim lsProcName As String, lsModelNme As String
   Dim lsConditn As String, lsOrderBy As String
   
   lsProcName = "prcSummaryPresentation"
   ''On Error GoTo errProc
   
   moveRepProgress "Retrieving Records..."
   Debug.Print psSQL
   p_oRawSource.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
   
   moveRepProgress "Processing Records..."
   If p_oRawSource.EOF Then
      CloseReport
      MsgBox "No Record has Pass the Given Report Critera!!!", vbInformation, "Notice"
      GoTo endProc
   End If
   
   ' open the report source recordset...
   openSource
   
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sBarrCode") & " » " & p_oRawSource("sDescript") & " » " & p_oRawSource("sModelNme")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If

         .AddNew
         .Fields("sField02").Value = p_oRawSource("sBarrCode")
         .Fields("sField03").Value = p_oRawSource("sDescript")
         .Fields("sField04").Value = IIf(IsNull(p_oRawSource("sModelNme")), "", p_oRawSource("sModelNme"))
         .Fields("nField01").Value = p_oRawSource("nQtyOnHnd")
         .Fields("lField01").Value = p_oRawSource("nSelPrice")
   
         p_oRawSource.MoveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
   End With
   
   If prcReportObject = False Then GoTo endProc
   
   setRepProgress 1
   moveRepProgress "Loading Report..."
   
   prcSummaryPresentation = True
   
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

   ' open the report source recordset...
   openSource

   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRawSource
      Do
         If moveRepProgress(.Fields("sBarrCode") & " » " & .Fields("sDescript")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If
         
         
         With p_oRepSource
            .AddNew
            .Fields("sField01").Value = IIf(IsNull(p_oRawSource("xLocation")), "", p_oRawSource("xLocation"))
            .Fields("sField02").Value = p_oRawSource("sBarrCode")
            .Fields("sField03").Value = p_oRawSource("sDescript")
            .Fields("sField04").Value = IIf(IsNull(p_oRawSource("sModelNme")), "", p_oRawSource("sModelNme"))
            .Fields("nField01").Value = p_oRawSource("nQtyOnHnd")
            .Fields("lField01").Value = p_oRawSource("nSelPrice")
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
      .Fields.Append "sField02", adVarChar, 50
      .Fields.Append "sField03", adVarChar, 90
      .Fields.Append "sField04", adVarChar, 200
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
