VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPSalesDetailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2004 and beyond
' All Rights Reserved
'
'     Main object for Branch Sales Summary Report Object
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
'  Sheryl [ 01/26/2011 10:25 am]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmCPSalesDet
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
   'On Error Goto errProc

   psReportID = ReportID
   psReportNm = ReportName

   Set p_oStandard = New frmCPSalesDet
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
   'On Error Goto errProc

   With p_oStandard
      .Caption = "Warranty Card Sales Summary"
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
      
      Select Case .Presentation
      Case 0   ' Sony
         If prcSalesSony = False Then GoTo endProc
      Case 1   ' Etouch
         If prcSalesEtouch = False Then GoTo endProc
      Case 2   ' Samsung
         If prcSalesSamsung = False Then GoTo endProc
      End Select
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

   lsProcName = "prcReportObject"
   'On Error Goto errProc

   setRepProgress 2
   moveRepProgress "Initializing Report File..."

   Set loRepApp = New Application
   Set p_oReport = loRepApp.OpenReport(p_oAppDrivr.AppPath & _
                        "\Reports\Definitions\" & psReportFl & ".rpt")
   Set loRepApp = Nothing

   Select Case p_oStandard.Presentation
   Case 0
      lsAddHead = " (Summarized)"
   Case 1
      lsAddHead = " (Detailed)"
   End Select

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

Private Function prcSalesSony() As Boolean
   Dim lsProcName As String, lsModelNme As String

   lsProcName = "prcSalesSony"
   'On Error Goto errProc

   With p_oStandard
      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      psSQL = "Select" _
                  & " f.sBranchNm" _
                  & ", a.dTransact" _
                  & ", CONCAT(g.sFrstName,' ' ,g.sMiddName,'' ,g.sLastName) xCust" _
                  & ", e.sModelNme" _
                  & ", c.sSerialNo" _
               & " From CP_SO_Master a" _
                    & ",CP_SO_Detail b" _
                    & " LEFT JOIN CP_Inventory_Serial c" _
                     & " ON b.sSerialId = c.sSerialId" _
                    & " LEFT JOIN CP_Inventory d" _
                     & " ON c.sStockIdx = d.sStockIdx" _
                    & " LEFT JOIN CP_Model e" _
                    & " ON d.sModelIdx = e.sModelIdx" _
                    & ", Branch f" _
                    & ", Client_Master g" _
               & " Where a.sTransNox = b.sTransNox" _
                  & " AND f.sBranchCd = LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ")" _
                  & " AND a.sClientId = g.sClientId" _
                  & " AND d.cHsSerial = 1" _
                  & " And a.dTransact Between " & dateParm(.DateFrom) & " And " & dateParm(.DateThru & " 23:59:59") _
                  & " And f.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) _
                  & " And a.cTranStat <> " & strParm(xeStateCancelled)

       psSQL = psSQL _
            & " Order By f.sBranchNm" _
               & ", a.dTransact"
   End With

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
         If moveRepProgress(.Fields("sBranchNm") & " » " & .Fields("sSerialNo")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If


         With p_oRepSource
            .AddNew

         .Fields("sField01").Value = ("Guanzon")
         .Fields("sField02").Value = p_oRawSource("sBranchNm")
         .Fields("sField03").Value = p_oRawSource("dTransact")
         .Fields("sField04").Value = p_oRawSource("xCust")
         .Fields("sField05").Value = IIf(IsNull(p_oRawSource("sModelNme")), "-", p_oRawSource("sModelNme"))
         .Fields("sField06").Value = p_oRawSource("sSerialNo")
         End With

         .MoveNext
      Loop Until .EOF
   End With

   If prcReportObject = False Then GoTo endProc

   setRepProgress 1
   moveRepProgress "Loading Report..."

   prcSalesSony = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcSalesEtouch() As Boolean
 Dim lsProcName As String, lsModelNme As String

   lsProcName = "prcSalesEtouch"
   'On Error Goto errProc

   With p_oStandard
      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      psSQL = "Select" _
                  & " f.sBranchNm" _
                  & ", a.dTransact" _
                  & ", CONCAT(g.sFrstName,' ' ,g.sMiddName,'' ,g.sLastName) xCust" _
                  & ", e.sModelNme" _
                  & ", c.sSerialNo" _
                  & ", g.cGenderCd" _
                  & ", g.dBirthDte" _
                  & ", CONCAT(h.sFrstName, ' ', h.sLastName)xSalesman" _
                  & ", a.sSalesInv" _
               & " From CP_SO_Master a" _
                    & " LEFT JOIN Salesman h" _
                    & " ON a.sSalesman = h.sEmployID" _
                  & ",CP_SO_Detail b" _
                    & " LEFT JOIN CP_Inventory_Serial c" _
                    & " ON b.sSerialId = c.sSerialId" _
                    & " LEFT JOIN CP_Inventory d" _
                    & " ON c.sStockIdx = d.sStockIdx" _
                    & " LEFT JOIN CP_Model e" _
                    & " ON d.sModelIdx = e.sModelIdx" _
                  & ", Branch f" _
                  & ", Client_Master g" _

         psSQL = psSQL _
               & " Where a.sTransNox = b.sTransNox" _
                  & " AND f.sBranchCd = LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ")" _
                  & " AND a.sClientId = g.sClientId" _
                  & " AND d.cHsSerial = 1" _
                  & " And a.dTransact Between " & dateParm(.DateFrom) & " And " & dateParm(.DateThru & " 23:59:59") _
                  & " And f.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) _
                  & " And a.cTranStat <> " & strParm(xeStateCancelled) _
            & " Order By f.sBranchNm" _
               & ", a.dTransact"
   End With

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
         If moveRepProgress(.Fields("sBranchNm") & " » " & .Fields("sSerialNo")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If


         With p_oRepSource
            .AddNew
         .Fields("sField02").Value = p_oRawSource("cGenderCd")
         .Fields("sField03").Value = p_oRawSource("dTransact")
         .Fields("sField04").Value = p_oRawSource("xCust")
         .Fields("sField05").Value = IIf(IsNull(p_oRawSource("sModelNme")), "-", p_oRawSource("sModelNme"))
         .Fields("sField06").Value = p_oRawSource("sSerialNo")
         .Fields("sField07").Value = p_oRawSource("dBirthDte")
         .Fields("sField08").Value = IIf(IsNull(p_oRawSource("xSalesman")), " ", p_oRawSource("xSalesman"))
         .Fields("sField09").Value = p_oRawSource("sBranchNm")
         
         End With

         .MoveNext
      Loop Until .EOF
   End With

   If prcReportObject = False Then GoTo endProc

   setRepProgress 1
   moveRepProgress "Loading Report..."

   prcSalesEtouch = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcSalesSamsung() As Boolean
 Dim lsProcName As String, lsModelNme As String

   lsProcName = "prcSalesSamsung"
   'On Error Goto errProc

   With p_oStandard
      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      psSQL = "Select" _
                  & " f.sBranchNm" _
                  & ", a.dTransact" _
                  & ", CONCAT(g.sFrstName,' ' ,g.sMiddName,'' ,g.sLastName) xCust" _
                  & ", e.sModelNme" _
                  & ", c.sSerialNo" _
                  & ", g.cGenderCd" _
                  & ", g.dBirthDte" _
                  & ", CONCAT(h.sFrstName, ' ', h.sLastName)xSalesman" _
                  & ", a.sSalesInv" _
                  & ", g.sEmailAdd" _
               & " From CP_SO_Master a" _
                    & " LEFT JOIN Salesman h" _
                    & " ON a.sSalesman = h.sEmployID" _
                  & ",CP_SO_Detail b" _
                    & " LEFT JOIN CP_Inventory_Serial c" _
                    & " ON b.sSerialId = c.sSerialId" _
                    & " LEFT JOIN CP_Inventory d" _
                    & " ON c.sStockIdx = d.sStockIdx" _
                    & " LEFT JOIN CP_Model e" _
                    & " ON d.sModelIdx = e.sModelIdx" _
                  & ", Branch f" _
                  & ", Client_Master g" _

         psSQL = psSQL _
               & " Where a.sTransNox = b.sTransNox" _
                  & " AND f.sBranchCd = LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ")" _
                  & " AND a.sClientId = g.sClientId" _
                  & " AND d.cHsSerial = 1" _
                  & " And a.dTransact Between " & dateParm(.DateFrom) & " And " & dateParm(.DateThru & " 23:59:59") _
                  & " And f.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) _
                  & " And a.cTranStat <> " & strParm(xeStateCancelled) _
            & " Order By f.sBranchNm" _
               & ", a.dTransact"
   End With

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
         If moveRepProgress(.Fields("sBranchNm") & " » " & .Fields("sSerialNo")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If


         With p_oRepSource
            .AddNew
         .Fields("sField01").Value = p_oRawSource("sEmailAdd")
         .Fields("sField02").Value = p_oRawSource("xCust")
         .Fields("sField03").Value = p_oRawSource("cGenderCd")
         .Fields("sField04").Value = p_oRawSource("xCust")
         .Fields("sField05").Value = IIf(IsNull(p_oRawSource("sModelNme")), "-", p_oRawSource("sModelNme"))
         .Fields("sField06").Value = p_oRawSource("sSerialNo")
         .Fields("sField07").Value = p_oRawSource("dBirthDte")
         .Fields("sField08").Value = IIf(IsNull(p_oRawSource("xSalesman")), " ", p_oRawSource("xSalesman"))
         .Fields("sField09").Value = p_oRawSource("sBranchNm")
         .Fields("sField10").Value = p_oRawSource("sSalesInv")
         
         
         End With

         .MoveNext
      Loop Until .EOF
   End With

   If prcReportObject = False Then GoTo endProc

   setRepProgress 1
   moveRepProgress "Loading Report..."

   prcSalesSamsung = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcSalesLG() As Boolean
 Dim lsProcName As String, lsModelNme As String

   lsProcName = "prcSalesLG"
   'On Error Goto errProc

   With p_oStandard
      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      psSQL = "Select" _
                  & " f.sBranchNm" _
                  & ", a.dTransact" _
                  & ", CONCAT(g.sFrstName,' ' ,g.sMiddName,'' ,g.sLastName) xCust" _
                  & ", e.sModelNme" _
                  & ", c.sSerialNo" _
                  & ", g.cGenderCd" _
                  & ", g.dBirthDte" _
                  & ", CONCAT(h.sFrstName, ' ', h.sLastName)xSalesman" _
                  & ", a.sSalesInv" _
                  & ", g.sEmailAdd" _
               & " From CP_SO_Master a" _
                    & " LEFT JOIN Salesman h" _
                    & " ON a.sSalesman = h.sEmployID" _
                  & ",CP_SO_Detail b" _
                    & " LEFT JOIN CP_Inventory_Serial c" _
                    & " ON b.sSerialId = c.sSerialId" _
                    & " LEFT JOIN CP_Inventory d" _
                    & " ON c.sStockIdx = d.sStockIdx" _
                    & " LEFT JOIN CP_Model e" _
                    & " ON d.sModelIdx = e.sModelIdx" _
                  & ", Branch f" _
                  & ", Client_Master g" _

         psSQL = psSQL _
               & " Where a.sTransNox = b.sTransNox" _
                  & " AND f.sBranchCd = LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ")" _
                  & " AND a.sClientId = g.sClientId" _
                  & " AND d.cHsSerial = 1" _
                  & " And a.dTransact Between " & dateParm(.DateFrom) & " And " & dateParm(.DateThru & " 23:59:59") _
                  & " And f.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) _
                  & " And a.cTranStat <> " & strParm(xeStateCancelled) _
            & " Order By f.sBranchNm" _
               & ", a.dTransact"
   End With

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
         If moveRepProgress(.Fields("sBranchNm") & " » " & .Fields("sSerialNo")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If


         With p_oRepSource
            .AddNew
         .Fields("sField01").Value = p_oRawSource("sEmailAdd")
         .Fields("sField02").Value = p_oRawSource("xCust")
         .Fields("sField03").Value = p_oRawSource("cGenderCd")
         .Fields("sField04").Value = p_oRawSource("xCust")
         .Fields("sField05").Value = IIf(IsNull(p_oRawSource("sModelNme")), "-", p_oRawSource("sModelNme"))
         .Fields("sField06").Value = p_oRawSource("sSerialNo")
         .Fields("sField07").Value = p_oRawSource("dBirthDte")
         .Fields("sField08").Value = IIf(IsNull(p_oRawSource("xSalesman")), " ", p_oRawSource("xSalesman"))
         .Fields("sField09").Value = p_oRawSource("sBranchNm")
         .Fields("sField10").Value = p_oRawSource("sSalesInv")
         
         
         End With

         .MoveNext
      Loop Until .EOF
   End With

   If prcReportObject = False Then GoTo endProc

   setRepProgress 1
   moveRepProgress "Loading Report..."

   prcSalesLG = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Sub openSource()
   Dim lsProcName As String

   lsProcName = "openSource"
   'On Error Goto errProc

   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 50
      .Fields.Append "sField02", adVarChar, 25
      .Fields.Append "sField03", adVarChar, 100
      .Fields.Append "sField04", adVarChar, 60
      .Fields.Append "sField05", adVarChar, 25
      .Fields.Append "sField06", adVarChar, 25
      .Fields.Append "sField07", adVarChar, 25
      .Fields.Append "sField08", adVarChar, 25
      .Fields.Append "sField09", adVarChar, 60
      .Open
   End With

endProc:
   Exit Sub
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
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

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "clsSalesDetailed", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub



