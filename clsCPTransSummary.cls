VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPTransSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Transaction Summary Report
'
' Copyright 2007 and Beyond
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
'  Jheff [ 09/16/2009 04:00 pm ]
'     Start coding this object...
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPTransSummary"

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmCPRepCriteria
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

   Set p_oStandard = New frmCPRepCriteria
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
      .Caption = "CP Transaction Summary Criteria"
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
      psSQL = "SELECT" & _
                  "  a.sStockIDx" & _
                  ", CONCAT(b.sBarrCode, ' » ', b.sDescript) xDescript" & _
                  ", a.dTransact" & _
                  ", a.sSourceNo" & _
                  ", a.sSourceCd" & _
                  ", c.sSourceNm" & _
                  ", a.nQtyInxxx" & _
                  ", a.nQtyOutxx" & _
                  ", a.nQtyOnHnd" & _
                  ", c.sTableNme" & _
                  ", b.sBarrCode" & _
               " FROM CP_Inventory_Ledger a" & _
                  ", CP_Inventory b" & _
                  ", xxxTransactionSource c" & _
               " WHERE a.sStockIDx = b.sStockIDx" & _
                  " AND a.sSourceCd = c.sSourceID" & _
                  " AND a.sBranchCd = " & strParm(p_oAppDrivr.BranchCode)
      If Not (.SpecifyFrom = "" And .SpecifyThru = "") Then
         psSQL = psSQL & _
               " AND " & IIf(.Presentation = 0, "b.sBarrCode", "b.sDescript") & _
                  " BETWEEN " & strParm(.SpecifyFrom) & " AND " & _
                  IIf(.SpecifyThru = "", strParm(.SpecifyFrom), strParm(.SpecifyThru))
      End If
      psSQL = psSQL & _
                  " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & _
                     " AND " & dateParm(.DateThru) & _
                  " AND a.nQtyInxxx + a.nQtyOutxx > 0" & _
               " ORDER BY " & IIf(.Presentation = 0, "sBarrCode", "sDescript") & _
                  ", a.dTransact" & _
                  ", a.nLedgerNo"
                  
      If prcCPTransSummary() = False Then GoTo endProc
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

Private Function prcCPTransSummary() As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsStockIDx As String
   Dim lnQtyOnHnd As Long

   lsProcName = "prcCPTransSummary"
   ''On Error GoTo errProc

   ' start retrieving
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
Debug.Print psSQL
   With p_oRawSource
      Do
         If moveRepProgress(.Fields("sBarrcode") & " » " & .Fields("xDescript")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If
         
         If LCase(.Fields("sSourceCd")) <> "apad" Then
         Select Case Right(LCase(.Fields("sSourceCd")), 2)
         Case "wa", "wt"   ' pxeCPAcceptWarrantyTransfer, pxeCPWarrantyTransfer
            lsSQL = "SELECT" & _
                        "  a.sTransNox xReferNox" & _
                        ", b.sBranchNm xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        ", Branch b" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo")) & _
                        " AND b.sBranchCd = " & _
                           IIf(Right(LCase(.Fields("sSourceCd")), 2) = "wa", _
                              strParm(Left(.Fields("sSourceNo"), Len(p_oAppDrivr.BranchCode))), _
                              " a.sDestinat")
         Case "dl", "dv"   ' pxeCPAcceptDelivery, pxeCPDelivery
            lsSQL = "SELECT" & _
                        "  a.sTransNox xReferNox" & _
                        ", b.sBranchNm xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        ", Branch b" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo")) & _
                        " AND b.sBranchCd = " & _
                           IIf(Right(LCase(.Fields("sSourceCd")), 2) = "dl", _
                              strParm(Left(.Fields("sSourceNo"), Len(p_oAppDrivr.BranchCode))), _
                              " a.sDestinat")
         Case "al", "av"   ' pxeCPAcceptAutoDelivery, pxeCPAutoDelivery
            lsSQL = "SELECT" & _
                        "  a.sTransNox xReferNox" & _
                        ", b.sBranchNm xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        ", Branch b" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo")) & _
                        " AND b.sBranchCd = " & _
                           IIf(Right(LCase(.Fields("sSourceCd")), 2) = "al", _
                              strParm(Left(.Fields("sSourceNo"), Len(p_oAppDrivr.BranchCode))), _
                              " a.sDestinat")
         Case "jo"   ' pxeCPJobOrder
            lsSQL = "SELECT" & _
                        "  IF(ISNULL(b.sSalesInv), a.sTransNox, b.sSalesInv) xReferNox" & _
                        ", IF(ISNULL(CONCAT(b.nGiftCpnx, ' Gift Coupon'))" & _
                           ", a.sJobDescr, CONCAT(b.nGiftCpnx, ' Gift Coupon')) xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        " LEFT JOIN CP_SO_Master b" & _
                           " ON a.sTransNox = b.sReferNox" & _
                              " AND b.sSourceCd = 'SCJO'" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo"))
         Case "da"  ' pxeCPPurchaseReceiving
            lsSQL = "SELECT" & _
                        "  a.sReferNox xReferNox" & _
                        ", b.sCompnyNm xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        ", Client_Master b" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo")) & _
                        " AND a.sSupplier = b.sClientID"
         Case "pp"   ' pxeCPPurchaseReplacement
            lsSQL = "SELECT" & _
                        "  a.sReferNox xReferNox" & _
                        ", b.sCompnyNm xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        ", Client_Master b" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo")) & _
                        " AND a.sClientID = b.sClientID"
         Case "pr"   ' pxeCPPurchaseReturn
            lsSQL = "SELECT" & _
                        "  a.sTransNox xReferNox" & _
                        ", b.sCompnyNm xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        ", Client_Master b" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo")) & _
                        " AND a.sClientID = b.sClientID"
         Case "wl", "wr", "wp"
            ' pxeCPWholesale, pxeCPWholesaleReturn, pxeCPWholesaleReplacement
            lsSQL = "SELECT" & _
                        "  a.sTransNox xReferNox" & _
                        ", b.sCompnyNm xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        ", Client_Master b" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo")) & _
                        " AND a.sClientID = b.sClientID"
         Case "sl"   ' pxeCPSales
            lsSQL = "SELECT" & _
                        "  a.sSalesInv xReferNox" & _
                        ", a.sRemarksx xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo"))
         Case "sr", "sp"   ' pxeCPSalesReturn, pxeCPSalesReplacement
            lsSQL = "SELECT" & _
                        "  a.sTransNox xReferNox" & _
                        ", a.sRemarksx xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo"))
         Case "ga"   ' pxeCPSalesGiveAway
            lsSQL = "SELECT" & _
                        "  b.sTransNox xReferNox" & _
                        ", b.sRemarksx xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                        ", MC_SO_Master b" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo")) & _
                        " AND a.sTransNox = b.sTransNox"
         Case "ws"   ' pxeCPWarrantyRelease
            lsSQL = "SELECT" & _
                        "  a.sWCRNoxxx xReferNox" & _
                        ", a.sRemarksx xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo"))
         Case "in", "ob", "oc", "ow"
            ' pxeCPPurchase, pxeCPBranchOrder, pxeCPBranchOrderConfirm, pxeCPCustomerOrder
         Case "cm", "dm", "ad"  ' pxeCPCreditMemo, pxeCPDebitMemo
            lsSQL = "SELECT" & _
                        "  a.sDocNmbrx xReferNox" & _
                        ", a.sRemarksx xRemarksx" & _
                     " FROM " & .Fields("sTableNme") & " a" & _
                     " WHERE a.sTransNox = " & strParm(.Fields("sSourceNo"))
         End Select
         Debug.Print lsSQL
         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
         If loRS.EOF Then
            Debug.Print lsSQL
         Else
            If lsStockIDx <> .Fields("sStockIDx") Then
               lnQtyOnHnd = .Fields("nQtyOnHnd") + .Fields("nQtyOutxx") - .Fields("nQtyInxxx")
               lsStockIDx = .Fields("sStockIDx")
            End If
            
            With p_oRepSource
               .AddNew
               .Fields("sField01").Value = p_oRawSource("xDescript")
               .Fields("sField02").Value = p_oRawSource("dTransact")
               .Fields("sField03").Value = IIf(IsNull(loRS("xReferNox")), _
                                             p_oRawSource("sSourceNo"), _
                                             loRS("xReferNox"))
               .Fields("sField04").Value = p_oRawSource("sSourceNm")
               .Fields("sField05").Value = Left(IFNull(loRS("xRemarksx")), 50)
               .Fields("nField01").Value = lnQtyOnHnd
               .Fields("nField02").Value = p_oRawSource("nQtyInxxx")
               .Fields("nField03").Value = p_oRawSource("nQtyOutxx")
               .Fields("nField04").Value = p_oRawSource("nQtyOnHnd")
            End With
         End If
         End If
         
         .MoveNext
      Loop Until .EOF
   End With

   ' after transferring record to source, release raw source...
   p_oRawSource.Close

   If prcReportObject = False Then GoTo endProc

   setRepProgress 1
   moveRepProgress "Loading Report..."

   prcCPTransSummary = True

endProc:
   Set loRS = Nothing
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
      .Fields.Append "sField02", adVarChar, 30
      .Fields.Append "sField03", adVarChar, 50
      .Fields.Append "sField04", adVarChar, 50
      .Fields.Append "sField05", adVarChar, 90
      .Fields.Append "nField01", adInteger, 4
      .Fields.Append "nField02", adInteger, 4
      .Fields.Append "nField03", adInteger, 4
      .Fields.Append "nField04", adInteger, 4
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
