VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPSales4Audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'  sheryl 2015-08-25 1:59 pm

Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRepMaster As Recordset
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
   ''On Error GoTo errProc

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
   Dim lors As Recordset, loTemp As Recordset
   Dim lsProcName As String
   Dim lnEntryNo As Integer

   lsProcName = "ProcessReport"
   ''On Error GoTo errProc

   With p_oStandard
      .Caption = "CP Sales Summary"
'      .optSerialized(0).Enabled = False
'      .optSerialized(1).Enabled = False
'      .optSerialized(2).Enabled = False
'      .optPresentation(0).Caption = "Summary"
'      .optPresentation(1).Caption = "Detail"
'
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
         Debug.Print psSQL
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
   
      If prcSalesSummary = False Then GoTo endProc
   End With

   ProcessReport = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcSalesSummary() As Boolean
   Dim lsProcName As String
   Dim lsTransNox As String
   Dim loRecData As Recordset
   Dim lnSRPAmtxx As Currency
   Dim lsAsterisk As String
   Dim lnReplAmtx As Currency
   Dim lsLoadQuery As String
   
   lsProcName = "prcSalesSummary"
   ''On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
      
      psSQL = "SELECT" & _
                  " a.sTransNox" & _
                  ", l.sAreaDesc" & _
                  ", e.sBranchNm" & _
                  ", a.sSalesInv" & _
                  ", f.sCompnyNm" & _
                  ", a.dTransact" & _
                  ", c.sBarrCode" & _
                  ", c.sDescript" & _
                  ", i.sBrandNme" & _
                  ", h.sModelNme" & _
                  ", d.sCategrNm" & _
                  ", j.sCategrNm `sSubCateg` " & _
                  ", b.nQuantity * b.nUnitPrce `nSRPAmtxx`" & _
                  ", g.sSerialNo" & _
                  ", a.nReplAmtx" & _
                  ", b.nUnitPrce" & _
                  ", b.nQuantity" & _
                  ", CONCAT(f.sFrstName, ' ', f.sLastName) `sSalesMan`"

      psSQL = psSQL & _
               " FROM CP_SO_Master a" & _
                     " LEFT JOIN Client_Master f" & _
                        " ON a.sClientID = f.sClientID" & _
                     " LEFT JOIN Salesman m ON a.sSalesman = m.sEmployID" & _
               ", CP_SO_Detail b" & _
                     " LEFT JOIN CP_Inventory_Serial g" & _
                        " ON b.sSerialID = g.sSerialID" & _
               ", CP_Inventory c" & _
                     " LEFT JOIN CP_Model h" & _
                        " ON c.sModelIDx = h.sModelIDx" & _
                     " LEFT JOIN CP_Brand i" & _
                        " ON c.sBrandIDx = i.sBrandIDx" & _
                     " LEFT JOIN Category j" & _
                        " ON c.sCategID2 = j.sCategrID" & _
               ", Category d" & _
               ", Branch e" & _
                     " LEFT JOIN Branch_Others k" & _
                        " ON e.sBranchCd = k.sBranchCd" & _
                     " LEFT JOIN Branch_Area l" & _
                        " ON k.sAreaCode = l.sAreaCode"
         psSQL = psSQL & _
               " WHERE a.sTransNox = b.sTransNox" & _
               " AND b.sStockIDx = c.sStockIDx" & _
               " AND c.sCategID1 = d.sCategrID" & _
               " AND LEFT(a.sTransNox,4) = e.sBranchCd" & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
            IIf(.Area = "", "", " AND l.sAreaCode = " & strParm(.Area)) & _
            IIf(.Branch = "", "", " AND a.sTransNox LIKE " & strParm(.Branch + "%")) & _
            IIf(.Brand = "", "", " AND c.sBrandIDx = " & strParm(.Brand)) & _
            IIf(.Model = "", "", " AND c.sModelIDx = " & strParm(.Model)) & _
            IIf(.Category = "", "", " AND c.sCategID1 = " & strParm(.Category)) & _
            IIf(.SubCategory = "", "", " AND c.sCategID2 = " & strParm(.SubCategory)) & _
               " AND a.cTranStat <> " & strParm(xeStateCancelled)
         
         lsLoadQuery = " SELECT " & _
                  " a.sTransNox" & _
                  ", e.sAreaDesc" & _
                  ", c.sBranchNm" & _
                  ", '' sSalesInv" & _
                  ", '' sCompnyNm" & _
                  ", a.dTransact" & _
                  ", b.sBarrCode" & _
                  ", b.sDescript" & _
                  ", 'LOAD WALLET' sBrandNme" & _
                  ", '' sModelNme" & _
                  ", 'LOAD WALLET' sCategrNm" & _
                  ", '' `sSubCateg` " & _
                  ", SUM(a.nAmtPaidx) `nSRPAmtxx`" & _
                  ", '' sSerialNo" & _
                  ", 0 nReplAmtx" & _
                  ", SUM(a.nAmtPaidx) `nUnitPrce`" & _
                  ", '0' nQuantity" & _
                  ", '' `sSalesMan`"

         lsLoadQuery = lsLoadQuery & _
                  " FROM CP_SO_Load_Wallet a" & _
                  ", CP_Load_Matrix b" & _
                  ", Branch c" & _
                  ", Branch_Others d" & _
                        " LEFT JOIN Branch_Area e ON d.sAreaCode = e.sAreaCode" & _
                  " WHERE LEFT(a.sTransNox,4) = c.sBranchCd" & _
                  " AND a.sStockIDx = b.sStockIDx" & _
                  " AND c.sBranchCd = d.sBranchCd" & _
                  IIf(.Area = "", "", " AND d.sAreaCode = " & strParm(.Area)) & _
                  IIf(.Branch = "", "", " AND d.sBranchCd = " & strParm(.Branch)) & _
                  " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & _
                     " AND " & dateParm(.DateThru) & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                  " GROUP BY c.sBranchCd, a.dTransact"
         
         lsLoadQuery = lsLoadQuery & _
                  " UNION SELECT " & _
                  " a.sTransNox" & _
                  ", e.sAreaDesc" & _
                  ", c.sBranchNm" & _
                  ", '' sSalesInv" & _
                  ", '' sCompnyNm" & _
                  ", a.dTransact" & _
                  ", b.sBarrCode" & _
                  ", b.sDescript" & _
                  ", 'LOAD RETAIL' sBrandNme" & _
                  ", '' sModelNme" & _
                  ", 'LOAD RETAIL' sCategrNm" & _
                  ", '' `sSubCateg` " & _
                  ", SUM(a.nTranAmtx) `nSRPAmtxx`" & _
                  ", '' sSerialNo" & _
                  ", 0 nReplAmtx" & _
                  ", SUM(a.nTranAmtx) `nUnitPrce`" & _
                  ", '0' nQuantity" & _
                  ", '' `sSalesMan`"

         lsLoadQuery = lsLoadQuery & _
                  " FROM CP_SO_Eload a" & _
                  ", CP_Load_Matrix b" & _
                  ", Branch c" & _
                  ", Branch_Others d" & _
                        " LEFT JOIN Branch_Area e ON d.sAreaCode = e.sAreaCode" & _
                  " WHERE LEFT(a.sTransNox,4) = c.sBranchCd" & _
                  " AND a.sStockIDx = b.sStockIDx" & _
                  " AND c.sBranchCd = d.sBranchCd" & _
                  IIf(.Area = "", "", " AND d.sAreaCode = " & strParm(.Area)) & _
                  IIf(.Branch = "", "", " AND d.sBranchCd = " & strParm(.Branch)) & _
                  " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & _
                     " AND " & dateParm(.DateThru) & _
                  " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                   " GROUP BY c.sBranchCd, a.dTransact" & _
               " ORDER BY sBranchNm,  dTransact, sSalesInv"

         If .Category = "" Then
            lsLoadQuery = " UNION " & lsLoadQuery
            psSQL = psSQL & " " & lsLoadQuery
         ElseIf .Category = "C001051" Then
            psSQL = lsLoadQuery
         Else
            psSQL = psSQL & " " & " ORDER BY sBranchNm, dTransact, sSalesInv"
         End If

      
'      If p_oAppDrivr.IsMainOffice Or p_oAppDrivr.IsWarehouse Then
'         lsBranch = IIf(.Branch = "", "", " AND e.sBranchCd = " & strParm(.Branch))
'      Else
'         lsBranch = " AND e.sBranchCd = " & strParm(.Branch)
'      End If
'
'      psSQL = "SELECT " & _
'               " a.sTransNox" & _
'               ", a.dTransact" & _
'               ", a.sSalesInv" & _
'               ", c.sBarrCode" & _
'               ", c.sDescript" & _
'               ", b.nQuantity" & _
'               ", b.nUnitPrce" & _
'               ", e.sBranchNm" & _
'               ", d.sCategrNm" & _
'               ", a.nReplAmtx" & _
'               " FROM CP_SO_Master a" & _
'               ", CP_SO_Detail b" & _
'               ", CP_Inventory c" & _
'               ", Category d" & _
'               ", Branch e" & _
'               " WHERE a.sTransNox = b.sTransNox" & _
'               " AND b.sStockIDx = c.sStockIDx" & _
'               " AND c.sCategID1 = d.sCategrID" & _
'               " AND LEFT(a.sTransNox,4) = e.sBranchCd" & _
'               " AND a.cTranStat NOT IN('3','7') " & _
'               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
'               lsBranch & _
'               IIf(.Category = "", "", " AND c.sCategID1 = " & strParm(.Category)) & _
'               " ORDER BY e.sBranchNm,a.dTransact, a.sSalesInv"
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
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      lsTransNox = ""
      lnSRPAmtxx = 0#
      lnReplAmtx = 0#

      Do
         If moveRepProgress(p_oRawSource("sBranchNm") & " » " & p_oRawSource("sTransNox")) = False Then
            CloseReport
               MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
                  GoTo endProc
         End If
         
         If p_oRawSource("nReplAmtx") > 0 And p_oRawSource("nSRPAmtxx") > 0 Then
            If lsTransNox <> p_oRawSource("sTransNox") Then
               lnReplAmtx = p_oRawSource("nReplAmtx")
               lnSRPAmtxx = p_oRawSource("nSRPAmtxx") - lnReplAmtx
               lnReplAmtx = lnReplAmtx - p_oRawSource("nSRPAmtxx")
            ElseIf lsTransNox = p_oRawSource("sTransNox") And lnReplAmtx > 0 Then
               lnSRPAmtxx = p_oRawSource("nSRPAmtxx") - lnReplAmtx
               lnReplAmtx = lnReplAmtx - p_oRawSource("nSRPAmtxx")
            ElseIf lsTransNox = p_oRawSource("sTransNox") And lnReplAmtx <= 0 Then
               lnSRPAmtxx = p_oRawSource("nSRPAmtxx")
            End If
            lsAsterisk = "*"
         Else
            lnSRPAmtxx = p_oRawSource("nSRPAmtxx")
            lnReplAmtx = 0#
            lsAsterisk = ""
         End If

         .AddNew
         .Fields("sField01").Value = p_oRawSource("sBranchNm") & " - " & p_oRawSource("dTransact")
         .Fields("sField02").Value = lsAsterisk & p_oRawSource("sSalesInv")
         .Fields("sField03").Value = p_oRawSource("sBarrCode")
         .Fields("sField04").Value = p_oRawSource("sDescript")
         .Fields("sField05").Value = p_oRawSource("sCategrNm")
         .Fields("sField10").Value = IFNull(p_oRawSource("sSalesman"), "")
         .Fields("nField01").Value = p_oRawSource("nQuantity")
         .Fields("lField01").Value = p_oRawSource("nUnitPrce")
         .Fields("lField02").Value = lnSRPAmtxx
         lsTransNox = p_oRawSource("sTransNox")
      
'      If IFNull(p_oRawSource("nReplAmtx"), 0#) > 0# Then
'         If lsTransNox <> p_oRawSource("sTransNox") Then
'            lnPrice = p_oRawSource("nQuantity") * p_oRawSource("nUnitPrce") - p_oRawSource("nReplAmtx")
'            .Fields("sField02").Value = p_oRawSource("sSalesInv") & "*"
'            .Fields("lField02").Value = IIf(lnPrice < 0, 0#, lnPrice)
'            lsTransNox = p_oRawSource("sTransNox")
'            lnReplAmt = p_oRawSource("nQuantity") * p_oRawSource("nUnitPrce") - p_oRawSource("nReplAmtx")
'         Else
'            lnPrice = p_oRawSource("nQuantity") * p_oRawSource("nUnitPrce") - p_oRawSource("nReplAmtx")
'            .Fields("sField02").Value = p_oRawSource("sSalesInv") & "*"
'            .Fields("lField02").Value = IIf(lnPrice < 0, 0#, lnPrice)
'         End If
'      Else
'         .Fields("sField02").Value = p_oRawSource("sSalesInv")
'         .Fields("lField02").Value = p_oRawSource("nQuantity") * p_oRawSource("nUnitPrce")
'      End If
'
      
      p_oRawSource.MoveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
   End With
   
   If prcReportObject = False Then GoTo endProc
   
   prcSalesSummary = True
   
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
   Set p_oRepMaster = New Recordset
End Sub

Private Sub Class_Terminate()
   Set p_oProgress = Nothing
   Set p_oRepSource = Nothing
   Set p_oReport = Nothing
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "clsCPSales4Audit", lsProcName, Erl
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
   'On Error GoTo errProc

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
   'On Error GoTo errProc
   
   Set p_oRepSource = New Recordset
   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 250
      .Fields.Append "sField02", adVarChar, 100
      .Fields.Append "sField03", adVarChar, 100
      .Fields.Append "sField04", adVarChar, 100
      .Fields.Append "sField05", adVarChar, 50
      .Fields.Append "sField07", adVarChar, 50
      .Fields.Append "sField08", adVarChar, 50
      .Fields.Append "sField09", adVarChar, 50
      .Fields.Append "sField10", adVarChar, 50
      .Fields.Append "nField01", adInteger
      .Fields.Append "lField01", adCurrency
      .Fields.Append "lField02", adCurrency
      .Open
   End With

endProc:
   Exit Sub
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Sub

