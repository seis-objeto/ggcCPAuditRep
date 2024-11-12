VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBolttechRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'maynard 09-16-2024
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmDateBranchCriteria
Private p_oReport As Report
Private p_oJSONParser As clsJSON

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

   Set p_oStandard = New frmDateBranchCriteria
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
      .Caption = "Bolttech Report Criteria"
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
   

   If prcSummary = False Then GoTo endProc

   
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
   ''On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
   
      psSQL = "SELECT" & _
                    "  a.sTransNox" & _
                    ", a.sSourceNo" & _
                    ", b.dTransact" & _
                    ", a.sPayloadx" & _
                    ", c.sFileName" & _
                    ", c.dCreatedx" & _
                    ", c.dDateSent" & _
                    ", IFNULL(b.sPartnerN, '') sPartnerN" & _
                " FROM CP_SO_Insurance a" & _
                    ", CP_SO_Master b" & _
                    ", Bolttech c" & _
                " WHERE a.sSourceNo = b.sTransNox" & _
                    " AND a.sBatchNox = c.sBatchNox" & _
                    " AND a.cTranStat = '1'" & _
                    " AND c.cTranStat = '1'" & _
                    " AND a.sSourceNo IN ('C00523103285', 'C00623001737', 'C02723005963', 'C02723005964', 'C02723005965', 'C02923002342', 'C03523004351', 'C03523004352', 'C03823001921', 'C07123001867', 'C09323001003', 'C09824000912', 'C10723001613')" & _
                " ORDER BY b.dTransact"
                
        'psSQL = AddCondition(psSQL, " b.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru))
        
        If .BranchCd <> "" Then
        psSQL = AddCondition(psSQL, " a.sSourceNo LIKE " & strParm(.BranchCd & "%"))
        End If
      Debug.Print psSQL
   End With
   
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
'   openSource
   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 100
      .Fields.Append "sField02", adVarChar, 100
      .Fields.Append "sField03", adVarChar, 100
      .Fields.Append "sField04", adVarChar, 100
      .Fields.Append "sField05", adVarChar, 100
      .Fields.Append "sField06", adVarChar, 100
      .Fields.Append "sField07", adVarChar, 100
      .Fields.Append "sField08", adVarChar, 250
      .Fields.Append "sField09", adVarChar, 250
      .Fields.Append "sField10", adVarChar, 250
      .Fields.Append "sField11", adVarChar, 250
      .Fields.Append "lField01", adCurrency
      .Fields.Append "lField02", adCurrency
      .Fields.Append "lField03", adCurrency
      
      .Open
   End With
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sTransNox") & " » " & p_oRawSource("sSourceNo")) = False Then
            CloseReport
            MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
            GoTo endProc
         End If
         
         Dim parsedObject As Object
         Set parsedObject = p_oJSONParser.parse(p_oRawSource("sPayloadx"))
         If p_oJSONParser.GetParserErrors() = "" Then
             Debug.Print "JSON Object: " & p_oJSONParser.toString(parsedObject)
         Else
             Debug.Print "Parsing Errors: " & p_oJSONParser.GetParserErrors()
         End If
         
         Dim lors As Recordset
         Set lors = New Recordset

         psSQL = " SELECT " & _
                     "  a.sTransNox" & _
                     ", b.sStockIDx" & _
                     ", c.sCategID1" & _
                     ", c.sCategID2" & _
                     ", f.sCategrNm" & _
                     ", e.sTransNox xTransNox" & _
                     ", b.nUnitPrce" & _
                     ", a.nTranTotl" & _
                     ", a.nCashAmtx" & _
                     ", a.nAmtPaidx" & _
                     ", a.sApplicNo" & _
                     ", a.sApproved" & _
                     ", IFNULL(g.nAmountxx,0) xCardAmnt" & _
                     ", IFNULL(h.nAmtPaidx,0) xEPaymAmt" & _
                     ", c.sDescript" & _
                     ", c.nSelPrice" & _
                  " FROM CP_SO_Master a" & _
                     " LEFT JOIN CP_SO_Insurance e ON a.sTransNox = e.sSourceNo AND e.cTranStat <> '3'" & _
                     " LEFT JOIN MP_SO_Credit_Card g ON a.sTransNox = g.sReferNox AND g.sSourceCd = 'CPSl'" & _
                     " LEFT JOIN Other_Payment_Received h ON a.sTransNox = h.sReferNox AND h.sSourceCd = 'CPSl'" & _
                     ", CP_SO_Detail b" & _
                     " LEFT JOIN CP_Inventory c ON b.sStockIDx = c.sStockIDx" & _
                     " LEFT JOIN Category d ON c.sCategID1 = d.sCategrID" & _
                     " LEFT JOIN Category f ON c.sCategID2 = f.sCategrID"
                     
         psSQL = psSQL & _
                  " WHERE a.sTransNox = b.sTransNox" & _
                     " AND a.cTranStat NOT IN ('3', '7')" & _
                  " HAVING c.sCategID1 = 'C001001'"
             
         psSQL = AddCondition(psSQL, "a.sTransNox = " & strParm(IIf(p_oRawSource("sPartnerN") = "", p_oRawSource("sSourceNo"), p_oRawSource("sPartnerN"))))
             
         Set lors = p_oAppDrivr.Connection.Execute(psSQL, , adCmdText)
             
         Dim lsMOP As String
         Dim pbCashPayment As Boolean
         Dim pbFinancePayment As Boolean
         Dim pbCreditPayment As Boolean
         Dim pbEPayment As Boolean
         
         pbCashPayment = False
         pbFinancePayment = False
         pbCreditPayment = False
         pbEPayment = False
      
         If lors.EOF Then
            GoTo moveNext
         End If
         
         If CDbl(lors("nTranTotl")) = CDbl(lors("nAmtPaidx")) And CDbl(lors("nTranTotl")) = CDbl(lors("nCashAmtx")) Then
            lsMOP = "Cash"
            pbCashPayment = True
         ElseIf CDbl(lors("nTranTotl")) = CDbl(lors("nAmtPaidx")) And lors("sApplicNo") <> "" Then
            lsMOP = "Financing"
            pbFinancePayment = True
         ElseIf CDbl(lors("nTranTotl")) <= CDbl(lors("xCardAmnt")) Then
            lsMOP = "Credit Card"
            pbCreditPayment = True
         ElseIf CDbl(lors("nTranTotl")) <= CDbl(lors("xEPaymAmt")) Then
            lsMOP = "E-Payment"
            pbCreditPayment = True
         End If
      
'         If pbFinancePayment Then
'
'            If CDbl(loRS("xEPaymAmt")) > 0 Then
'                lsMOP = lsMOP + ", E-Payment"
'            End If
'
'            If CDbl(loRS("nCashAmtx")) > 0 Then
'                lsMOP = lsMOP + ", w/Down Payment"
'            End If
'
'            If CDbl(loRS("xCardAmnt")) > 0 Then
'                lsMOP = lsMOP + ", Credit Card"
'            End If
'            End If
'         End If
'
'         If pbCashPayment Then
'            If CDbl(loRS("xEPaymAmt")) > 0 Then
'                lsMOP = lsMOP + ", E-Payment"
'            End If
'
'            If CDbl(loRS("xCardAmnt")) > 0 Then
'                lsMOP = lsMOP + ", Credit Card"
'            End If
'            End If
'         End If
      
         If pbCreditPayment Then
            If CDbl(lors("xEPaymAmt")) > 0 Then
                lsMOP = lsMOP + ", E-Payment"
            End If
            
            If CDbl(lors("nCashAmtx")) > 0 Then
                lsMOP = lsMOP + ", Cash"
            End If
         
         ElseIf pbEPayment Then
            If CDbl(lors("xCardAmnt")) > 0 Then
                lsMOP = lsMOP + ", Credit Card"
            End If
            
            If CDbl(lors("nCashAmtx")) > 0 Then
                lsMOP = lsMOP + ", Cash"
            End If
         End If
              
         .AddNew
         .Fields("sField01").Value = parsedObject.Item("CLIENT_TRANS_NO")
         .Fields("sField02").Value = parsedObject.Item("CONTRACT_SOLD_DATE")
         .Fields("sField03").Value = lors("sDescript") 'getProductName(p_oRawSource("sSourceNo"))  'parsedObject.Item("PRODUCT_NAME")
         .Fields("lField01").Value = Format(lors("nUnitPrce"), "#,##.00")
         .Fields("lField03").Value = Format(lors("nSelPrice"), "#,##.00")
         .Fields("sField04").Value = parsedObject.Item("CUST_NAME")
         .Fields("sField05").Value = parsedObject.Item("STORE_NAME")
         .Fields("sField06").Value = parsedObject.Item("SALES_REP_NAME")
         .Fields("sField07").Value = parsedObject.Item("DEVICE_MAKE")
         .Fields("sField08").Value = parsedObject.Item("DEVICE_MODEL")
         .Fields("sField09").Value = parsedObject.Item("COLOR")
         .Fields("lField02").Value = Format(parsedObject.Item("DEVICE_RRP"), "#,##.00")
         .Fields("sField10").Value = lsMOP
         .Fields("sField11").Value = IIf(parsedObject.Item("IMEI") = "", parsedObject.Item("SERIALNO"), parsedObject.Item("IMEI"))
         
moveNext:
         lors.Close
         p_oRawSource.moveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
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
   Set p_oJSONParser = New clsJSON
   Set p_oRepSource = New Recordset
   Set p_oRawSource = New Recordset
End Sub

Private Sub Class_Terminate()
   Set p_oProgress = Nothing
   Set p_oJSONParser = Nothing
   Set p_oRepSource = Nothing
   Set p_oReport = Nothing
   Set p_oRawSource = Nothing
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "clsMCReservationRep", lsProcName, Erl
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
      p_oReport.Sections("RH").ReportObjects("txtCompany").SetText .ClientName
      p_oReport.Sections("RH").ReportObjects("txtAddress").SetText .Address & _
                  ", " & .TownCity & ", " & .Province & " " & .ZippCode
      p_oReport.Sections("RH").ReportObjects("txtHeading1").SetText psReportHd & " " & lsAddHead
      p_oReport.Sections("RH").ReportObjects("txtHeading2").SetText lsSubHead

      p_oReport.Sections("PF").ReportObjects("txtRptUser").SetText .UserName
   End With

   prcReportObject = True
   
endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Function getProductName(lsTransNox As String) As String
    Dim lsSQL As String
    Dim loData As Recordset

    lsSQL = "SELECT a.sTransNox, b.sBarrCode, b.sDescript" & _
               " FROM CP_SO_Detail a" & _
               ", CP_Inventory b" & _
               " WHERE a.sStockIDx = b.sStockIdx" & _
               " AND b.sCategID1 = 'C001052'" & _
               " AND a.sTransNox = " & strParm(lsTransNox)
   Debug.Print lsSQL
    Set loData = New Recordset
    loData.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

    If Not loData.EOF Then
         getProductName = loData("sDescript")
    Else
        getProductName = ""
    End If
End Function

