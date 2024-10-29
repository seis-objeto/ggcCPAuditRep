VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinanceSummary"
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
   Dim lors As Recordset, loTemp As Recordset
   Dim lsProcName As String
   Dim lnEntryNo As Integer

   lsProcName = "ProcessReport"
   '''On Error GoTo errProc

   With p_oStandard
      .Caption = "Finance Report"
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
   '''On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
      
'      If p_oAppDrivr.IsMainOffice Or p_oAppDrivr.IsWarehouse Then
'         lsBranch = IIf(.Branch = "", "", " AND d.sBranchCd = " & strParm(.Branch))
'      Else
'         lsBranch = " AND d.sBranchCd = " & strParm(.Branch)
'      End If
      
      psSQL = "SELECT" & _
                  " k.sAreaDesc" & _
                  ", e.sBranchNm" & _
                  ", a.dTransact" & _
                  ", j.sCompnyNm" & _
                  ", h.sBrandNme" & _
                  ", i.sModelNme" & _
                  ", d.sBarrcode" & _
                  ", d.sDescript" & _
                  ", g.sSerialNo" & _
                  ", c.nFinAmtxx" & _
                  ", c.nAmtPaidx" & _
                  ", a.sSalesInv" & _
                  ", c.sReferNox" & _
               " FROM CP_SO_Master a" & _
                  " LEFT JOIN Client_Master j ON a.sClientID = j.sClientID" & _
               ", CP_SO_Detail b" & _
                  " LEFT JOIN CP_Inventory_Serial g ON b.sSerialID = g.sSerialID" & _
               ", CP_SO_Finance c" & _
               ", CP_Inventory d" & _
                  " LEFT JOIN CP_Brand h ON d.sBrandIDx = h.sBrandIDx" & _
                  " LEFT JOIN CP_Model i ON d.sModelIDx = i.sModelIDx" & _
               ", Branch e" & _
               ", Branch_Others f" & _
                     " LEFT JOIN Branch_Area k ON f.sAreaCode = k.sAreaCode"
      psSQL = psSQL & _
               " WHERE a.sTransNox = b.sTransNox" & _
               " AND b.sTransNox = c.sTransNox" & _
               " AND b.sStockIDx = d.sStockIDx" & _
               " AND LEFT(a.sTransNox,4) = e.sBranchCd" & _
               " AND e.sBranchCd = f.sBranchCd" & _
               " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
               IIf(.Area = "", "", " AND f.sAreaCode = " & strParm(.Area)) & _
               IIf(.Branch = "", "", " AND e.sBranchCd = " & strParm(.Branch)) & _
               IIf(.Category = "", "", " AND d.sCategID1 = " & strParm(.Category)) & _
               IIf(.SubCategory = "", "", " AND d.sCategID2 = " & strParm(.SubCategory)) & _
               IIf(.Brand = "", "", " AND d.sBrandIDx = " & strParm(.Brand)) & _
               IIf(.Model = "", "", " AND d.sModelIDx = " & strParm(.Brand)) & _
               IIf(.Financer = "", "", " AND c.sClientID = " & strParm(.Financer)) & _
               " ORDER BY f.sAreaCode, e.sBranchNm, a.dTransact"

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
      .Fields.Append "sField01", adVarChar, 250
      .Fields.Append "sField02", adVarChar, 200
      .Fields.Append "sField03", adVarChar, 200
      .Fields.Append "sField04", adVarChar, 200
      .Fields.Append "sField05", adVarChar, 200
      .Fields.Append "sField06", adVarChar, 200
      .Fields.Append "sField07", adVarChar, 200
      .Fields.Append "sField08", adVarChar, 200
      .Fields.Append "sField09", adVarChar, 200
      .Fields.Append "sField10", adVarChar, 200
      .Fields.Append "lField01", adCurrency
      .Fields.Append "lField02", adCurrency
            
      .Open
   End With
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sBranchNm") & " » " & p_oRawSource("sSalesInv")) = False Then
            CloseReport
               MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
                  GoTo endProc
         End If
      
      .AddNew
      If p_oStandard.Presentation = 0 Then
         .Fields("sField01").Value = p_oRawSource("sBranchNm")
'         .Fields("sField02").Value = p_oRawSource("sBranchNm")
'         .Fields("sField03").Value = p_oRawSource("sBrandNme")
         .Fields("lField01").Value = p_oRawSource("nFinAmtxx")
'         .Fields("lField02").Value = p_oRawSource("nAmtPaidx")
      Else
         .Fields("sField01").Value = p_oRawSource("sBranchNm")
         .Fields("sField02").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
         .Fields("sField03").Value = IFNull(p_oRawSource("sSalesInv"), "")
         .Fields("sField04").Value = p_oRawSource("sReferNox")
         .Fields("sField05").Value = IFNull(p_oRawSource("sCompnyNm"), "")
         .Fields("sField06").Value = p_oRawSource("sBarrCode")
         .Fields("sField07").Value = p_oRawSource("sDescript")
         .Fields("sField08").Value = p_oRawSource("sBrandNme")
         .Fields("sField09").Value = p_oRawSource("sModelNme")
         .Fields("sField10").Value = IFNull(p_oRawSource("sSerialNo"), "")
         .Fields("lField01").Value = p_oRawSource("nAmtPaidx")
         .Fields("lField02").Value = p_oRawSource("nFinAmtxx")
      End If
     
      p_oRawSource.MoveNext
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
      .xLogError Err.Number, Err.Description, "clsFinanceSummary", lsProcName, Erl
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


