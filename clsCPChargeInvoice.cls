VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPChargeInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'  sheryl 2015-09-03 1:47 pm
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmRecvAgingCriteria
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

   Set p_oStandard = New frmRecvAgingCriteria
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
      .cmbSupportType.Enabled = False
      .txtDateFrom.Enabled = False
      .txtSearch(0).Enabled = False
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
   
      If prcChargeInv = False Then GoTo endProc
   End With

   ProcessReport = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Private Function prcChargeInv() As Boolean
   Dim lsProcName As String
   Dim lsBranch As String
   Dim lsAddCondtn As String
   Dim lsDateThru As String

   lsProcName = "prcChargeInv"
   ''On Error GoTo errProc
   
   With p_oStandard
      setRepProgress 5
      moveRepProgress "Setting Report Retrieval Info..."
      
      Select Case .cmbAging.ListIndex
      Case 0 'all
         lsAddCondtn = ""
      Case 1 'Due
         lsAddCondtn = "AND h.nTermDays >= DateDiff( " & dateParm(.DateThru) & ", a.dTransact)" & _
                        " AND a.cTranStat <> " & xeStateUnknown
      Case 2 'Overdue
         lsAddCondtn = "AND h.nTermDays < DateDiff( " & dateParm(.DateThru) & ", a.dTransact)" & _
                        " AND a.cTranStat <> " & xeStateUnknown
      End Select

      If p_oAppDrivr.IsMainOffice Or p_oAppDrivr.IsWarehouse Then
         lsBranch = IIf(.Branch = "", "", " AND d.sBranchCd = " & strParm(.Branch))
      Else
         lsBranch = " AND d.sBranchCd = " & strParm(.Branch)
      End If

      lsDateThru = .DateThru
      psSQL = "SELECT " & _
                  " d.sBranchNm" & _
                  ", a.sTransNox" & _
                  ", a.dTransact" & _
                  ", h.nTermDays" & _
                  ", DateDiff( " & dateParm(.DateThru) & ", a.dTransact) `nAgeDate`" & _
                  ", a.dStatChng" & _
                  ", e.sSalesInv" & _
                  ", a.sChrgeInv" & _
                  ", f.sCompnyNm" & _
                  ", c.sBarrCode" & _
                  ", c.sDescript" & _
                  ", b.nQuantity" & _
                  ", b.nUnitPrce" & _
                  ", a.cTranStat" & _
                  ", c.cHsSerial" & _
                  ", g.sSerialNo" & _
               " FROM CP_CO_Master a" & _
                  " LEFT JOIN CP_SO_Master e ON a.sReferNox = e.sTransNox" & _
                  " LEFT JOIN Client_Master f ON a.sClientID = f.sClientID" & _
                  " LEFT JOIN Term h ON a.sTermIDxx = h.sTermIDxx" & _
               ", CP_CO_Detail b" & _
                  " LEFT JOIN CP_Inventory_Serial g ON b.sSerialID = g.sSerialID" & _
               ", CP_Inventory c" & _
               ", Branch d"
         
      psSQL = psSQL & _
               " WHERE a.sTransNox = b.sTransNox" & _
               " AND b.sStockIDx = c.sStockIDx" & _
               " AND LEFT(a.sTransNox,4) = d.sBranchCd" & _
               " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
               lsBranch & _
               lsAddCondtn & _
               " AND a.dTransact <= " & dateParm(.DateThru) & _
               " ORDER BY d.sBranchCd, nAgeDate DESC"
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
      .Fields.Append "sField07", adVarChar, 20
      .Fields.Append "nField01", adInteger
      .Fields.Append "nField02", adInteger
      .Fields.Append "lField01", adCurrency
      .Fields.Append "lField02", adCurrency
            
      .Open
   End With
   ' after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource
      Do
         If moveRepProgress(p_oRawSource("sBranchNm") & " » " & p_oRawSource("sTransNox")) = False Then
            CloseReport
               MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
                  GoTo endProc
         End If
         
         
         
      .AddNew
      .Fields("sField01").Value = p_oRawSource("sBranchNm")
      .Fields("sField02").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
      .Fields("sField03").Value = p_oRawSource("sChrgeInv")
      .Fields("sField04").Value = p_oRawSource("sBarrCode")
      .Fields("sField06").Value = Format(p_oRawSource("dStatChng"), "YYYY-MM-DD")
      If p_oRawSource("cHsSerial") = 0 Then
         .Fields("sField05").Value = p_oRawSource("sDescript")
      Else
         .Fields("sField05").Value = IFNull(p_oRawSource("sSerialNo"), "")
      End If
      
      If Not IsNull(p_oRawSource("dStatChng")) And p_oRawSource("cTranStat") = xeStateUnknown Then
         .Fields("sField07").Value = "PAID"
         .Fields("nField02").Value = 0
      ElseIf IsNull(p_oRawSource("dStatChng")) Then
         If p_oRawSource("nAgeDate") > p_oRawSource("nTermDays") Then
            .Fields("sField07").Value = "OVERDUE"
            .Fields("nField02").Value = p_oRawSource("nAgeDate")
         ElseIf p_oRawSource("nTermDays") >= p_oRawSource("nAgeDate") Then
            .Fields("sField07").Value = "DUE DATE"
            .Fields("nField02").Value = p_oRawSource("nAgeDate")
         End If
      ElseIf Not IsNull(p_oRawSource("dStatChng")) And p_oRawSource("cTranStat") = xeStatePosted Then
         .Fields("sField07").Value = "POSTED"
         .Fields("nField02").Value = p_oRawSource("nAgeDate")
      Else
         .Fields("sField07").Value = "VERIFY PAYMENT"
         .Fields("nField02").Value = p_oRawSource("nAgeDate")
      End If
       
      .Fields("nField01").Value = p_oRawSource("nQuantity")
      
      .Fields("lField01").Value = p_oRawSource("nUnitPrce")
      .Fields("lField02").Value = p_oRawSource("nQuantity") * p_oRawSource("nUnitPrce")
      
      p_oRawSource.MoveNext
      Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
      p_oRawSource.Close
   End With
   
   If prcReportObject = False Then GoTo endProc
   
   setRepProgress 1
   moveRepProgress "Loading Report..."
  
   prcChargeInv = True
   
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
      .xLogError Err.Number, Err.Description, "clsCPChargeInvoice", lsProcName, Erl
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
