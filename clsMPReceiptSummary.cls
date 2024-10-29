VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMPReceiptSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'She 2019-10-30 2:52 pm
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmSuReTrackCriteria
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
   On Error GoTo errProc

   psReportID = ReportID
   psReportNm = ReportName

   Set p_oStandard = New frmSuReTrackCriteria
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
   On Error GoTo errProc

   With p_oStandard
      .Caption = "Receipt Series"
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
   
      If prcReceiptSummary = False Then GoTo endProc
   
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
   On Error GoTo errProc

   setRepProgress 2
   moveRepProgress "Initializing Report File..."

   Set loRepApp = New Application
   Set p_oReport = loRepApp.OpenReport(p_oAppDrivr.AppPath & _
                        "\Reports\Definitions\" & psReportFl & ".rpt")
   Set loRepApp = Nothing

   With p_oStandard

      moveRepProgress "Assigning Record to Report..."
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

   End With

   prcReportObject = True

endProc:
   Exit Function
errProc:
   CloseReport
   ShowError lsProcName & "( " & " )"
End Function

Function prcReceiptSummary() As Boolean
   Dim lsOldProc As String
   Dim lsCondition As String
   
   lsOldProc = "prcReceiptSummary"
   On Error GoTo errProc

   With p_oStandard
      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      'SI
      psSQL = "SELECT b.sBranchNm" & _
                  ", a.dTransact" & _
                  ", a.sSalesInv sReferNox" & _
                  ", a.nAmtPaidx nTranTotl" & _
                  ", 'SI' `sSourceCd`" & _
                  ", a.cTranStat" & _
               " FROM CP_SO_Master a" & _
               ", Branch b" & _
               " WHERE LEFT(a.sTransNox,4) = b.sBranchCd" & _
               " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode + "%") & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru)
      'TDR
      psSQL = psSQL & _
               " UNION ALL SELECT b.sBranchNm" & _
                  ", a.dTransact" & _
                  ", a.sChrgeInv sReferNox" & _
                  ", a.nTranTotl nTranTotl" & _
                  ", 'TDR' `sSourceCd`" & _
                  ", a.cTranStat" & _
               " FROM CP_CO_Master a" & _
               ", Branch b" & _
               " WHERE LEFT(a.sTransNox,4) = b.sBranchCd" & _
               " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode + "%") & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru)
      'OR NORTHPOINT
      psSQL = psSQL & _
               " UNION ALL SELECT b.sBranchNm" & _
                  ", a.dTransact" & _
                  ", a.sReferNox" & _
                  ", a.nAmountxx nTranTotl" & _
                  ", 'OR NORTHPOINT' `sSourceCd`" & _
                  ", a.cPostedxx cTranStat" & _
               " FROM LR_Payment_Master a" & _
               ", Branch b" & _
               " WHERE LEFT(a.sTransNox,4) = b.sBranchCd" & _
               " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode + "%") & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru)
      'PR NORTHPOINT
      psSQL = psSQL & _
               " UNION ALL SELECT b.sBranchNm" & _
                  ", a.dTransact" & _
                  ", a.sReferNox" & _
                  ", a.nAmountxx nTranTotl" & _
                  ", 'PR NORTHPOINT' `sSourceCd`" & _
                  ", a.cPostedxx cTranStat" & _
               " FROM LR_Payment_Master_PR a" & _
               ", Branch b" & _
               " WHERE LEFT(a.sTransNox,4) = b.sBranchCd" & _
               " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode + "%") & _
               " AND a.dTransact BETWEEN " & dateParm(.DateFrom) & " AND " & dateParm(.DateThru) & _
               " ORDER BY `sSourceCd`, sReferNox "
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

   'open the report source recordset...
   openSource
   
   'after retrieving record process report
   setRepProgress p_oRawSource.RecordCount, "Assigning Record to the Report..."

   With p_oRepSource

         Do
            If moveRepProgress(p_oRawSource("sBranchNm") & " » " & p_oRawSource("sReferNox")) = False Then
              CloseReport
               MsgBox "Report Generation was Cancelled!!!", vbInformation, "Notice"
               GoTo endProc
            End If
         
            .AddNew
'            .Fields("sField01").Value = p_oRawSource("sAreaDesc")
            .Fields("sField02").Value = p_oRawSource("sBranchNm")
            .Fields("sField04").Value = Format(p_oRawSource("dTransact"), "YYYY-MM-DD")
            .Fields("sField05").Value = p_oRawSource("sSourceCd")
            .Fields("sField06").Value = p_oRawSource("sReferNox")
            .Fields("lField01").Value = p_oRawSource("nTranTotl")
            Select Case p_oRawSource("cTranstat")
            Case 0
               .Fields("sField07").Value = "OPEN"
            Case 1
               .Fields("sField07").Value = "CLOSED"
            Case 2
               .Fields("sField07").Value = "POSTED"
            Case 3
               .Fields("sField07").Value = "CANCELLED"
            End Select
     p_oRawSource.MoveNext
         Loop Until p_oRawSource.EOF
   
      ' after transferring record to source, release raw source...
         p_oRawSource.Close
      End With
   
      If prcReportObject = False Then GoTo endProc
   
      setRepProgress 1
     moveRepProgress "Loading Report..."
   
      prcReceiptSummary = True
   
endProc:
     Exit Function
errProc:
      CloseReport
      ShowError lsOldProc & "( " & " )"
End Function

Private Sub openSource()
   Dim lsOldProc As String

   lsOldProc = "openSource"
   On Error GoTo errProc

   With p_oRepSource
      .Fields.Append "sField01", adVarChar, 50
      .Fields.Append "sField02", adVarChar, 50
      .Fields.Append "sField03", adVarChar, 100
      .Fields.Append "sField04", adVarChar, 100
      .Fields.Append "sField05", adVarChar, 50
      .Fields.Append "sField06", adVarChar, 50
      .Fields.Append "sField07", adVarChar, 200
      .Fields.Append "sField08", adVarChar, 50
      .Fields.Append "nField01", adInteger, 20
      .Fields.Append "lField01", adCurrency, 20
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

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "clsMPReceiptSummary", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub


