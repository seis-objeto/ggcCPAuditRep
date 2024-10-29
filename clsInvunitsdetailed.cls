VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvunitsdetailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'  Sheryl [ 07/09/2013 04:55 pm]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRepSource As Recordset
Private p_oRawSource As Recordset
Private p_oProgress As clsSpeedometer
Private p_oStandard As frmCPTranSumCriteria
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

   Set p_oStandard = New frmCPTranSumCriteria
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
               
         If prcSalesSony = False Then GoTo endProc

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
      
      p_oReport.Sections("PH").ReportObjects("txtHeading1").SetText psReportHd
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

Private Function prcInventoryUnits() As Boolean
 Dim lsProcName As String, lsModelNme As String

   lsProcName = "prcInventoryUnits"
   'On Error Goto errProc

   With p_oStandard
      setRepProgress 3
      moveRepProgress "Setting Report Retrieval Info..."

      psSQL = " SELECT sSerialNo" & _
                  " ,IFNULL(xLastRecv,sBranchCD)xLastRecv" & _
                  " ,xBranch" & _
                  " ,sBranchCD" & _
                  " ,sBrandNme" & _
                  " ,sModelNme" & _
                  " ,sDescript" & _
                  " ,sColorNme" & _
                  " ,cLocation" & _
               " FROM " & _
                  " ( SELECT  a.sSerialID" & _
                        " ,a.sSerialNo" & _
                        " ,c.sSourceCd" & _
                        " ,b.dTransact dDteInxxx" & _
                        " ,c.dTransact dDteOutxx" & _
                        " ,a.sBranchCd" & _
                        " ,d.sBarrCode" & _
                        " ,d.sDescript" & _
                        " ,e.sBrandNme" & _
                        " ,f.sModelNme" & _
                        " ,g.sColorNme" & _
                        " ,a.cLocation" & _
                        " ,d.nSelPrice" & _
                        " ,h.sDescript xLastRecv" & _
                        " ,h.cAutomate xLastAuto"
         
         psSQL = psSQL & _
                        " ,i.sDescript xBranch" & _
                        " ,i.cAutomate" & _
                  " FROM CP_Inventory_Serial a" & _
                        " LEFT JOIN CP_Inventory_Serial_Ledger b" & _
                           " ON a.sSerialID = b.sSerialID" & _
                           " AND b.sSourceCd IN ('CPDA','CPDl','CPDm','CPSR','CPSp')" & _
                        " LEFT JOIN Branch h" & _
                           " ON b.sBranchCd = h.sBranchCd" & _
                        " LEFT JOIN CP_Inventory_Serial_Ledger c" & _
                           " ON a.sSerialID = c.sSerialID" & _
                           " AND c.sSourceCd IN ('CPSl','CPPR','CPCm')" & _
                           " AND c.dTransact >= " & dateParm(.DateThru) & _
                  ", CP_Inventory d" & _
                        " LEFT JOIN CP_Brand e" & _
                           " ON d.sBrandIDx = e.sBrandIDx" & _
                        " LEFT JOIN CP_Model f" & _
                           " ON d.sModelIDx = f.sModelIDx" & _
                        " LEFT JOIN Color g" & _
                           " ON d.sColorIDx = g.sColorIDx" & _
                  ", Branch i"
                  
         psSQL = psSQL & _
                  " WHERE a.sStockIDx = d.sStockIDx" & _
                  " AND d.sCategID1 = 'C001001'" & _
                  " AND a.cUnitClas = '0'" & _
                  " AND a.cUnitType = '1'" & _
                  " AND ( a.cSoldStat = '0' OR ( a.cSoldStat = '1' AND c.dTransact >= " & dateParm(.DateThru) & ")" & ")" & _
                  " AND a.sBranchCd = i.sBranchCd" & _
                  " HAVING IF(dDteInxxx IS NULL,  1 = 1  , dDteInxxx <= " & dateParm(.DateThru) & ")" & _
                  " AND IF(dDteOutxx IS NULL, (cLocation = '0' OR ( cLocation = '1' AND i.cAutomate = '1') OR ( cLocation = '4' AND i.cAutomate = '1')) " & _
                  " , dDteOutxx > dDteInxxx AND dDteOutxx > " & dateParm(.DateThru) & ")" & _
                  " ORDER BY dDteInxxx DESC) xSource " & _
            " WHERE LEFT(xBranch,1) IN('1','2','3','W')" & _
            " GROUP BY sSerialNo" & _
            " ORDER BY  sBranchCD, xBranch, sModelNme"
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

   prcInventoryUnits = True

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





