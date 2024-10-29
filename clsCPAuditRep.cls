VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPAuditRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Main object for Cellphone Integrated System - Cellphone Audit Report Object
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
'  Jheff [ 09/16/2009 04:45 pm ]
'     Start coding this object...
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPAuditRep"

Private p_oAppDrivr As clsAppDriver
Private p_oRepBrowse As frmReports
Private p_oRepSource As Object

Private p_sRepName As String

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Source() As Report
   Set Source = p_oRepSource.Source
End Property

Public Function ShowReport() As Boolean
   Dim lsProcName As String
   Dim lsRepName As String
   Dim lsSQL As String
   Dim lnRow As Long
   
   lsProcName = "ShowReports"
   ''On Error GoTo errProc
   
   Set p_oRepBrowse = New frmReports
   Set p_oRepBrowse.AppDriver = p_oAppDrivr
   
   With p_oRepBrowse
      .Show 1
      If .Preview = False Then GoTo endProc
      
      Set p_oRepSource = CreateObject(Trim(.ReportLibrary) & "." & Trim(.ReportClass))
      Set p_oRepSource.AppDriver = p_oAppDrivr
      p_oRepSource.InitReport .ReportID, .ReportName
      
      If p_oRepSource.ProcessReport = False Then GoTo endProc
      
      lsRepName = getReportCode
      If .SaveReport Then
'         p_oRepSource.Source.DisplayProgressDialog = False
'         p_oRepSource.Source.ExportOptions.DestinationType = 1
'         p_oRepSource.Source.ExportOptions.FormatType = 1
'         p_oRepSource.Source.ExportOptions.DiskFileName = p_oAppDrivr.AppPath & "\Reports\" & lsRepName
'         p_oRepSource.Source.Export False
      End If
      
      If .LogReport Then
         With p_oAppDrivr
            lsSQL = "INSERT INTO xxxReportsLog (" & _
                        "  sReportID" & _
                        ", dGenerate" & _
                        ", sUserIDxx" & _
                        ", sRepFName" & _
                     " ) VALUES (" & _
                        strParm(p_oRepBrowse.ReportID) & _
                        ", " & dateParm(.ServerDate) & _
                        ", " & strParm(.UserID) & _
                        ", " & strParm(lsRepName) & " )"

            .Connection.Execute lsSQL, lnRow, adCmdText
            If lnRow = 0 Then
               MsgBox "Unable to Register Report Generation!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End With
      End If
   End With
   
   ShowReport = True
   
endProc:
   Unload p_oRepBrowse
   Set p_oRepBrowse = Nothing

   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Sub CloseReport()
   p_oRepSource.CloseReport
End Sub

Private Function getReportCode() As String
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "getReportCode"
   ''On Error GoTo errProc
   
   ' first get the computer id
   getReportCode = Format(p_oAppDrivr.ServerDate, "yy") & _
            p_oAppDrivr.BranchCode & _
            p_oAppDrivr.ComputerID
   
   ' then get the latest Report name based on the year and computer id
   lsSQL = "SELECT sRepFName " & _
            " FROM xxxReportsLog" & _
            " WHERE sRepFName LIKE " & strParm(getReportCode & "%") & _
            " ORDER BY sRepFName DESC" & _
            " LIMIT 1"
   
   Debug.Print lsSQL
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

   If loRS.EOF Then
      getReportCode = getReportCode & Format(1, "0000")
   Else
      getReportCode = getReportCode & Format(CInt(Right(loRS(0), 4)) + 1, "0000")
   End If
   
endProc:

   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

Private Sub Class_Terminate()
   Set p_oRepSource = Nothing
End Sub
