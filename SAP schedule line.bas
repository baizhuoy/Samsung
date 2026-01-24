Attribute VB_Name = "modFunctions"
Option Explicit

Public SapGuiAuto As Object, WScript, msgcol
Public objGui As Object
Public objConn As Object
Public objSess As Object
Public objSBar As Object

Public objSheet As Worksheet
Dim W_System

#If VBA7 Then
    Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As LongPtr, ByVal bInheritHandle As LongPtr, ByVal dwProcessId As LongPtr) As LongPtr
    Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As LongPtr) As LongPtr
#Else
    Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
#End If

Public Sub Main()
    Dim W_Ret As Boolean
    
    W_Ret = Attach_Session("SEP100")
    If Not W_Ret Then
        GoTo MyEnd
    End If
    
    Call ProcessExcelData
    
MyEnd:
    Set objSess = Nothing
    Set objGui = Nothing
    Set SapGuiAuto = Nothing
End Sub

Private Function Attach_Session(SID, Optional mysystem As String) As Boolean
    Dim il As Long, it As Long, W_conn, W_Sess
    
    If mysystem = "" Then
        W_System = SID
    Else
        W_System = mysystem
    End If
    
    If W_System = "" Then
        Attach_Session = False
        Exit Function
    End If
    
    If Not objSess Is Nothing Then
        If objSess.Info.SystemName & objSess.Info.Client = W_System Then
            Attach_Session = True
            Exit Function
        End If
    End If
    
    If objGui Is Nothing Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set objGui = SapGuiAuto.GetScriptingEngine
    End If
    
    For il = 0 To objGui.Children.Count - 1
        Set W_conn = objGui.Children(il + 0)
        For it = 0 To W_conn.Children.Count - 1
            Set W_Sess = W_conn.Children(it + 0)
            If W_Sess.Info.SystemName & W_Sess.Info.Client = W_System Then
                Set objConn = objGui.Children(il + 0)
                Set objSess = objConn.Children(it + 0)
                Exit For
            End If
        Next
    Next
    
    If objSess Is Nothing Then
        MsgBox "No active objSess To system " + W_System + ", Or scripting Is Not enabled.", vbCritical + vbOKOnly
        Attach_Session = False
        Exit Function
    End If
    
    If IsObject(WScript) Then
        WScript.ConnectObject objSess, "on"
        WScript.ConnectObject objGui, "on"
    End If
    
    Set objSBar = objSess.findById("wnd[0]/sbar")
    'objSess.findById("wnd[0]").Maximize
    Attach_Session = True

End Function

Public Sub ProcessExcelData()
Dim i, lr As Long
Dim SO, SOLINE, Sch, Status As String
        
    Set objSheet = ActiveWorkbook.ActiveSheet
    lr = objSheet.Range("A1").End(xlDown).Row
    
    
    Columns("A:A").TextToColumns Destination:=Range("A1")
    Columns("B:B").TextToColumns Destination:=Range("B1")
    
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("A2:A" & lr) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("B2:B" & lr) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:D" & lr)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    For i = 2 To lr
        SO = Cells(i, 1).Value
        SOLINE = Cells(i, 2).Value
        Sch = Cells(i, 3).Value
        Call ProcessRowInSAP(i, SO, SOLINE, Sch)
    Next i
    objSess.EndTransaction
End Sub

Sub ProcessRowInSAP(i, SO, SOLINE, Sch)
Dim WINDOW_TITLE, TEST_SUB_STRING, LC_WINDOW_CHECK As String
    
    'Enter Transaction
    objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/nVA02"
    objSess.findById("wnd[0]").sendVKey 0
        
    'Enter SO Number
    objSess.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = SO
    objSess.findById("wnd[0]").sendVKey 0


    On Error Resume Next
    
    objSess.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press 'position
    objSess.findById("wnd[1]/usr/txtRV45A-POSNR").Text = SOLINE
    objSess.findById("wnd[1]").sendVKey 0
    objSess.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_PEIN").press 'go to schedule line
    objSess.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/ctxtVBEP-ETTYP[8,0]").Text = Sch
    objSess.findById("wnd[0]/tbar[1]/btn[29]").press
    'objSess.findById("wnd[0]").sendVKey 0
    
'    If Cells(i, 1) = Cells(i + 1, 1) Then
'        Cells(i, 4) = "Processd"
'
'        Do
'            objSess.findById("wnd[0]/tbar[1]/btn[19]").press 'next item
            
            WINDOW_TITLE = objSess.findById("wnd[0]").Text
            TEST_SUB_STRING = "ATP Change"
            LC_WINDOW_CHECK = InStr(1, WINDOW_TITLE, TEST_SUB_STRING)
            If LC_WINDOW_CHECK <> 0 Then
            On Error Resume Next
                objSess.findById("wnd[0]/tbar[1]/btn[14]").press 'continue shift+f2
                objSess.findById("wnd[0]/tbar[1]/btn[6]").press 'accept policy f6
                objSess.findById("wnd[0]/tbar[1]/btn[6]").press 'accept policy f6
            objSess.findById("wnd[0]/tbar[1]/btn[6]").press
            On Error GoTo 0
            End If
            'If objSess.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").Text = CStr(Cells(i + 1, 2)) Then
        
                'objSess.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/ctxtVBEP-ETTYP[8,0]").Text = Cells(i + 1, 3)
                'objSess.findById("wnd[0]/tbar[1]/btn[29]").press
                'Cells(i + 1, 4) = "Processd"
                'i = i + 1
            'End If
            
'        Loop Until Not Cells(i, 1) = Cells(i + 1, 1)
        
    'End If
    
    
    'Save
    objSess.findById("wnd[0]").sendVKey 11
    
    'Status Message
    Cells(i, 4) = objSBar.Text
End Sub




