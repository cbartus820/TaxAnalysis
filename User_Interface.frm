VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} User_Interface 
   Caption         =   "Configure The Program"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8505.001
   OleObjectBlob   =   "User_Interface.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "User_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

Dim Vertex_LR As Long
Dim Vertex_LC As Long
Dim FilterRange As Range
Dim Posting_Date_Col As Long
Dim Situs_Main_Div_Col As Long


Posting_Date_Col = WorksheetFunction.Match("Posting Date", Vertex.Range("1:1"), 0)
Situs_Main_Div_Col = WorksheetFunction.Match("Situs Main Division", Vertex.Range("1:1"), 0)
Vertex_LR = WorksheetFunction.CountA(Vertex.Range("BR:BR"))
Vertex_LC = WorksheetFunction.CountA(Vertex.Range("1:1"))
 
 StartDate = Format(Start_Date.Value, "yyyy-mm-dd")
 EndDate = Format(End_Date.Value, "yyyy-mm-dd")
 
 
 ReconState = State_List.Value
Set FilterRange = Vertex.Range(Vertex.Cells(1, 1), Vertex.Cells(Vertex_LR, Vertex_LC))
 
If InStr(State_List.Value, "ALL") > 0 Then
FilterRange.AutoFilter Field:=Posting_Date_Col, Criteria1:=">=" & StartDate, Operator:=xlAnd, Criteria2:="<=" & EndDate
FilterRange.SpecialCells(xlCellTypeVisible).Copy

Else
FilterRange.AutoFilter Field:=Situs_Main_Div_Col, Criteria1:=WorksheetFunction.Index(Config.Range("A:B"), WorksheetFunction.Match(State_List.Value, Config.Range("B:B"), 0), 1)
FilterRange.AutoFilter Field:=Posting_Date_Col, Criteria1:=">=" & StartDate, Operator:=xlAnd, Criteria2:="<=" & EndDate
FilterRange.SpecialCells(xlCellTypeVisible).Copy

End If

Set Vertex_Extract = Workbooks.Add
Set Vertex = Vertex_Extract.ActiveSheet
Vertex.Range("A1").PasteSpecial xlPasteAll


 If User_Interface.CheckBox1.Value = True Then
 Credit_Report_Inquiry = vbYes
 Else
 End If
 
 'Call Delete_Empty_Columns
 

 
 
 
  Call B_FN_Tax_Recon_FN_Call_Order


End Sub



Private Sub CommandButton2_Click()
ReconState = "All"
For i = 2 To WorksheetFunction.CountA(Config.Range("A:A"))
   State_List.Value = Config.Range("A" & i).Value
    StartDate = Format(Start_Date.Value, "yyyy-mm-dd")
    EndDate = Format(End_Date.Value, "yyyy-mm-dd")
    ReconState = Config.Range("B" & i).Value
    Call Tax_Recon_FN_Call_Order
Next i
 'StartDate = Format(Start_Date.Value, "yyyy-mm-dd")
 'EndDate = Format(End_Date.Value, "yyyy-mm-dd")
  Unload Me

End Sub



Private Sub Quanda_Option_Click()

End Sub

Private Sub Start_Date_Change()

End Sub



Private Sub UserForm_Activate()


With State_List
    For i = 2 To Config_LastRow
    If InStr(Config.Range("C" & i).Value, "US") <> 0 Or InStr(Config.Range("C" & i).Value, "USA") <> 0 Then
    .AddItem (Config.Range("B" & i).Value)
    Else
    End If
    Next i
End With

With Start_Date
    For i = 2 To Config_Date_LastRow
    .AddItem (Format(Config.Range("D" & i).Value, "mm/dd/yyyy"))
    Next i
End With

With End_Date
    For i = 2 To Config_Date_LastRow
    .AddItem (Format(Config.Range("D" & i).Value, "mm/dd/yyyy"))
    Next i
End With


If Date <= WorksheetFunction.EoMonth(Date, 0) Then

End_Date.Value = Format(DateSerial(Year(Format(WorksheetFunction.Max(Config.Range("D:D")), "mm/dd/yyyy")), Month(Format(WorksheetFunction.Max(Config.Range("D:D")), "mm/dd/yyyy")) + 1, 0), "mm/dd/yyyy")
Start_Date.Value = Format(DateSerial(Year(Format(WorksheetFunction.Max(Config.Range("D:D")), "mm/dd/yyyy")), Month(Format(WorksheetFunction.Max(Config.Range("D:D")), "mm/dd/yyyy")), 1), "mm/dd/yyyy")
State_List.Value = Config.Range("B2").Value
Else
End_Date.Value = WorksheetFunction.EoMonth(Format(WorksheetFunction.Min(Config.Range("C:C")), "mm/dd/yyyy"), 0)
Start_Date.Value = Format(WorksheetFunction.Min(Config.Range("C:C")), "mm/dd/yyyy")
State_List.Value = Config.Range("B2").Value
End If

End Sub