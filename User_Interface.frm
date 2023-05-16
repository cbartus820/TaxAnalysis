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

    ' Find the column numbers for "Posting Date" and "Situs Main Division"
    Posting_Date_Col = WorksheetFunction.Match("Posting Date", Vertex.Range("1:1"), 0)
    Situs_Main_Div_Col = WorksheetFunction.Match("Situs Main Division", Vertex.Range("1:1"), 0)

    ' Calculate the last row and column in the Vertex sheet
    Vertex_LR = WorksheetFunction.CountA(Vertex.Range("BR:BR"))
    Vertex_LC = WorksheetFunction.CountA(Vertex.Range("1:1"))

    ' Get user-selected Start Date, End Date, and Recon State values
    StartDate = Format(Start_Date.Value, "yyyy-mm-dd")
    EndDate = Format(End_Date.Value, "yyyy-mm-dd")
    ReconState = State_List.Value

    ' Define the range to apply filters on
    Set FilterRange = Vertex.Range(Vertex.Cells(1, 1), Vertex.Cells(Vertex_LR, Vertex_LC))

    ' Apply filters based on user-selected parameters
    If InStr(State_List.Value, "ALL") > 0 Then
        FilterRange.AutoFilter Field:=Posting_Date_Col, Criteria1:=">=" & StartDate, Operator:=xlAnd, Criteria2:="<=" & EndDate
    Else
        FilterRange.AutoFilter Field:=Situs_Main_Div_Col, Criteria1:=WorksheetFunction.Index(Config.Range("A:B"), WorksheetFunction.Match(State_List.Value, Config.Range("B:B"), 0), 1)
        FilterRange.AutoFilter Field:=Posting_Date_Col, Criteria1:=">=" & StartDate, Operator:=xlAnd, Criteria2:="<=" & EndDate
    End If

    ' Copy the filtered data to a new workbook
    FilterRange.SpecialCells(xlCellTypeVisible).Copy
    Set Vertex_Extract = Workbooks.Add
    Set Vertex = Vertex_Extract.ActiveSheet
    Vertex.Range("A1").PasteSpecial xlPasteAll

    ' Check if the CheckBox1 is selected and set the Credit_Report_Inquiry variable accordingly
    If User_Interface.CheckBox1.Value = True Then
        Credit_Report_Inquiry = vbYes
    End If

    ' Call the next subroutine in the sequence
    Call Step_02_Main_Call_Recon_Subroutines

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

Private Sub UserForm_Activate()
Call Populate_Userform
End Sub

Sub Populate_Userform()

    ' Calculate the last row in the "Settings" sheet for destination states (column C) and posting dates (column D)
    Config_LastRow_States = WorksheetFunction.CountA(SourceData.Sheets("Settings").Range("C:C"))
    Config_LastRow_Dates = WorksheetFunction.CountA(SourceData.Sheets("Settings").Range("D:D"))

    ' Populate the State_List ComboBox with US or USA destination states
    With State_List
        For i = 2 To Config_LastRow_States
            If InStr(Config.Range("C" & i).Value, "US") <> 0 Or InStr(Config.Range("C" & i).Value, "USA") <> 0 Then
                If Config.Range("B" & i).Value <> "" Then
                    .AddItem (Config.Range("B" & i).Value)
                End If
            End If
        Next i
    End With

    ' Populate the Start_Date ComboBox with available posting dates from the Config sheet
    With Start_Date
        For i = 2 To Config_LastRow_Dates
            .AddItem (Format(Config.Range("D" & i).Value, "mm/dd/yyyy"))
        Next i
    End With

    ' Populate the End_Date ComboBox with available posting dates from the Config sheet
    With End_Date
        For i = 2 To Config_LastRow_Dates
            .AddItem (Format(Config.Range("D" & i).Value, "mm/dd/yyyy"))
        Next i
    End With

    ' Set default values for Start_Date, End_Date, and State_List ComboBoxes based on the current date
    If Date <= WorksheetFunction.EoMonth(Date, 0) Then
        End_Date.Value = Format(DateSerial(Year(Format(WorksheetFunction.Max(Config.Range("D:D")), "mm/dd/yyyy")), Month(Format(WorksheetFunction.Max(Config.Range("D:D")), "mm/dd/yyyy")) + 1, 0), "mm/dd/yyyy")
        Start_Date.Value = Format(DateSerial(Year(Format(WorksheetFunction.Max(Config.Range("D:D")), "mm/dd/yyyy")), Month(Format(WorksheetFunction.Max(Config.Range("D:D")), "mm/dd/yyyy")), 1), "mm/dd/yyyy")
    Else
        End_Date.Value = WorksheetFunction.EoMonth(Format(WorksheetFunction.Min(Config.Range("C:C")), "mm/dd/yyyy"), 0)
        Start_Date.Value = Format(WorksheetFunction.Min(Config.Range("C:C")), "mm/dd/yyyy")
    End If
    State_List.Value = Config.Range("B2").Value

End Sub
