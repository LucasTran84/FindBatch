#NoEnv
#SingleInstance Force

FilePath_D := A_ScriptDir "\The_List.xlsx"

if(!WinExist("The_List - Excel"))
{
        Run, % FilePath_D
        while !WinExist("The_List - Excel")
            Sleep, 500
}

WinActivate, The_List - Excel
oExcel := ComObjActive("Excel.Application")

FilePath_S := oExcel.Worksheets("Setting").Range("B2").Value
oWorkSheetD := oExcel.Worksheets("Main")

oWorkbookS := ComObjGet(FilePath_S)
SLRow := oWorkbookS.Worksheets("RCV").Cells(oWorkbookS.Worksheets("RCV").Rows.Count, 2).End(xlUp:=-4162).Row

LRow := oWorkSheetD.Cells(oWorkSheetD.Rows.Count, 1).End(xlUp:=-4162).Row - 1
Loop, % LRow
{
    vData := oWorkSheetD.Range("A" A_Index + 1).Value
    vRow := oWorkbookS.Worksheets("RCV").Cells(oWorkbookS.Worksheets("RCV").Rows.Count, 5).Find(vData).Row
    if(vRow <> "")
    {
        oWorkSheetD.Range("B" A_Index + 1).Value := oWorkbookS.Worksheets("RCV").Cells(vRow, 4).Value
        oWorkSheetD.Range("C" A_Index + 1).Value := oWorkbookS.Worksheets("RCV").Cells(vRow, 5).Text
        oWorkSheetD.Range("D" A_Index + 1).Value := oWorkbookS.Worksheets("RCV").Cells(vRow, 6).Value
        vDN := oWorkSheetD.Range("E" A_Index + 1).Value := oWorkbookS.Worksheets("RCV").Cells(vRow, 11).Value
        oWorkSheetD.Range("F" A_Index + 1).Value := oWorkbookS.Application.WorksheetFunction.COUNTIF(oWorkbookS.Worksheets("RCV").Range("K4:K" vRow), vDN)
        oWorkSheetD.Range("G" A_Index + 1).Value := oWorkbookS.Application.WorksheetFunction.COUNTIF(oWorkbookS.Worksheets("RCV").Range("E4:E" SLRow), vData)
    }
}

ExitApp

^!p::ExitApp