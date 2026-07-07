Option Explicit

Sub MergeSurvey()

    Dim Folder As String
    Dim FileList As Variant
    Dim FileName As String
    Dim i As Long

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet

    Dim LastRow As Long
    Dim LastCol As Long
    Dim NextRow As Long

    Dim FirstFile As Boolean

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    '========================================
    '  — Ì» «·„·›« 
    '========================================
    FileList = Array( _
        "S1.xlsx", _
        "S2.xlsx", _
        "S3.xlsx", _
        "S4.xlsx" _
    )

    '========================================
    ' „”«— „Ã·œ Modules
    '========================================
    Folder = ThisWorkbook.Path & "\Modules\"

    If Dir(Folder, vbDirectory) = "" Then
        MsgBox "«·„Ã·œ Modules €Ì— „ÊÃÊœ.", vbCritical
        GoTo ExitSub
    End If

    '========================================
    ' „”Õ Ê—ﬁ… Survey ›ﬁÿ
    '========================================
    ThisWorkbook.Worksheets("survey").Cells.Clear

    Set wsDst = ThisWorkbook.Worksheets("survey")

    FirstFile = True

    '========================================
    ' œ„Ã «·„·›« 
    '========================================
    For i = LBound(FileList) To UBound(FileList)

        FileName = FileList(i)

        If Dir(Folder & FileName) <> "" Then

            Set wb = Workbooks.Open(Folder & FileName, ReadOnly:=True)

            Set wsSrc = Nothing

            On Error Resume Next
            Set wsSrc = wb.Worksheets("survey")
            On Error GoTo ErrorHandler

            If Not wsSrc Is Nothing Then

                LastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
                LastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

                If wsDst.Cells(1, 1).Value = "" Then
                    NextRow = 1
                Else
                    NextRow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Row + 1
                End If

                If FirstFile Then

                    '‰”Œ «·⁄‰«ÊÌ‰ Ê«·»Ì«‰« 
                    wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(LastRow, LastCol)).Copy _
                        Destination:=wsDst.Cells(NextRow, 1)

                    FirstFile = False

                Else

                    '‰”Œ «·»Ì«‰«  ›ﬁÿ
                    wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(LastRow, LastCol)).Copy _
                        Destination:=wsDst.Cells(NextRow, 1)

                End If

            Else

                MsgBox "Ê—ﬁ… survey €Ì— „ÊÃÊœ… ›Ì «·„·›: " & FileName, vbExclamation

            End If

            wb.Close SaveChanges:=False

        Else

            MsgBox "«·„·› €Ì— „ÊÃÊœ: " & FileName, vbExclamation

        End If

    Next i

    MsgBox " „ œ„Ã „·›«  Survey »‰Ã«Õ.", vbInformation

ExitSub:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    Exit Sub

ErrorHandler:

    MsgBox "ÕœÀ Œÿ√: " & Err.Description, vbCritical

    Resume ExitSub

End Sub