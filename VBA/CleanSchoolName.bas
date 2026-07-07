Attribute VB_Name = "Module1"
Function CleanSchoolName(ByVal txt As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    Dim forbiddenWords As Variant
    Dim word As Variant
    
    ' ÞÇÆãÉ ÇáßáãÇÊ ÇáããäæÚÉ
    forbiddenWords = Array("ãÏÑÓÉ", "ÇÚÏÇÏíÉ", "ËÇäæíÉ", "ááÈäÇÊ", "ááÈäíä", "Èäíä", "ÈäÇÊ", "ÇáãÎÊáØÉ", "ãÎÊáØÉ", "ãÎ", "ÓÇÏÓ", "Úáãí", "ÇáÚáãí", "ÇÏÈí", "ÃÏÈí", "ÇáÇÏÈí", "ÊÑÈíÉ ÎÇÕÉ", "ÅÚÏÇÏíÉ", "ÇáãåäíÉ", "ÞÑÖ","ÓÇÏÓ","Úáãí","ÇÏÈí")
    
    ' 1. ÅÒÇáÉ ÇáäÞÇØ æÇáÃÑÞÇã æÇáÑãæÒ æÇáãÏÇÊ (íÈÞì ÝÞØ ÇáÚÑÈí æÇáãÓÇÝÉ)
    txt = Replace(txt, ".", "")
    txt = Replace(txt, "?", "")
    txt = Replace(txt, "/", "")
    txt = Replace(txt, "(", "")
    txt = Replace(txt, ")", "")
    txt = Replace(txt, "Ü", "")
    ' íãßäß ÅÖÇÝÉ ÇáãÒíÏ ãä ÇáÑãæÒ åäÇ
    
    ' 2. ÅÒÇáÉ ÇáÃÑÞÇã
    For i = 0 To 9
        txt = Replace(txt, i, "")
    Next i
    
    ' 3. ÅÒÇáÉ ÇáßáãÇÊ ÇáããäæÚÉ
    For Each word In forbiddenWords
        txt = Replace(txt, word, "")
    Next word
    
    ' 4. ÅÒÇáÉ ÇáÍÑæÝ ÇáãÝÑÏÉ (ãËá Ú ¡ ã ¡ Ë)
    Dim wordsArray() As String
    wordsArray = Split(txt, " ")
    For i = LBound(wordsArray) To UBound(wordsArray)
        If Len(Trim(wordsArray(i))) > 1 Then
            result = result & wordsArray(i) & " "
        End If
    Next i
    
    CleanSchoolName = Trim(result)
End Function
