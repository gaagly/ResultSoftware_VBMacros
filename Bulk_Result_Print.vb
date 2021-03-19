Sub PrintFullResult()
    Dim RollNo As Integer
    Dim GoodFileName As String
    RollNo = 0
    For RollNo = 1 To 39
        Range("d9").Value = RollNo
        If RollNo < 10 Then
            GoodFileName = "0" & CStr(RollNo) & CStr(Range("d10").Value)
        Else
            GoodFileName = CStr(RollNo) & CStr(Range("d10").Value)
        End If
        ActiveSheet.ExportAsFixedFormat _
        Type:=x1TypePDF, _
        Filename:=GoodFileName, _
        Quality:=x1QualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        From:=1, _
        To:=5, _
        OpenAfterPublish:=False
    Next RollNo
End Sub
