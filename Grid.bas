Attribute VB_Name = "Grid"
Public Function MarcaLinha(ByVal ParGrid As fpSpread, ByVal parLinha As Long, ByVal ParColunaCodigo As Long, Optional ByVal ParColuna As Long _
        , Optional ByRef ParCodigoRetorno As String, Optional ByVal parForcaLinhaInformada As Boolean, Optional ByVal ParOrdena As Boolean = True)
On Error GoTo Trata
Dim sColuna As Long
With ParGrid
    sColuna = .ActiveCol
    If parLinha = 0 And .UserColAction <> 1 Then Exit Function
    .Col = ParColunaCodigo
    .Row = IIf(parForcaLinhaInformada And parLinha > 0, parLinha, .ActiveRow)
    ParCodigoRetorno = .Text
    .Col = sColuna
End With
Exit Function
Resume
Trata:
MsgBox "", "Grid.MarcaLinha", , , Erl
End Function

