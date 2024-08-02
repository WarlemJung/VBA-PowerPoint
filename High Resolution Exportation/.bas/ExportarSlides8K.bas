Attribute VB_Name = "ExportarSlide8K"
Sub ExportarSlides8K()

    Dim caminho As String, nome As String, formato As String, slideAtual As Integer
    Const largura As Integer = 7680, altura As Integer = 4320
    
    usuario = Environ("USERNAME")
    slideAtual = ActiveWindow.Selection.SlideRange.SlideIndex
    nome = Application.ActivePresentation.FullName & "_8K_" & slideAtual
    formato = ".png"
    caminho = "C:\Users\" & usuario & "\Pictures\" & nome & formato
    Application.ActivePresentation.Slides(slideAtual).Export caminho, formato, largura, altura
    MsgBox "Slide Nº " & slideAtual & " foi salvo como: " & caminho, vbExclamation, ""
    
End Sub
