Attribute VB_Name = "ExportarSlide12K"
Sub ExportarSlides12K()

    Dim caminho As String, nome As String, formato As String, slideAtual As Integer
    Const largura As Integer = 12288, altura As Integer = 6480
    
    usuario = Environ("USERNAME")
    slideAtual = ActiveWindow.Selection.SlideRange.SlideIndex
    nome = Application.ActivePresentation.FullName & "_12K_" & slideAtual
    formato = ".png"
    caminho = "C:\Users\" & usuario & "\Pictures\" & nome & formato
    Application.ActivePresentation.Slides(slideAtual).Export caminho, formato, largura, altura
    MsgBox "Slide Nº " & slideAtual & " foi salvo como: " & caminho, vbExclamation, ""
    
End Sub
