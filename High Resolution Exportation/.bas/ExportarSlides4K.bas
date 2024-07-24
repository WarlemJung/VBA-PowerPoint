Attribute VB_Name = "ExportarSlide4K"
Sub ExportarSlides4K()

    Dim usuario As String, caminho As String, nome As String, formato As String, slideAtual As Integer
    Const largura As Integer = 3840, altura As Integer = 2160
    
    usuario = Environ("USERNAME")
    slideAtual = ActiveWindow.Selection.SlideRange.SlideIndex
    nome = Application.ActivePresentation.FullName & "_4K_" & slideAtual
    formato = ".png"
    caminho = "C:\Users\" & usuario & "\Pictures\" & nome & formato
    Application.ActivePresentation.Slides(slideAtual).Export caminho, formato, largura, altura
    MsgBox "Slide Nº " & slideAtual & " foi salvo como: " & caminho, vbExclamation, ""
    
End Sub
