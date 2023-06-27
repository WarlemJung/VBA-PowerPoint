Attribute VB_Name = "ExportarSlide4K"
Sub ExportarSlides4K()

    Dim caminho As String, nome As String, formato As String, slideAtual As Integer
    Const largura As Integer = 3840, altura As Integer = 2160
    
    slideAtual = ActiveWindow.Selection.SlideRange.SlideIndex
    nome = "4K_" & slideAtual
    formato = ".png"
    caminho = "E:\Users\EAGLE\Pictures\" & nome & formato
    Application.ActivePresentation.Slides(slideAtual).Export caminho, formato, largura, altura
    MsgBox "Slide Nº " & slideAtual & " foi salvo como: " & caminho, vbExclamation, ""
    
End Sub
