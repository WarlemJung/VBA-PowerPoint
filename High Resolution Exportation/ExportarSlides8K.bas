Attribute VB_Name = "ExportarSlide8K"
Sub ExportarSlides8K()

    Dim caminho As String, nome As String, formato As String, slideAtual As Integer
    Const largura As Integer = 7680, altura As Integer = 4320
    
    slideAtual = ActiveWindow.Selection.SlideRange.SlideIndex
    nome = "8K_" & slideAtual
    formato = ".png"
    caminho = "E:\Users\EAGLE\Pictures\" & nome & formato
    Application.ActivePresentation.Slides(slideAtual).Export caminho, formato, largura, altura
    MsgBox "Slide Nº " & slideAtual & " foi salvo como: " & caminho, vbExclamation, ""
    
End Sub
