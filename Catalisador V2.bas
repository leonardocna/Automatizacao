Attribute VB_Name = "NewMacros"
Sub DataBringer()
Attribute DataBringer.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
' Substitui o texto informado, de acordo com os dados fornecidos pela planilha base.

Dim xl As New Excel.Application

Set xl = New Excel.Application

xl.Visible = True

'abre o Excel com a base de dados
Excel.Workbooks.Open FileName:="C:\Users\cesar_valerio\Desktop\MMSP Laudo V1.xltm"

'Determina as variáveis, de acordo com as células do excel
var1 = xl.Cells(2, 2).Value
var2 = xl.Cells(2, 4).Value

Windows("Documento1").Activate
Selection.WholeStory
'Substitui as variáveis no texto
    With Selection.Find
        .Text = "#nome"
        .Replacement.Text = var1
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
    With Selection.Find
        .Text = "#empresa"
        .Replacement.Text = var2
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub
Sub AbreModelo()
' Utiliza o modelo informado

'abre o Word
Word.Documents.Open FileName:="C:\Users\cesar_valerio\Desktop\Teste Modelo.docx"

' Copia o texto da fonte
    Windows("Teste Modelo").Activate
    Application.WindowState = wdWindowStateNormal
    Selection.WholeStory
    Selection.Copy
' Cola o texto no arquivo que unifica os modelos
    Windows("Documento1").Activate
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    
    Windows("Teste Modelo").Close

End Sub
