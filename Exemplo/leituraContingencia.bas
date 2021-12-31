Attribute VB_Name = "leituraContingencia"
Private Sub Leitura_Contingencia()

Dim loc As String
loc = "CHAVE OU VAARIAVEL QUE CONTEM A CHAVE"

'Abre o arquivo contingencia_ret para leitura EX.de caminho: "D:\NSNFCe\Processados\nsConcluido\contingencia_ret"
n = FreeFile()
Open "CAMINHO DO NSNFCe CLOUD QUE CONTEM O ARQUIVO contingencia_ret" For Binary As #n
conteúdo = Input(LOF(n), n)
Close #n
linhas = Split(conteúdo, vbCrLf)

'Faz a leitura das linhas no arquivo contingencia_ret para achar a chave correspondente
linhadesejadaret = ""
For Each linha In linhas
    posição = InStr(1, linha, loc)
    If posição > 0 Then
        linhadesejadaret = linha
        Exit For
    End If
Next linha

'Se achar a linha que contem a chave, retorna a linha com todas as informações
If linhadesejadaret <> "" Then
    MsgBox linhadesejadaret
Else
    MsgBox "Não achou!"
End If
End Sub




