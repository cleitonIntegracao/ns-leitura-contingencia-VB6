Attribute VB_Name = "leituraContingencia"
Private Sub Leitura_Contingencia()

Dim loc As String
loc = "CHAVE OU VAARIAVEL QUE CONTEM A CHAVE"

'Abre o arquivo contingencia_ret para leitura EX.de caminho: "D:\NSNFCe\Processados\nsConcluido\contingencia_ret"
n = FreeFile()
Open "CAMINHO DO NSNFCe CLOUD QUE CONTEM O ARQUIVO contingencia_ret" For Binary As #n
conte�do = Input(LOF(n), n)
Close #n
linhas = Split(conte�do, vbCrLf)

'Faz a leitura das linhas no arquivo contingencia_ret para achar a chave correspondente
linhadesejadaret = ""
For Each linha In linhas
    posi��o = InStr(1, linha, loc)
    If posi��o > 0 Then
        linhadesejadaret = linha
        Exit For
    End If
Next linha

'Se achar a linha que contem a chave, retorna a linha com todas as informa��es
If linhadesejadaret <> "" Then
    MsgBox linhadesejadaret
Else
    MsgBox "N�o achou!"
End If
End Sub




