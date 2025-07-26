VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
CriarBitmap32Bits
End Sub
Sub CriarBitmap32Bits()

    Dim largura As Long
    Dim altura As Long
    Dim corVGA As Integer
    Dim vgaColors(0 To 15) As Long
    Dim i As Long

    ' Paleta de cores VGA (formato BGR - byte order invertido)
    vgaColors(0) = RGB(0, 0, 0)
    vgaColors(1) = RGB(0, 0, 170)
    vgaColors(2) = RGB(0, 170, 0)
    vgaColors(3) = RGB(0, 170, 170)
    vgaColors(4) = RGB(170, 0, 0)
    vgaColors(5) = RGB(170, 0, 170)
    vgaColors(6) = RGB(170, 85, 0)
    vgaColors(7) = RGB(170, 170, 170)
    vgaColors(8) = RGB(85, 85, 85)
    vgaColors(9) = RGB(85, 85, 255)
    vgaColors(10) = RGB(85, 255, 85)
    vgaColors(11) = RGB(85, 255, 255)
    vgaColors(12) = RGB(255, 85, 85)
    vgaColors(13) = RGB(255, 85, 255)
    vgaColors(14) = RGB(255, 255, 85)
    vgaColors(15) = RGB(255, 255, 255)

    largura = CLng(InputBox("Digite a largura do bitmap:", "Largura"))
    altura = CLng(InputBox("Digite a altura do bitmap:", "Altura"))
    corVGA = CInt(InputBox("Digite o número da cor VGA (0-15):", "Cor VGA"))

    If corVGA < 0 Or corVGA > 15 Then
        MsgBox "Cor inválida!"
        Exit Sub
    End If

    Dim fileHeaderSize As Long: fileHeaderSize = 14
    Dim infoHeaderSize As Long: infoHeaderSize = 40
    Dim pixelDataSize As Long: pixelDataSize = largura * altura * 4 ' 4 bytes por pixel (BGRA)
    Dim fileSize As Long: fileSize = fileHeaderSize + infoHeaderSize + pixelDataSize

    Dim bmpBytes() As Byte
    ReDim bmpBytes(0 To fileSize - 1)

    ' === Cabeçalho BMP (14 bytes) ===
    bmpBytes(0) = &H42 ' "B"
    bmpBytes(1) = &H4D ' "M"
    PutLong bmpBytes, 2, fileSize          ' Tamanho total do ficheiro
    PutLong bmpBytes, 6, 0                 ' Reservado
    PutLong bmpBytes, 10, fileHeaderSize + infoHeaderSize ' Offset dados

    ' === Cabeçalho DIB (40 bytes) ===
    PutLong bmpBytes, 14, infoHeaderSize   ' Tamanho cabeçalho DIB
    PutLong bmpBytes, 18, largura
    PutLong bmpBytes, 22, altura
    PutInt bmpBytes, 26, 1                 ' Planos
    PutInt bmpBytes, 28, 32                ' Bits por pixel
    PutLong bmpBytes, 30, 0                ' Sem compressão
    PutLong bmpBytes, 34, pixelDataSize
    PutLong bmpBytes, 38, 2835             ' Resol. horizontal (px/m)
    PutLong bmpBytes, 42, 2835             ' Resol. vertical (px/m)
    PutLong bmpBytes, 46, 0                ' Cores na paleta
    PutLong bmpBytes, 50, 0                ' Cores importantes

    ' === Dados dos pixels (BGRA) ===
    Dim r As Byte, g As Byte, b As Byte
    b = vgaColors(corVGA) And &HFF
    g = (vgaColors(corVGA) \ &H100) And &HFF
    r = (vgaColors(corVGA) \ &H10000) And &HFF

    Dim pos As Long: pos = fileHeaderSize + infoHeaderSize
    For i = 1 To largura * altura
        bmpBytes(pos) = b: pos = pos + 1
        bmpBytes(pos) = g: pos = pos + 1
        bmpBytes(pos) = r: pos = pos + 1
        bmpBytes(pos) = 0: pos = pos + 1 ' Alpha = 0
    Next i

    ' === Gravação para disco ===
    Dim caminho As String
    caminho = ".\new.bmp"
    MsgBox (CurDir())
    Dim f As Integer: f = 1
    Open caminho For Binary As #1
    Put #1, , bmpBytes
    Close #1

    MsgBox "Imagem gravada em: " & caminho

End Sub

' === Funções auxiliares ===
Private Sub PutLong(ByRef arr() As Byte, ByVal offset As Long, ByVal value As Long)
    arr(offset) = value And &HFF
    arr(offset + 1) = (value \ &H100) And &HFF
    arr(offset + 2) = (value \ &H10000) And &HFF
    arr(offset + 3) = (value \ &H1000000) And &HFF
End Sub

Private Sub PutInt(ByRef arr() As Byte, ByVal offset As Long, ByVal value As Integer)
    arr(offset) = value And &HFF
    arr(offset + 1) = (value \ &H100) And &HFF
End Sub

Private Sub UserForm_Click()

End Sub
