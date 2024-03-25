VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   555
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   2500
      Left            =   1845
      Top             =   945
      Width           =   2500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function GerarQRCodePIX(strChave As String, strValor As String, strNomeLoja As String, strCidadeLoja As String) As String
    Dim strFinal As String
    Dim strURL As String
    
    
    Dim sMerchantCategoryCode As String
    Dim sMerchantAccountInformation As String
    Dim sTransactionCurrency As String
    Dim sTransactionValue As String
    Dim sMerchantName  As String
    Dim sMerchantCity As String
    
    
 
    strFinal = "000201"
    
    strURL = "br.gov.bcb.pix"
    strURL = UCase(strURL)
    
    sMerchantAccountInformation = "00" & Format(Len(strURL), "00") & strURL
    sMerchantAccountInformation = sMerchantAccountInformation & "01" & Len(strChave) & strChave
    
    strFinal = strFinal & "26" & Len(sMerchantAccountInformation) & sMerchantAccountInformation '26 - Merchant Account Informatio
    
    
    sMerchantCategoryCode = "52" & Format(Len("0000"), "00") & "0000"
    strFinal = strFinal & sMerchantCategoryCode
    
    sTransactionCurrency = "53" & Format(Len("986"), "00") & "986"
    strFinal = strFinal & sTransactionCurrency
    
    sTransactionValue = "54" & Format(Len(strValor), "00") & strValor
    strFinal = strFinal & sTransactionValue
    
    strFinal = strFinal & "5802BR" 'CountryCode
    
    sMerchantName = "59" & Format(Len(strNomeLoja), "00") & strNomeLoja
    strFinal = strFinal & sMerchantName
    
    sMerchantCity = "60" & Format(Len(strCidadeLoja), "00") & strCidadeLoja
    strFinal = strFinal & sMerchantCity
    
    strFinal = strFinal & "62070503***" 'Additional Data Field Template
    strFinal = strFinal & "6304"
    strFinal = strFinal & GetCRC16(strFinal)
    GerarQRCodePIX = strFinal
End Function
Private Function GetCRC16(payload As String) As String
    Dim polinomio As Long
    Dim resultado As Long
    Dim i As Integer
    Dim j As Integer
    Dim length As Integer
    Dim offset As Integer
    Dim bitwise As Integer
    Dim tempByte As Integer
    
 
    
    ' Dados definidos pelo BACEN
    polinomio = &H1021
    resultado = &HFFFF
    
    ' Checksum
    length = Len(payload)
    If length > 0 Then
        For offset = 1 To length
            resultado = resultado Xor (Asc(Mid$(payload, offset, 1)) * &H100)
            For bitwise = 0 To 7
                If (resultado And &H8000) <> 0 Then
                    resultado = (resultado And &H7FFF) * 2
                    resultado = resultado Xor polinomio
                Else
                    resultado = (resultado And &HFFFF) * 2
                End If
            Next bitwise
        Next offset
    End If
    resultado = resultado And &HFFFF
    GetCRC16 = UCase$(Right$("0000" & Hex(resultado), 4))
End Function

Public Sub PrintImagem(p As IPictureDisp, Optional ByVal X, Optional ByVal y, Optional ByVal resize)
    If IsMissing(resize) Then resize = 0.2
    If IsMissing(y) Then y = Printer.CurrentY

    Dim imageWidth As Long
    imageWidth = p.Width * resize
    
    If IsMissing(X) Then
        X = (Printer.Width - imageWidth) / 2
    End If

    Printer.FontBold = True


    Printer.PaintPicture p, X, y, imageWidth, p.Height * resize

    Printer.CurrentY = y + p.Height * resize
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    Printer.Print " Leia o QR-Code no aplicativo do seu banco"
    Printer.EndDoc
    
End Sub
Public Function SelectPrinter(ByVal Nome As String) As Boolean
    Dim X As Printer
    For Each X In Printers
    If UCase(Mid(X.DeviceName, 1, 8)) = UCase(Mid(Nome, 1, 8)) Then
    Set Printer = X
    SelectPrinter = True
    Exit For
    End If
    Next
    SelectPrinter = False
End Function

Private Sub Command1_Click()
  
  PrintImagem Image1.Picture
End Sub

Private Sub Form_Load()
    Set Image1.Picture = QRCodegenBarcode(GerarQRCodePIX("seuemail@seudominio.com", "18.30", "Nome do Recebedor", "Cidade do Recebedor"))
    SelectPrinter ("Sua Impressora aqui")
    Command1_Click
End Sub
