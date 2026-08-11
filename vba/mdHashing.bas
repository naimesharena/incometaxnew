Attribute VB_Name = "mdHashing"
Public Function Base64_HMACSHA256(ByVal iteration As Long, ByVal sTextToHash As String, ByVal sSharedSecretKey As String)

    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")

    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.Key = SharedSecretKey

    Dim byteS() As Byte
    byteS = enc.ComputeHash_2((TextToHash))
    
    For i = 1 To iteration
        
        byteS = enc.ComputeHash_2((byteS))
    Next
     
    Base64_HMACSHA256 = EncodeBase64(byteS)
    Set asc = Nothing
    Set enc = Nothing

End Function
Public Function Base64_HMACSHA256_test(ByVal iteration As Long, sTextToHash() As Variant, ByVal sSharedSecretKey As String)

    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    
    Dim tempTextToHash() As Byte
    Dim tempString, CountTextToHash As Variant
    Dim i, j As Variant
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")
    ReDim TextToHash(0)
    subProcCaption = "Encrypting XML"
    noOfProcessSub = UBound(sTextToHash)
    For i = 0 To UBound(sTextToHash)
        tempString = sTextToHash(i)
        tempTextToHash = asc.Getbytes_4(tempString)
        CountTextToHash = UBound(TextToHash)
        If i > 0 Then
            ReDim Preserve TextToHash(CountTextToHash + UBound(tempTextToHash) + 1)
            For j = 0 To UBound(tempTextToHash)
                CountTextToHash = CountTextToHash + 1
                TextToHash(CountTextToHash) = tempTextToHash(j)
            Next
        Else
            ReDim Preserve TextToHash(CountTextToHash + UBound(tempTextToHash))
            For j = 0 To UBound(tempTextToHash)
                TextToHash(j) = tempTextToHash(j)
            Next
        End If
        UpdateProgressBar
     Next
    
   ' TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.Key = SharedSecretKey

    Dim byteS() As Byte
    byteS = enc.ComputeHash_2((TextToHash))
    
    'for test purpose only not required for xml generation
   ' Open ThisWorkbook.Path & "/testnew.txt" For Output As #2
   ' For i = 0 To UBound(TextToHash)
   '     Print #2, TextToHash(i)
   ' Next
    'Close #2
    
    For i = 1 To iteration
        byteS = enc.ComputeHash_2((byteS))
    Next
     
    Base64_HMACSHA256_test = EncodeBase64(byteS)
    Set asc = Nothing
    Set enc = Nothing

End Function
'function added by Chetan C M 03/01/2025
Function HMACSHA256A(strToSign As String, strKey() As Byte, ByVal iteration As Long)

On Error GoTo endline

    Dim lngLoop As Long

    Dim oUTF, oEnc

    Dim HMAC() As Byte

    Dim lastrow As Long

    Dim byteString() As Byte

     Dim byteS() As Byte



    On Error GoTo err_handler

    Set test = New HS256

    test.InitHmac strKey

    byteString = StrConv(strToSign, vbFromUnicode)

'    byteString = Test.ToUTF8(strToSign)

    byteS = test.HMACSHA256(byteString)

    Dim i, j As Variant
    
    UpdateProgressBar

     For i = 1 To iteration

        test.InitHmac strKey

        byteString = ((byteS))

        byteS = test.HMACSHA256(byteString)
        

    Next

    HMACSHA256A = EncodeBase64(byteS)

'    Worksheets("Log Sheet").Cells(lastrow, 4) = "Pass"

    Exit Function



err_handler:

'    Worksheets("Log Sheet").Cells(lastrow, 4) = "Fail"

'    Worksheets("Log Sheet").Cells(lastrow, 5) = Err.Description

    'MsgBox Err.Description, vbCritical

endline:

'MsgBox ("033")

End Function


Private Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing

End Function
