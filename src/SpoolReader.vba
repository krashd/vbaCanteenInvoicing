Dim arrSpoolData() As Variant
Dim arrSplitData() As Variant
Dim arrProcessedData() As Variant
Dim arrProductLines() As Variant
Dim wbCurrent As Workbook
Dim DestWS As Worksheet
Dim FinalMessage As String

Dim boolReadingHeader As Boolean

Const selArrayMethod = 0

Dim lReadingPosition As Long

Dim PageNo As Long
Dim InvoiceNo As String
Dim ClientAddress(4) As String
Dim SenderAddress(4) As String
Dim StoreNo As String
Dim InvoiceDate As Date
Dim CustomerOrder As String
Dim OrderDate As Date
Dim DispatchDate As Date
Dim ProductCode() As String
Dim ProductDescription() As String
Dim PackSize() As Long
Dim QtyDelivered() As Long
Dim Price() As Double
Dim Ammount() As Double
Dim VAT() As Double
Dim VATBands(2) As Double
Dim GrossTotal(3) As Double
Dim VATTotal(3) As Double
Dim NETTotal(3) As Double
Dim TotalCases As String
Dim RouteDrop As String
Dim SLAccountNo As String

Private Sub setupVariables()
    boolReadingHeader = True
    FinalMessage = ""
    Set wbCurrent = ThisWorkbook
    DebugText.initialiseSettings
    Call eraseSmallArrays
    Call eraseSingleArray(arrSplitData)
End Sub

Private Sub eraseSingleArray(arrayName As Variant)
    If GetArrayLength(arrayName) > -1 Then Erase arrayName
End Sub

Private Sub eraseSmallArrays()
    Call eraseSingleArray(arrSpoolData)
    Call eraseSingleArray(arrProcessedData)
    Call eraseSingleArray(arrProductLines)
    Call eraseSingleArray(ProductCode)
    Call eraseSingleArray(ProductDescription)
    Call eraseSingleArray(PackSize)
    Call eraseSingleArray(QtyDelivered)
    Call eraseSingleArray(Price)
    Call eraseSingleArray(Ammount)
    Call eraseSingleArray(VAT)
    Call eraseSingleArray(VATBands)
End Sub

Private Sub cleanUp()
    DebugText.cleanUp
    Set wbCurrent = Nothing
    Call eraseSmallArrays
    Call eraseSingleArray(arrSplitData)
End Sub

Public Sub readSpoolData()
    Dim arrProcessedString() As Variant
    Dim arrSplitString() As String

    Dim lProcessedStringPosition As Long
    Dim lSplitStringArrayPosition As Long
    
    Dim inputFileName As String
    
    Dim fileNum As Long
    Dim dataLine As Variant
    Dim fileName As String
    Dim i As Long
    Dim x As Long
        
    Call setupVariables
    
    inputFileName = Application.GetOpenFilename("Spool File (SPOOL.*), SPOOL.*", , "Select Spool File")
    If inputFileName = "False" Then
        MsgBox "No Spool File Selected, Aborting Import"
        Exit Sub
    End If
    
    fileNum = FreeFile()
    
    Open inputFileName For Input As #fileNum    'Open input file for reading

    lProcessedStringPosition = 0
    
    ReDim arrProcessedString(0)
    
    While Not EOF(fileNum)
        Line Input #fileNum, dataLine
        arrSplitString = Split(dataLine, vbLf)              'Split the incoming data line into seperate lines
        
        For lSplitStringArrayPosition = 0 To UBound(arrSplitString)
            ReDim Preserve arrProcessedString(lProcessedStringPosition)
            arrProcessedString(lProcessedStringPosition) = Replace(arrSplitString(lSplitStringArrayPosition), vbLf, vbCrLf)
            lProcessedStringPosition = lProcessedStringPosition + 1
        Next lSplitStringArrayPosition
        
    Wend
    Close fileNum
    
    For lProcessedStringPosition = 0 To UBound(arrProcessedString)
        ReDim Preserve arrSplitData(Checks.GetArrayLength(arrSplitData) + 1)
        arrSplitData(UBound(arrSplitData)) = arrProcessedString(lProcessedStringPosition)
    Next lProcessedStringPosition
    
    DebugText.PrintText "Done Reading File, Processing."
    
    Call buildBlock
'    Call outputTestData
    Call cleanUp
End Sub

Private Sub outputTestData()
    Dim lArrayLength As Long
    Dim lCurrentPosition As Long
    Dim strDebugText As String
    
    lArrayLength = Checks.GetArrayLength(arrProcessedData)
    
    For lCurrentPosition = 0 To lArrayLength
        strDebugText = arrProcessedData(0, lCurrentPosition) & ", " & arrProcessedData(16, lCurrentPosition)
        DebugText.PrintText strDebugText
    Next lCurrentPosition
End Sub

Private Sub testybits()
    'Const constTest = "SHP TUNA & CUCUMBER         X4         "
    'Dim strTest As String
    'Dim strOutput As String
    'Dim strOutput2 As String
    
    Dim arrTest(1, 4, 0) As Variant
    Dim lArrayRank As Long
    Dim lArryaLength As Long
        
    arrTest(0, 0, 0) = "Test1"
    arrTest(1, 0, 0) = "Test2"
    arrTest(0, 1, 0) = "Test3"
    arrTest(1, 1, 0) = "Test4"
    arrTest(0, 2, 0) = 5
    arrTest(0, 3, 0) = 6
    arrTest(0, 4, 0) = 7
    arrTest(1, 2, 0) = 8
    arrTest(1, 3, 0) = 9
    arrTest(1, 4, 0) = 10
    
    lArrayRank = Checks.countDimensions(arrTest)
    lArrayLength = Checks.GetArrayLength(arrTest)
    
    MsgBox "Array Dimensions: " & lArrayRank & ", Array Length: " & lArrayLength
    
    'strTest = Trim(constTest)
    'strOutput = Trim(Mid(strTest, Len(strTest) - 0, 1))
    'strOutput2 = Trim(Left(strTest, Len(strTest) - 2))
    
    'MsgBox """" & constTest & """" & vbCrLf & """" & strTest & """" & vbCrLf & """" & strOutput & """" & vbCrLf & """" & strOutput2 & """"
End Sub

Private Sub buildBlock()
    Dim lProcessedArrayPosition As Long
    Dim lArraySize As Long
    Dim lArrayPosition As Long
    Dim lNewArraySize As Long
    Dim lCurrentProcessingPosition As Long
    Dim strTemporary As String
    Dim strProcessing As String
    Dim strPackSize As String
    Dim boolStartedReading As Boolean
    Dim boolSpansPages As Boolean
    
    'Call ADO_Conn.Open_Connection
    
    ReDim arrProcessedData(22, 0)
    
    lArraySize = UBound(arrSplitData)
    
    For lProcessedArrayPosition = 0 To lArraySize
        strProcessing = arrSplitData(lProcessedArrayPosition)
        
        If boolStartedReading = False Then
            If Len(strProcessing) > 20 Then
                If Len(strProcessing) >= 67 Then
                    strTemporary = Trim(strProcessing)
                    If Len(strTemporary) = 1 Or Len(strTemporary) = 2 Then
                        If IsNumeric(strTemporary) Then
                            'PageNo = strTemporary
                            lCurrentProcessingPosition = 0
                            boolStartedReading = True
                        End If
                    End If
                End If
            End If
        End If
        
        If boolStartedReading = True Then
            If Len(strProcessing) > 20 Then
                If boolSpansPages = False Then
                    If lCurrentProcessingPosition = 3 Then
                        PageNo = 1
                        InvoiceNo = Trim(Mid(strProcessing, 68))
                    End If
                
                    If lCurrentProcessingPosition = 5 Then
                        ClientAddress(0) = Trim(Mid(strProcessing, 2, 30))
                        SenderAddress(0) = Trim(Mid(strProcessing, 35, 30))
                    End If
                
                    If lCurrentProcessingPosition = 6 Then
                        ClientAddress(1) = Trim(Mid(strProcessing, 2, 30))
                        SenderAddress(1) = Trim(Mid(strProcessing, 35, 30))
                        StoreNo = Trim(Mid(strProcessing, 68, 10))
                    End If
                
                    If lCurrentProcessingPosition = 7 Then
                        ClientAddress(2) = Trim(Mid(strProcessing, 2, 30))
                        SenderAddress(2) = Trim(Mid(strProcessing, 35, 30))
                    End If
                
                    If lCurrentProcessingPosition = 8 Then
                        ClientAddress(3) = Trim(Mid(strProcessing, 2, 30))
                        SenderAddress(3) = Trim(Mid(strProcessing, 35, 30))
                    End If
                
                    If lCurrentProcessingPosition = 9 Then
                        ClientAddress(4) = enterSomeValue(Trim(Mid(strProcessing, 2, 30)))
                        'If Len(Trim(Mid(strProcessing, 2, 30))) > 0 Then
                        '    ClientAddress(4) = Trim(Mid(strProcessing, 2, 30))
                        'End If
                        SenderAddress(4) = Trim(Mid(strProcessing, 35, 30))
                        InvoiceDate = Format(Trim(Mid(strProcessing, 68, 8)), "dd/mm/yy")
                    End If
                
                    If lCurrentProcessingPosition = 12 Then
                        CustomerOrder = Trim(Mid(strProcessing, 2, 13))
                        OrderDate = Format(Trim(Mid(strProcessing, 19, 8)), "dd/mm/yy")
                        DispatchDate = Format(Trim(Mid(strProcessing, 53, 8)), "dd/mm/yy")
                    End If
                End If
                
                If lCurrentProcessingPosition >= 14 And lCurrentProcessingPosition <= 48 Then
                    If Len(strProcessing) > 0 Then
                        lNewArraySize = Checks.GetArrayLength(ProductCode) + 1
                        
                        ReDim Preserve ProductCode(lNewArraySize)
                        ReDim Preserve ProductDescription(lNewArraySize)
                        ReDim Preserve PackSize(lNewArraySize)
                        ReDim Preserve QtyDelivered(lNewArraySize)
                        ReDim Preserve Price(lNewArraySize)
                        ReDim Preserve Ammount(lNewArraySize)
                        ReDim Preserve VAT(lNewArraySize)
                        
                        ProductCode(lNewArraySize) = Trim(Mid(strProcessing, 2, 15))
                        strTemporary = Trim(Mid(strProcessing, 18, 34))
                        strPackSize = Trim(Mid(strTemporary, Len(strTemporary) - 1, 2))
                        If Left(strPackSize, 1) = "x" Or Left(strPackSize, 1) = "X" Then
                            PackSize(lNewArraySize) = Mid(strPackSize, 2, 1)
                            ProductDescription(lNewArraySize) = Trim(Left(strTemporary, Len(strTemporary) - 2))
                        Else
                            PackSize(lNewArraySize) = strPackSize
                            ProductDescription(lNewArraySize) = Trim(Left(strTemporary, Len(strTemporary) - 3))
                        End If
                        QtyDelivered(lNewArraySize) = Trim(Mid(strProcessing, 55, 3))
                        Price(lNewArraySize) = Trim(Mid(strProcessing, 59, 9))
                        Ammount(lNewArraySize) = Trim(Mid(strProcessing, 69, 7))
                        strTemporary = Trim(Mid(strProcessing, 77))
                        VAT(lNewArraySize) = Left(strTemporary, Len(strTemporary) - 1)
                        If VAT(lNewArraySize) > 0 Then VAT(lNewArraySize) = VAT(lNewArraySize) / 100
                        'Call updateArray
                    End If
                End If
            End If
                
            If lCurrentProcessingPosition = 50 Then
                If Len(strProcessing) < 10 Then
                    boolSpansPages = True
                Else
                    TotalCases = Trim(strProcessing)
                    boolSpansPages = False
                End If
            End If
                
            If Len(strProcessing) > 20 Then
                If boolSpansPages = False Then
                    If lCurrentProcessingPosition = 51 Then
                        RouteDrop = Trim(strProcessing)
                    End If
                
                    If lCurrentProcessingPosition >= 52 And lCurrentProcessingPosition <= 54 Then
                        lArrayPosition = (lCurrentProcessingPosition - 52)
                        VATBands(lArrayPosition) = Trim(Mid(strProcessing, 51, 4))
                        If VATBands(lArrayPosition) > 0 Then VATBands(lArrayPosition) = VATBands(lArrayPosition) / 100
                        GrossTotal(lArrayPosition) = Trim(Mid(strProcessing, 58, 7))
                        VATTotal(lArrayPosition) = Trim(Mid(strProcessing, 66, 6))
                        NETTotal(lArrayPosition) = Trim(Mid(strProcessing, 74, 7))
                    End If
                
                    If lCurrentProcessingPosition = 57 Then
                        SLAccountNo = Trim(Left(strProcessing, 51))
                        GrossTotal(3) = Trim(Mid(strProcessing, 58, 7))
                        VATTotal(3) = Trim(Mid(strProcessing, 66, 6))
                        NETTotal(3) = Trim(Mid(strProcessing, 74, 7))
                    End If
                End If
            End If
            
            If lCurrentProcessingPosition = 61 Then
                boolStartedReading = False
                If boolSpansPages = False Then
                    Call Copy_Template(StoreNo)
                    Call write_data
                    Call Transfer_Invoices(StoreNo)
                    Call eraseSmallArrays
                End If
                lCurrentProcessingPosition = 0
            End If
        End If
        
        lCurrentProcessingPosition = lCurrentProcessingPosition + 1
    Next lProcessedArrayPosition
    
    DebugText.PrintText FinalMessage
    MsgBox FinalMessage
End Sub

Private Sub updateArray()
    Dim lArrayPosition As Long
    
    ReDim Preserve arrProcessedData(0 To 22, 0 To Checks.GetArrayLength(arrProcessedData) + 1)
    lArrayPosition = Checks.GetArrayLength(arrProcessedData)
    arrProcessedData(0, lArrayPosition) = InvoiceNo
    arrProcessedData(1, lArrayPosition) = ClientAddress(0)
    arrProcessedData(2, lArrayPosition) = ClientAddress(1)
    arrProcessedData(3, lArrayPosition) = ClientAddress(2)
    arrProcessedData(4, lArrayPosition) = enterSomeValue(ClientAddress(3))
    arrProcessedData(5, lArrayPosition) = enterSomeValue(ClientAddress(4))
    arrProcessedData(6, lArrayPosition) = SenderAddress(0)
    arrProcessedData(7, lArrayPosition) = SenderAddress(1)
    arrProcessedData(8, lArrayPosition) = SenderAddress(2)
    arrProcessedData(9, lArrayPosition) = enterSomeValue(SenderAddress(3))
    arrProcessedData(10, lArrayPosition) = enterSomeValue(SenderAddress(4))
    arrProcessedData(11, lArrayPosition) = StoreNo
    arrProcessedData(12, lArrayPosition) = InvoiceDate
    arrProcessedData(13, lArrayPosition) = CustomerOrder
    arrProcessedData(14, lArrayPosition) = OrderDate
    arrProcessedData(15, lArrayPosition) = DispatchDate
    arrProcessedData(16, lArrayPosition) = ProductCode
    arrProcessedData(17, lArrayPosition) = ProductDescription
    arrProcessedData(18, lArrayPosition) = PackSize
    arrProcessedData(19, lArrayPosition) = QtyDelivered
    arrProcessedData(20, lArrayPosition) = Price
    arrProcessedData(21, lArrayPosition) = Ammount
    arrProcessedData(22, lArrayPosition) = VAT
End Sub

Private Function Copy_Template(BranchNo As String) As Long
        
    On Error Resume Next
    Set DestWS = wbCurrent.Sheets(BranchNo)
    On Error GoTo 0
    
    If DestWS Is Nothing Then
        'create the DestWS
        wbCurrent.Sheets("Canteen Template").Copy after:=Sheets("Canteen Template")
        'Set DestWS = wbCurrent.Sheets.Add(Type:="I:\Dronfield\Stores\Weekly Reports\Canteen Invoices\Templates\Canteen Invoice.xltx", after:=wbCurrent.Worksheets(Worksheets.Count))
        DebugText.PrintText "Opening template for usage."
        Set DestWS = ActiveSheet
        DestWS.Visible = xlSheetVisible
        DestWS.Name = BranchNo
    End If
    
End Function

Private Function enterSomeValue(varInput As Variant) As Variant
    If varInput <> "" Then
        enterSomeValue = varInput
    ElseIf Not IsNull(varInput) Then
        enterSomeValue = varInput
    Else
        enterSomeValue = ""
    End If
End Function

Private Function valuePresent(strInput As String) As Boolean
    If strInput <> "" Then valuePresent = True
    If Not IsNull(strInput) Then valuePresent = True
End Function

Private Function write_data() As Long
    Dim lArrayLength As Long
    Dim i As Long
    
    DestWS.Range("J2").Value = PageNo
    DestWS.Range("J5").Value = InvoiceNo
    DestWS.Range("J8").Value = StoreNo
    For i = 0 To 4
        DestWS.Cells(i + 7, 1).Value = ClientAddress(i)
        DestWS.Cells(i + 7, 5).Value = SenderAddress(i)
    Next i
    DestWS.Range("J11").Value = InvoiceDate
    DestWS.Range("A14").Value = CustomerOrder
    DestWS.Range("C14").Value = OrderDate
    DestWS.Range("G14").Value = DispatchDate
    DebugText.PrintText "Writing Product List."
    
    lArrayLength = Checks.GetArrayLength(ProductCode)
    
    For i = 0 To lArrayLength
        DestWS.Range("A" & i + 16).Value = ProductCode(i)
        DestWS.Range("C" & i + 16).Value = ProductDescription(i)
        DestWS.Range("G" & i + 16).Value = PackSize(i)
        DestWS.Range("H" & i + 16).Value = QtyDelivered(i)
        DestWS.Range("I" & i + 16).Value = Price(i)
        DestWS.Range("J" & i + 16).Value = Ammount(i)
        DestWS.Range("K" & i + 16).Value = VAT(i)
    Next i
    DestWS.Range("A62").Value = TotalCases
    DestWS.Range("A63").Value = RouteDrop
    DestWS.Range("A66").Value = SLAccountNo
    For i = 0 To 2
        If GrossTotal(i) > 0 Then
            DestWS.Range("G" & i + 64).Value = VATBands(i)
            DestWS.Range("I" & i + 64).Value = GrossTotal(i)
            DestWS.Range("J" & i + 64).Value = VATTotal(i)
            DestWS.Range("K" & i + 64).Value = NETTotal(i)
        Else
            DestWS.Range("G" & i + 64).Value = ""
            DestWS.Range("I" & i + 64).Value = ""
            DestWS.Range("J" & i + 64).Value = ""
            DestWS.Range("K" & i + 64).Value = ""
        End If
    Next i
    DestWS.Range("I67").Value = GrossTotal(3)
    DestWS.Range("J67").Value = VATTotal(3)
    DestWS.Range("K67").Value = NETTotal(3)
    
    Set DestWS = Nothing
    For i = 0 To 2
        VATBands(i) = 0
        GrossTotal(i) = 0
        VATTotal(i) = 0
        NETTotal(i) = 0
    Next i
    GrossTotal(3) = 0
    VATTotal(3) = 0
    NETTotal(3) = 0
End Function

Private Function Transfer_Invoices(tabName As String) As Long
    Dim DestWB As Workbook
    Dim SourceWS As Worksheet
    Dim invoice_dispatch_date As Date
    Dim stYear As String, stMonth As String, stDay As String
    Dim baseLoc As String, saveLoc As String
    
    Debug.Print "Setting Sheet to transfer."
    DebugText.PrintText "Setting Sheet to transfer."
    
    Set SourceWS = wbCurrent.Sheets(tabName)
    invoice_dispatch_date = SourceWS.Range("G14").Value
    stYear = Format(invoice_dispatch_date, "yyyy")
    stMonth = Format(invoice_dispatch_date, "mm")
    stDay = Format(invoice_dispatch_date, "dd")
            
    baseLoc = "I:\Dronfield\Stores\Weekly Reports\Canteen Invoices"
    saveLoc = baseLoc & "\" & stYear & "\" & stMonth & "\" & stYear & "-" & stMonth & "-" & stDay & ".xlsx"
    
    If Not FileExists(saveLoc) Then
        Call CheckDir(baseLoc & "\" & stYear)
        Call CheckDir(baseLoc & "\" & stYear & "\" & stMonth)
        
        Debug.Print "Destination Workbook doesn't exist, creating new workbook."
        DebugText.PrintText "Destination Workbook doesn't exist, creating new workbook."
        Set DestWB = Workbooks.Add()
        DestWB.SaveAs fileName:=saveLoc
    Else
        Debug.Print "Destination Workbook exists, opening workbook."
        DebugText.PrintText " Destination Workbook exists, opening workbook."
        Set DestWB = Workbooks.Open(fileName:=saveLoc)
    End If
            
    Debug.Print "Moving sheet to destination workbook."
    DebugText.PrintText "Moving Sheet to Destination Workbook."
    SourceWS.Move after:=DestWB.Worksheets(DestWB.Worksheets.Count)
            
    'test to see if the destination sheets(sheet1, sheet2, sheet3) exist
    If Checks.Check_Worksheet_Exists(DestWB, "Sheet1") Then Call Checks.Quiet_Delete_Worksheet(DestWB.Sheets("Sheet1"))
    If Checks.Check_Worksheet_Exists(DestWB, "Sheet2") Then Call Checks.Quiet_Delete_Worksheet(DestWB.Sheets("Sheet2"))
    If Checks.Check_Worksheet_Exists(DestWB, "Sheet3") Then Call Checks.Quiet_Delete_Worksheet(DestWB.Sheets("Sheet3"))
    
    Debug.Print "Updating Final Message"
    DebugText.PrintText "Updating Final Message."
    FinalMessage = FinalMessage & "Invoice " & tabName & " moved to file " & stYear & "-" & stMonth & "-" & stDay & ".xls" & vbCrLf
    
    Debug.Print "Closing the destination workbook."
    DebugText.PrintText "Closing the destination workbook."
    DestWB.Close savechanges:=True
    Set DestWB = Nothing
End Function
