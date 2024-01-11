Attribute VB_Name = "CreateMassQuote"

Sub CreateMailMerge()

Application.ScreenUpdating = False
Dim coll As New Collection
Dim lLastRow As Long
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet
Set ws = wb.Sheets("Data Entry")
ws.Activate

answer = MsgBox("Do you want the merge to create individual PDFs? Select No if you want to make your own pdf groupings, otherwise this will split by contract#", vbYesNo)

If answer = vbYes Then
With Application.FileDialog(msoFileDialogFolderPicker)
.Show
myfolder = .SelectedItems(1) & "\"
End With
End If

With ws
lLastRow = .Range("A" & Sheet1.Rows.Count).End(xlUp).Row
End With

Dim oCustomer As clsCustomer
 Dim i As Long
    ' Read through the list of customers and create customer object, then add to collection
    For i = 2 To lLastRow
        Set oCustomer = New clsCustomer
            oCustomer.BillToCustomerName = Range("A" & i)
            oCustomer.BillToAddress = Range("B" & i)
            oCustomer.BillToTown = Range("C" & i)
            oCustomer.BillToState = Range("D" & i)
            oCustomer.BillToZipCode = Range("E" & i)
            oCustomer.BillToContactName = Range("F" & i)
            oCustomer.BillToPhoneNumber = Range("G" & i)
            oCustomer.BillToFaxNumber = Range("H" & i)
            oCustomer.BillToEmail = Range("I" & i)
            oCustomer.ShipToCustomerName = Range("J" & i)
            oCustomer.ShipToAddress = Range("K" & i)
            oCustomer.ShipToTown = Range("L" & i)
            oCustomer.ShipToState = Range("M" & i)
            oCustomer.ShipToZipCode = Range("N" & i)
            oCustomer.ShipToContactName = Range("O" & i)
            oCustomer.ShipToPhoneNumber = Range("P" & i)
            oCustomer.ShipToFaxNumber = Range("Q" & i)
            oCustomer.ShipToEmail = Range("R" & i)
            oCustomer.ContractAwardNumber = Range("S" & i)
            oCustomer.CurrentPOPStartDate = Range("T" & i)
            oCustomer.CurrentPOPEndDate = Range("U" & i)
            oCustomer.QuoteInfoEmailAddress = Range("V" & i)
            oCustomer.QuoteInfoBillerFirstName = Range("W" & i)
            oCustomer.QuoteInfoBillerLastName = Range("X" & i)
            oCustomer.QuoteInfoDate = Range("Y" & i)
            oCustomer.QuoteInfoAppendix = Range("Z" & i)
            oCustomer.QuoteInfoBillersManagerName = Range("AA" & i)
            oCustomer.QuoteInfoQuoteNumber = Range("AB" & i)
            oCustomer.MeterAdmin = Range("AC" & i)
            oCustomer.Model = Range("AD" & i)
            oCustomer.Serial = Range("AE" & i)
            oCustomer.Contract = Range("AF" & i)
            oCustomer.MABase = Range("AG" & i)
            oCustomer.RentalBase = Range("AH" & i)
            oCustomer.Allowance = Range("AI" & i)
            oCustomer.MeterName = Range("AJ" & i)
            oCustomer.OverageRate = Range("AK" & i)
            oCustomer.BaseBillFrequency = Range("AL" & i)
            oCustomer.UsageBillFrequency = Range("AM" & i)
            oCustomer.ContactName = Range("AN" & i)
            'Range("AN" & i) not currently used was previous billed read
            oCustomer.CurrentRead = Range("AO" & i)
            oCustomer.GroupContract = Range("AP" & i)
            oCustomer.NumPeriods = Range("AQ" & i)
            
            'If you create a new column of data you must also add the column name to the clsCustomer class module
            
            
            'Do calculations to create NewPOP Start and End
            oCustomer.NewPOPStartDate = Format(DateAdd("d", 1, CDate(oCustomer.CurrentPOPEndDate)), "mm/dd/yyyy")
            Select Case oCustomer.BaseBillFrequency
            Case "Monthly"
                oCustomer.NewPOPEndDate = Format(DateAdd("d", -1, DateAdd("m", CDec(oCustomer.NumPeriods), CDate(oCustomer.NewPOPStartDate))), "mm/dd/yyyy")
            Case "Quarterly"
                oCustomer.NewPOPEndDate = Format(DateAdd("d", -1, DateAdd("m", (CDec(oCustomer.NumPeriods) * 3), CDate(oCustomer.NewPOPStartDate))), "mm/dd/yyyy")
            Case "Semi-Annually"
                oCustomer.NewPOPEndDate = Format(DateAdd("d", -1, DateAdd("m", (CDec(oCustomer.NumPeriods) * 6), CDate(oCustomer.NewPOPStartDate))), "mm/dd/yyyy")
            Case "Annually"
                oCustomer.NewPOPEndDate = Format(DateAdd("d", -1, DateAdd("m", (CDec(oCustomer.NumPeriods) * 12), CDate(oCustomer.NewPOPStartDate))), "mm/dd/yyyy")
            Case Else
                oCustomer.NewPOPEndDate = Format(DateAdd("d", -1, DateAdd("m", 12, CDate(oCustomer.NewPOPStartDate))), "mm/dd/yyyy")
            End Select

        coll.Add oCustomer
        
            
    Next i
    
    'create list of contracts without duplicates
    
    Dim contractColl As New Collection
    Dim line As Variant
    For Each line In coll
        On Error Resume Next
        contractColl.Add line.Contract, CStr(line.Contract)
    Next line
    
    
    'loop through collection if contract matches
 
    Dim l As Variant
    Dim checkContract As String
    
    For Each contractNum In contractColl
        checkContract = contractNum
                'Create quote form
                Debug.Print ("Creating Quote")
                Sheets("New Quote Form").Copy Before:=Sheets(1)
                ActiveSheet.Name = contractNum & "QuoteFormContract"
                'loop through each serial/line in contract and fill out quote form
                Call createQuote(coll, contractColl, checkContract)
                'Count serials to determine how many appendixes there will be
                numSerials = 0
                    For Each Serial In coll
                        If Serial.Contract = checkContract Then
                        numSerials = numSerials + 1
                        End If
                    Next Serial
                'Create CC Form
                Debug.Print ("Creating CCForm")
                Sheets("New CC Form").Copy Before:=Sheets(1)
                ActiveSheet.Name = contractNum & "QuoteCCForm"
                Call createCCForm(coll, contractColl, checkContract)
        
                If answer = vbYes Then
                    Dim arrayOfSheets As Variant
                    arrayOfSheets = Array(contractNum & "QuoteCCForm", contractNum & "QuoteFormContract")
                    numAppendix = 0
                    If numSerials > 20 Then
                       numAppendix = (((numSerials - 70) / 50) + 1)
                    End If
                    
                    
                    
                    
                    For n = 0 To numAppendix - 1
                        ReDim Preserve arrayOfSheets(UBound(arrayOfSheets) + 1)
                        arrayOfSheets(UBound(arrayOfSheets)) = contractNum & "Appendix" & n + 1
                        Debug.Print (n)
                    Next n
     
                    Call printPDF(contractNum, myfolder, arrayOfSheets)
                End If
                
        Next contractNum
    
   Application.ScreenUpdating = True


End Sub

Public Function createQuote(coll, contractColl, checkContract)
    
    Dim i As Integer
    i = 0

    For Each Serial In coll
        
        If Serial.Contract = checkContract Then
              
            'Run this part once per contract
            If i = 0 Then
                'Bill To Information
                Range("C9").Value = Serial.BillToCustomerName
                Range("C10").Value = Serial.BillToAddress
                Range("C11").Value = Serial.BillToTown
                Range("C12").Value = Serial.BillToState
                Range("F12").Value = Serial.BillToZipCode
                Range("C13").Value = Serial.BillToContactName
                Range("C14").Value = Serial.BillToPhoneNumber
                Range("C15").Value = Serial.BillToFaxNumber
                Range("C16").Value = Serial.BillToEmail
                
                
                'POP Info
                Range("C17").Value = Serial.NewPOPStartDate
                Range("E17").Value = Serial.NewPOPEndDate
                Range("F19").Value = Serial.BaseBillFrequency
                Range("G19").Value = Serial.GroupContract
                
                'Ship To Information
                Range("I9").Value = Serial.ShipToCustomerName
                Range("I10").Value = Serial.ShipToAddress
                Range("I11").Value = Serial.ShipToTown
                Range("I12").Value = Serial.ShipToState
                Range("K12").Value = Serial.ShipToZipCode
                Range("I13").Value = Serial.ShipToContactName
                Range("I14").Value = Serial.ShipToPhoneNumber
                Range("I15").Value = Serial.ShipToFaxNumber
                Range("I16").Value = Serial.ShipToEmail
                Range("H19").Value = Serial.UsageBillFrequency
                
                'Quote Info
                Range("J17").Value = Serial.QuoteInfoQuoteNumber
                Range("H64").Value = Serial.QuoteInfoEmailAddress
                Range("D62").Value = Serial.ContractAwardNumber
            
            End If
            
            If i < 20 Then
                'Line information
                Range("B" & (22 + i)).Value = Serial.Model
                Range("C" & (22 + i)).Value = Serial.CurrentRead
                Range("D" & (22 + i)).Value = Serial.Serial
                Range("F" & (22 + i)).Value = Serial.Contract
                Range("G" & (22 + i)).Value = Serial.MABase
                Range("I" & (22 + i)).Value = Serial.RentalBase
                Range("J" & (22 + i)).Value = Serial.MeterName
                Range("K" & (22 + i)).Value = Serial.Allowance
                Range("L" & (22 + i)).Value = Serial.OverageRate
            End If
            
            If i > 19 And i < 70 Then
                If i = 20 Then
                    'create new page
                    Sheets("Quote Overflow Page").Copy After:=Sheets(checkContract & "QuoteFormContract")
                    ActiveSheet.Name = checkContract & "Appendix1"
                End If
                

                
                'Line information
                Range("A" & (i - 18)).Value = Serial.Model
                Range("B" & (i - 18)).Value = Serial.CurrentRead
                Range("C" & (i - 18)).Value = Serial.Serial
                Range("E" & (i - 18)).Value = Serial.Contract
                Range("F" & (i - 18)).Value = Serial.MABase
                Range("H" & (i - 18)).Value = Serial.RentalBase
                Range("I" & (i - 18)).Value = Serial.MeterName
                Range("J" & (i - 18)).Value = Serial.Allowance
                Range("K" & (i - 18)).Value = Serial.OverageRate
                
            End If
            
            
                If i > 69 Then
                    If ((i - 70) Mod 50) = 0 Then '
                        j = 2
                        Debug.Print ("appendix " & ((i - 70) / 50) + 2)
                        'create new page
                        Sheets("Quote Overflow Page").Copy After:=Sheets(checkContract & "Appendix" & ((i - 70) / 50) + 1)
                        ActiveSheet.Name = checkContract & "Appendix" & ((i - 70) / 50) + 2
                    End If

                    Range("A" & j).Value = Serial.Model
                    Range("B" & j).Value = Serial.CurrentRead
                    Range("C" & j).Value = Serial.Serial
                    Range("E" & j).Value = Serial.Contract
                    Range("F" & j).Value = Serial.MABase
                    Range("H" & j).Value = Serial.RentalBase
                    Range("I" & j).Value = Serial.MeterName
                    Range("J" & j).Value = Serial.Allowance
                    Range("K" & j).Value = Serial.OverageRate
                    j = j + 1
                End If

           
            'calculate total for MA Base and Rent
            currMATotal = currMATotal + CDec(Serial.MABase)
            currRentTotal = currRentTotal + CDec(Serial.RentalBase)
           
            i = i + 1
            
        End If
            
        
    Next Serial
            
            'print totals to first page
            Worksheets(checkContract & "QuoteFormContract").Activate
            Range("G42").Value = currMATotal
            Range("I42").Value = currRentTotal


    
End Function

Public Function createCCForm(coll, contractColl, checkContract)
    Dim i As Integer
    i = 0
    Dim currTotal As Double
    currTotal = 0
    For Each Serial In coll
        If Serial.Contract = checkContract Then
            
            If i = 0 Then
            
                Range("G3").Value = Serial.ContractAwardNumber
            
                'Bill To Info
                Range("F12").Value = Serial.BillToCustomerName
                Range("F13").Value = Serial.BillToAddress
                Range("F14").Value = Serial.BillToTown
                Range("F15").Value = Serial.BillToState
                Range("F16").Value = Serial.BillToZipCode
                'Ship To Info
                Range("G12").Value = Serial.ShipToCustomerName
                Range("G13").Value = Serial.ShipToAddress
                Range("G14").Value = Serial.ShipToTown
                Range("G15").Value = Serial.ShipToState
                Range("G16").Value = Serial.ShipToZipCode
                

                
            End If

                                        'calculate totals
                Range("E24").Value = Serial.NumPeriods
                currTotal = currTotal + CDec(Serial.MABase) + CDec(Serial.RentalBase)
                Range("F24").Value = currTotal
                Range("G24").Value = Range("F24").Value * Range("E24").Value
                Range("G32").Value = Range("G24").Value
                Range("G35").Value = Range("G24").Value

          i = i + 1
        End If
    Next Serial
End Function

Public Function printPDF(contractNum, myfolder, arrayOfSheets)


Dim str As String
Dim myfile As String

Sheets(arrayOfSheets).Select

myfile = contractNum
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
myfolder & myfile _
, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False






End Function

