# Email Automation

## Basic Overview
Using data from excel documents submitted to building property management, builds a message in outlook to send to new employees with the proper details including the subject, personalized message body, and attachment. 

## Code Breakdown
The main procedure contains class and variable declarations. We establish the first row in which data will be read from.

    Sub main()

    Dim FileName As String
    Dim wb As Workbook
    Dim Ws As Worksheet
    Dim NameVar As CNameType
    Set NameVar = New CNameType
    Dim IncVar As CIncType
    Set IncVar = New CIncType
    FileName = InputBox("Enter Move Sheet File Name")
    Set wb = Workbooks.Open("W:\IOS\IA\CCU\WO SUBMITTED 2018 JULY THRU DECEMBER\" & FileName & ".xlsx")
    Set Ws = wb.Worksheets("Sheet1")
    Dim IsNewEmp As Boolean
    IncVar.inc = 9
    IncVar.ph_letter = "f" + CStr(IncVar.inc)
    IncVar.phtype_letter = "e" + CStr(IncVar.inc)
    IncVar.name_letter = "d" + CStr(IncVar.inc)
    NameVar.ph = Ws.Range(IncVar.ph_letter).Value
    NameVar.phtype = Ws.Range(IncVar.phtype_letter).Value
    NameVar.Name = Ws.Range(IncVar.name_letter).Value


Next, a while loop tests for a blank string value in the name cell. If the string length doesn't equal zero, we proceed to split the name into an array and assign these values to extract the first and last name, and whether it is indicated that this is a "new" employee.


    While (Not (Len(NameVar.Name) = 0)) 'if cell is not blank, continue
      Call SplitName(NameVar) 'splits name string into parts
      IsNewEmp = IsNew(NameVar)
    If IsNewEmp = True Then
      Call CreateMessage(NameVar) 'build message body, attachment, and TO field
    End If
    Set NameVar = New CNameType 'create new object for next loop
    Call Increment(NameVar, IncVar, Ws) 'increment cells to pull info from move sheet
         Wend
     MsgBox "Message(s) sent successfully"
    End Sub


If there's "new" keyword in the string, set flag, then call create message subroutine. UBound function counts the number of words separated by a space.


###### SplitName

    Sub SplitName(NameVar As CNameType)

    Dim TempResult As Variant
    Dim WordsNumber As Integer
    Dim strCnt As Integer

    TempResult = Split(NameVar.Name, " ")
    strCnt = UBound(TempResult) + 1

    If strCnt = 3 Then
    NameVar.first = TempResult(0)
    NameVar.last = TempResult(1)
    NameVar.newemp = TempResult(2)

    ElseIf strCnt = 2 Then
    NameVar.first = TempResult(0)
    NameVar.last = TempResult(1)

    End If

    End Sub


 
CreateMessage subroutine creates outlook object and completes email recipient and address, body details, and adds attachment.


###### CreateMessage


    Sub CreateMessage(NameVar As CNameType)

     Dim OutApp As Object
     Dim OutMail As Object
     Set OutApp = CreateObject("Outlook.Application")
     Set OutMail = OutApp.CreateItem(0)
     Fname = "C:\Users\ctrueman\Documents\NavigationMap.pdf" 'File path/name of the Navigation Map
  
      On Error Resume Next
      With OutMail
        .To = NameVar.first & "." & NameVar.last & "@arb.ca.gov"
        .CC = ""
        .BCC = ""
        .Subject = "AT&T Phone Setup"
        .Body = BodySelect(OutMail, NameVar)
        .Attachments.Add Fname
        .Send   'or use .Display
        '.Display
      End With
    
      Set OutMail = Nothing
      Set OutApp = Nothing
    End Sub
    
    

BodySelect builds email message body. The If statement will breakdown one of two options, "D" for digital or "A" for analog, and a template will be used and returned to the CreateMessage subroutine.

###### BodySelect

    
    Function BodySelect(OutMail As Object, NameVar As CNameType)
      NameVar.phtype = Trim(NameVar.phtype)
      If NameVar.phtype = "D" Or NameVar.phtype = "d" Then
      
        OutMail.Body = "Hello " & NameVar.first & "," & Chr(10) & Chr(10) & "Your new number is 916-" & NameVar.ph & ". To access your           voicemail box and begin your setup dial 327-1944. The pin number has been set to your seven digit phone number " & NameVar.ph &         ". Once logged in please record a new greeting and your name as the previous user has one set.  You will also be able to set a           new pin number if you choose to. After the setup is complete if you have a voicemail waiting your phone will show an arrow               pointing towards key 8 labeled 'message waiting'. Press that key and you can log in from there.  Attached is a PDF chart to help         with navigating the ATT messaging system during the setup and for future reference." & Chr(10) & Chr(10) & "Calvin Trueman" &           Chr(10) & "IT Administration" & Chr(10) & _
        "Air Resources Board | Office of Information Services" & Chr(10) & "(916) 322-2908 desk | (916) 327-0640 fax" & Chr(10) &               "Calvin.Trueman@arb.ca.gov"
        
      ElseIf NameVar.phtype = "A" Or NameVar.phtype = "a" Then
 
        OutMail.Body = "Hello " & NameVar.first & "," & Chr(10) & Chr(10) & "Your new number is 916-" & NameVar.ph & ". To access your           voicemail box and begin your set up dial 327-1944. The pin number has been set to your seven digit phone number " & NameVar.ph &         ". Once logged in please record a new greeting and your name as the previous use has one set. You will also be able to set a             new pin number if you choose to. After the setup is complete if you have a voicemail waiting dial 327-1944 and log in from               there. If there are old voicemails in the queue please contact your supervisor on how to handle the message. Attached is a PDF           chart to help with navigating the ATT messaging system during the setup and for future reference." & Chr(10) & Chr(10) &                 "Calvin Trueman" & Chr(10) & "IT Administration" & Chr(10) & _
        "Air Resources Board | Office of Information Services" & Chr(10) & "(916) 322-2908 desk | (916) 327-0640 fax" & Chr(10) &               "Calvin.Trueman@arb.ca.gov"
        
      Else
        MsgBox "Incorrect input for A/D"
        
        End If   
      Set BodySelect = OutMail
     End Function
     
     
     
     
Increment subroutine increments the cell column number and concatenates it with the row letter. We use this to assign the next cell value to the string variables.
     
###### Increment


     
    Sub Increment(NameVar As CNameType, IncVar As CIncType, Ws As Worksheet)

    IncVar.inc = IncVar.inc + 1

    IncVar.ph_letter = "f" + CStr(IncVar.inc)
    IncVar.phtype_letter = "e" + CStr(IncVar.inc)
    IncVar.name_letter = "d" + CStr(IncVar.inc)
 
  
    NameVar.ph = Ws.Range(IncVar.ph_letter).Value
    NameVar.phtype = Ws.Range(IncVar.phtype_letter).Value
    NameVar.Name = Ws.Range(IncVar.name_letter).Value
 
    End Sub
    
    
    
IsNew function returns to a flag in main, determines whether or not to proceed to create message. If not a new employee returns false to local variable and doesn't create message.   

    
##### IsNew


    Function IsNew(ByVal NameVar) As Boolean

    If NameVar.newemp = "(New)" Or NameVar.newemp = "(new)" Or NameVar.newemp = "new" Or NameVar.newemp = "New" Then
     IsNew = True
    Else
     IsNew = False
    End If

    End Function
    
    
    
