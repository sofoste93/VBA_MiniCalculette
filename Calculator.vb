Private Sub addition_Click()

Dim num_1 As Variant, num_2 As Variant, result As Variant

If (firstNum.value = "" Or secondNum.value = "") Then

    MsgBox ("Invalid Format! Please check your input")

Else
   ' convert input to double
    num_1 = CDbl(firstNum.value)
    num_2 = CDbl(secondNum.value)
    
'If (num_1 = "" Or num_2 = "") Then
'    MsgBox ("PLEASE CHECK YOUR INPUT"), vbCritical + vbOKCancel, "FORMAT ERROR"
'Else
    result = num_1 + num_2
    
    Display_Result.value = result

End If


'MsgBox ("Addition result: " & result) ' do not forget to set a data type

End Sub

Private Sub closeBtn_Click()
    Unload Me ' this will close the app
End Sub

Private Sub clrBtn_Click()
' clear content
    Display_Result.value = ""
    firstNum.value = ""
    secondNum.value = ""
End Sub

Private Sub division_Click()

Dim num_1 As Variant, num_2 As Variant, result As Variant

If (firstNum.value = "" Or secondNum.value = "") Then

    MsgBox ("Invalid Format! Please check your input")

Else
   ' convert input to double
    num_1 = CDbl(firstNum.value)
    num_2 = CDbl(secondNum.value)
    
'If (num_1 = "" Or num_2 = "") Then
'    MsgBox ("PLEASE CHECK YOUR INPUT"), vbCritical + vbOKCancel, "FORMAT ERROR"
'Else
    result = num_1 / num_2
    
    Display_Result.value = result

End If

'MsgBox ("Addition result: " & result)
End Sub

Private Sub eight_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub five_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub four_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub Label1_Click()

End Sub

Private Sub multiplication_Click()

Dim num_1 As Variant, num_2 As Variant, result As Variant

If (firstNum.value = "" Or secondNum.value = "") Then

    MsgBox ("Invalid Format! Please check your input")

Else
   ' convert input to double
    num_1 = CDbl(firstNum.value)
    num_2 = CDbl(secondNum.value)
    
'If (num_1 = "" Or num_2 = "") Then
'    MsgBox ("PLEASE CHECK YOUR INPUT"), vbCritical + vbOKCancel, "FORMAT ERROR"
'Else
    result = num_1 * num_2
    
    Display_Result.value = result

End If

End Sub

Private Sub nine_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub one_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub seven_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub six_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub square_Click()
'Dim value
'value = CDbl(Display_Result.value) ' convert to double

'Sqr (value)
Dim num_1 As Variant, num_2 As Variant, result As Variant

'result = CDbl(Display_Result.value)

'Sqr (result)
Display_Result.value = "unavailable yet"

End Sub

Private Sub subtraction_Click()

Dim num_1 As Variant, num_2 As Variant, result As Variant

If (firstNum.value = "" Or secondNum.value = "") Then

    MsgBox ("Invalid Format! Please check your input")

Else
   ' convert input to double
    num_1 = CDbl(firstNum.value)
    num_2 = CDbl(secondNum.value)
    
'If (num_1 = "" Or num_2 = "") Then
'    MsgBox ("PLEASE CHECK YOUR INPUT"), vbCritical + vbOKCancel, "FORMAT ERROR"
'Else
    result = num_1 - num_2
    
    Display_Result.value = result

End If

' even if we don't specify a datatype, thi will work for subtraction

End Sub

Private Sub moduloBtn_Click()
'Der Ausdruck a Mod b entspricht einer der folgenden Formeln:
'
'a -(b * (a \ b))
'
'a -(b * Fix(a / b))
Dim num_1 As Variant, num_2 As Variant, result As Variant

If (firstNum.value = "" Or secondNum.value = "") Then

    MsgBox ("Invalid Format! Please check your input")

Else
   ' convert input to double
    num_1 = CDbl(firstNum.value)
    num_2 = CDbl(secondNum.value)
    
'If (num_1 = "" Or num_2 = "") Then
'    MsgBox ("PLEASE CHECK YOUR INPUT"), vbCritical + vbOKCancel, "FORMAT ERROR"
'Else
    result = num_1 Mod num_2
    
    Display_Result.value = result

End If
    
End Sub

Private Sub three_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub two_Click()
Display_Result.value = "unavailable yet"
End Sub

Private Sub UserForm_Click()
    MsgBox ("Thanks for using the program!")
End Sub

Private Sub zero_Click()
Display_Result.value = "unavailable yet"
End Sub
