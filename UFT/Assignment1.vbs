dim lownumber 
lownumber = 0
dim highnumber 
highnumber = 0
dim temp 
temp = 0
dim remainder 
remainder = 0
dim result 
result = 0
dim i 
i = 0
dim StartTime
dim EndTime
dim UserInput

StartTime = Timer()
UserInput = InputBox("Enter the Lowest Range: ") 
lownumber = CInt(UserInput)
UserInput = InputBox("Enter the Highest Range: ")
highnumber = CInt(UserInput)
For i = lownumber To highnumber Step 1 
    temp = i
    result = 0
    While (temp > 0)
        remainder = (temp Mod 10)
        result = result + (remainder * remainder * remainder)
        temp = temp \ 10
    Wend
    If (i = result) Then
        MsgBox(i)
    End If
Next

EndTime = Timer()
MsgBox("time taken - " & FormatNumber(EndTime - StartTime, 2))

