Attribute VB_Name = "MakeDateLookNice"
Public Function ProcessDate(ByVal DateString As String) As String
Dim TheYear As Integer
Dim TheMonth As Integer
Dim TheDay As Integer
Dim TheHour As Integer
'for leading 0 in mins
Dim TheMin As String


TheYear = ((DateString \ &H2000000) And &H1F) + 1980 'yay. Int div used otherwise it give slightly wrong value. _
value needs to be "shifted" then masked to get small deci val. h2000000 =2^25 i.e. move 25 to the right
'MsgBox PropFileYear
TheMonth = ((DateString \ &H200000) And &HF) 'month code. h200000 = 2^21 i.e. move 21 right
'MsgBox PropFileMonth
TheDay = ((DateString \ &H10000) And &H1F) 'day code. h10000 = 2^16
TheHour = ((DateString \ &H800) And &H1F) 'hrs code. h800 =2^11
TheMin = ((DateString \ &H20) And &H3F) 'mins code. h20 = 2^5


ProcessDate = Format(TheDay & "/" & TheMonth & "/" & TheYear & "  " & TheHour & ":" & TheMin, "dd/mm/yyyy hh:mm")

End Function
