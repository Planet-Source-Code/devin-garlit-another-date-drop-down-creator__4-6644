<div align="center">

## Another date drop down creator


</div>

### Description

This has probably been done a million times but this is another variant of how to create dropdowns for the month, day and year. It bases the days on the month and year, so the user won't select day 31 for february. I'm happy to say I wrote this in <a href='http://www.developerspad.com'>Developers Pad</a>.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Devin Garlit](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/devin-garlit.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/devin-garlit-another-date-drop-down-creator__4-6644/archive/master.zip)





### Source Code

```
<script language=javascript>
function fnSubmit(strPage)
{
 document.forms[0].action= strPage
 document.forms[0].submit()
}
</script>
<%
'here is the call
writedropdowns
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'writeDropDowns
'purpose - write three drop down boxes where the days available are based on month and year
'notes - this submits to the page each time a month or year is selected and recalculates the number
'      of days
'    -I use a javascript function to submit my form, this may seem a bit much
'      but it allows me to customize better, If i want I can throw other validation in it, or
'      a frame target
'by: Devin Garlit dgarlit@hotmail.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub writeDropDowns()
 Dim strSelfLink
 strSelfLink = request.servervariables("SCRIPT_NAME")
 response.Write "<form name=dates method=post>" & vbcrlf
 response.Write MonthDropDown("month1",False,request("month1"),strSelfLink) & " " & DayDropDown("day1", "",getDaysInMonth(request("month1"),request("year1")),request("day1")) & " " & YearDropDown("year1","","", request("year1"),strSelfLink) & vbcrlf
 response.Write "</form>"	& vbcrlf
End Sub
'MonthDropDown
'strName = name of drop down
'blnNum = 'If blnNUM Is True, Then show As numbers
'strSelected = the currenct selected month
'strSelfLink = link to current page
Function MonthDropDown(strName, blnNum, strSelected, strSelfLink)
 Dim strTemp, i, strSelectedString
 strTemp = "<select name='" & strName& "' onchange='javascript: fnSubmit(" & chr(34) & strSelfLink & chr(34) & ")'>" & vbcrlf
 strTemp = strTemp & "<option value='" & 0 & "'>" & "Month" & "</option>" & vbcrlf
 For i = 1 To 12
  If strSelected = CStr(i) Then
	strSelectedString = "Selected"
  Else
	strSelectedString = ""
  End If
  If blnNum Then
   strTemp = strTemp & "<option value='" & i & "' " & strSelectedString & " >" & i & "</option>" & vbcrlf
  Else
	strTemp = strTemp & "<option value='" & i & "' " & strSelectedString & " >" & MonthName(i) & "</option>" & vbcrlf
  End If
 Next
 strTemp = strTemp & "</select>" & vbcrlf
 MonthDropDown = strTemp
End Function
'YearDropDown
'strName = name of dropdown
'intStartYear = year to start options list
'intEndYear = year to end options list
'strSelected = the currenct selected year
'strSelfLink = link To currentpage
Function YearDropDown(strName, intStartYear, intEndYear, strSelected, strSelfLink)
 Dim strTemp, i, strSelectedString
 If intStartYear = "" Then
  intStartYear = Year(now())
 End If
 If intEndYear = "" Then
  intEndYear = Year(now()) + 9
 End If
 strTemp = "<select name='" & strName& "' onchange='javascript: fnSubmit(" & chr(34) & strSelfLink & chr(34) & ")'>" & vbcrlf
 strTemp = strTemp & "<option value='" & 0 & "'>" & "Year" & "</option>" & vbcrlf
 For i = intStartYear To intEndYear
  If strSelected = CStr(i) Then
   strSelectedString = "Selected"
  Else
   strSelectedString = ""
  End If
  strTemp = strTemp & "<option value='" & i & "' " & strSelectedString & " >" & i & "</option>" & vbcrlf
  Next
  strTemp = strTemp & "</select>" & vbcrlf
  YearDropDown = strTemp
End Function
'DayDropDown
'strName = name of drop down
'intStartDay = day to start with
'intEndDay = day to end with
'strSelected = current slected day
Function DayDropDown(strName, intStartDay, intEndDay, strSelected )
	Dim strTemp, i, strSelectedString
	If intStartDay = "" Then
	 intStartDay = 1
	End If
	If intEndDay = "" Then
	 intEndDay = getDaysInMonth(Month(now()),Year(now()))
	End If
	strTemp = "<select name='" & strName& "'>" & vbcrlf
	strTemp = strTemp & "<option value='" & 0 & "'>" & "Day" & "</option>" & vbcrlf
	For i = intStartDay To intEndDay
	 If strSelected = CStr(i) Then
	  strSelectedString = "Selected"
	 Else
	  strSelectedString = ""
	 End If
	 strTemp = strTemp & "<option value='" & i & "' " & strSelectedString & " >" & i & "</option>" & vbcrlf
	Next
	strTemp = strTemp & "</select>" & vbcrlf
	DayDropDown = strTemp
End Function
'getDaysInMonth
'strMonth = month as number
'strYear = year
Function getDaysInMonth(strMonth,strYear)
		Dim strDays
  Select Case CInt(strMonth)
    Case 1,3,5,7,8,10,12:
		strDays = 31
    Case 4,6,9,11:
	    strDays = 30
    Case 2:
		If ( (CInt(strYear) Mod 4 = 0 And CInt(strYear) Mod 100 <> 0) Or ( CInt(strYear) Mod 400 = 0) ) Then
		 strDays = 29
		Else
		 strDays = 28
		End If
    'Case Else:
  End Select
  getDaysInMonth = strDays
End Function
%>
```

