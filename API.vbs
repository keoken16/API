
strFileLocation = Wscript.Arguments(0)
strTable = Wscript.Arguments(1)


 Dim oFS    : Set oFS = CreateObject("Scripting.FileSystemObject")
    Dim sFSpec : sFSpec  = oFS.GetAbsolutePathName(strFileLocation)
  Dim oXML   : Set oXML = CreateObject("Msxml2.DOMDocument.6.0")
  oXML.load sFSpec
  If 0 = oXML.parseError Then
     recursiveTraversalAtt oXML.documentElement, 0, strType, strInsert
  Else
     WScript.Echo objMSXML.parseError.reason
  End If


  
Sub recursiveTraversalAtt(oElm, nIndent, strType, strInsert)
	if oElm.getAttribute("name") = "insert" then
		strType = "insert"
	end if

	if len(strType) > 0 then
		if oElm.getAttribute("maxOccurs") = 1 then
			if strType = "insert" then
				if len(strInsert) > 0 then 
					strInsert = strInsert & "<SN>" & oElm.getAttribute("name")
				else
				strInsert = oElm.getAttribute("name")
				end if
			end if
		'WScript.Echo Space(nIndent), oElm.getAttribute("name") & " " & strType
		end if
	end if
	If 0 < oElm.childNodes.length Then
			For Each oChild In oElm.childNodes
				recursiveTraversalAtt oChild,0,strType,strInsert
			Next
			strType = ""
	End If

End Sub

str1 = vbtab
str2 = vbtab & vbtab
str3 = vbtab & vbtab & vbtab
str4 = vbtab & vbtab & vbtab & vbtab
str5 = vbtab & vbtab & vbtab & vbtab & vbtab
str6 = vbtab & vbtab & vbtab & vbtab & vbtab & vbtab

strEvn = vbtab & vbtab & "envelope = envelope & " & chr(34)
strSoap = "SOAP"
strSoapEnd = "]]"

strAction = "insert"
strFunction = strTable & strAction



strXMLUpdateHead = "'Update Requires sys_id" & VbCrLf	
strXmlHead = "function " & strFunction & "(ftnInstance,ftnUIDPWD,ftnData)" & VbCrLf
strXmlHead = strXmlHead &  str1 & "Set wsReq = CreateObject(" & chr(34) & "Microsoft.XMLHTTP" & chr(34) & ")" & VbCrLf
strXmlHead = strXmlHead &  str1 & "url = " & chr(34) & "https://" & chr(34) & " & ftnInstance & " & chr(34) &  ".service-now.com/" & strTable & ".do?SOAP" & chr(34) & VbCrLf
strXmlHead = strXmlHead &  str1 & "wsReq.open " & chr(34) & "POST" & chr(34) & ", url, false" & VbCrLf
strXmlHead = strXmlHead &  str1 & "wsReq.setRequestHeader " & chr(34) & "Content-Type" & chr(34) & ", " & chr(34) & "text/xml;charset=UTF-8" & chr(34) & "" & VbCrLf
strXmlHead = strXmlHead &  str1 & "wsReq.setRequestHeader " & chr(34) & "SOAPAction" & chr(34) & ", _" & VbCrLf
strXmlHead = strXmlHead &  str1 & "                       " & chr(34) & "http://www.service-now.com/" & strTable & "/" & strAction & chr(34) & "" & VbCrLf
strXmlHead = strXmlHead &  str1 & "" & VbCrLf
strXmlHead = strXmlHead &  str1 & "' Password/Username needs to be in base64 set ftnUIDPWD" & VbCrLf
strXmlHead = strXmlHead &  str1 & "' Use the following to create http://www.base64encode.org/" & VbCrLf
strXmlHead = strXmlHead &  str1 & "' paste in username:password  ex admin:admin would be YWRtaW46YWRtaW4=" & VbCrLf
strXmlHead = strXmlHead &  str1 & "wsReq.setRequestHeader " & chr(34) & "Authorization" & chr(34) & ", _" & VbCrLf
strXmlHead = strXmlHead &  str1 & "                       " & chr(34) & "Basic " & chr(34) & " & ftnUIDPWD" & VbCrLf
strXmlHead = strXmlHead &  str1 & "" & VbCrLf
strXmlHead = strXmlHead &  str1 & "envelope = " & chr(34) & "<soapenv:Envelope xmlns:soapenv=" & chr(34) & "" & chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & chr(34) & "" & chr(34) & " xmlns:sys=" & chr(34) & "" & chr(34) & "http://www.service-now.com/" & strTable & chr(34) & "" & chr(34) & ">" & chr(34) & " & VbCrLf & _" & VbCrLf
strXmlHead = strXmlHead &  str2 & "   " & chr(34) & "<soapenv:Header/>" & chr(34) & " & VbCrLf & _ " & VbCrLf
strXmlHead = strXmlHead &  str3 & "   " & chr(34) & "<soapenv:Body>" & chr(34) & " & VbCrLf & _ " & VbCrLf
strXmlHead = strXmlHead &  str4 & "	" & chr(34) & "<sys:" & strAction & ">" & chr(34) & " & VbCrLf" & VbCrLf


	strInsertArr = split(strInsert,"<SN>")
	for each x in strInsertArr
	strXmlBody = strXmlBody &  str5 & vbtab & "if len(ftnFilterData(arrData," & chr(34) & "SOAP" & x & chr(34) & ")) > 0 then" & VbCrLf
	strXmlBody = strXmlBody &  str5 & strEvn & "<" & x & ">" & chr(34) & " & ftnFilterData(arrData," & chr(34) & "SOAP" & x & chr(34) & ") & " & chr(34) & "</" & x & ">" & chr(34) & " & VbCrLf" & VbCrLf
	strXmlBody = strXmlBody &  str5 & vbtab & "end if" & VbCrLf
	next
	
	strgetKeys = "sys_class_name<SN>sys_created_by<SN>sys_created_on<SN>sys_domain<SN>sys_mod_count<SN>sys_updated_by<SN>sys_updated_on<SN>__use_view<SN>__encoded_query<SN>__limit<SN>__first_row<SN>__last_row"
	
	strgetKeysArr = split(strgetKeys,"<SN>")
	for each x in strgetKeysArr
	strXmlBodyUpdateKeys = strXmlBodyUpdateKeys &  str5 & vbtab & "if len(ftnFilterData(arrData," & chr(34) & "SOAP" & x & chr(34) & ")) > 0 then" & VbCrLf
	strXmlBodyUpdateKeys = strXmlBodyUpdateKeys &  str5 & strEvn & "<" & x & ">" & chr(34) & " & ftnFilterData(arrData," & chr(34) & "SOAP" & x & chr(34) & ") & " & chr(34) & "</" & x & ">" & chr(34) & " & VbCrLf" & VbCrLf
	strXmlBodyUpdateKeys = strXmlBodyUpdateKeys &  str5 & vbtab & "end if" & VbCrLf
	next
	
	strXmlBodyUpdate = str5 & vbtab & "if len(SOAP" & "sys_id" & ") > 0 then" & VbCrLf
	strXmlBodyUpdate = str5 & strEvn & "<" & "sys_id" & ">" & chr(34) & " & ftnFilterData(arrData," & chr(34) & "SOAPsys_id"& chr(34) & ") & " & chr(34) & "</" & "sys_id" & ">" & chr(34) & " & VbCrLf" & VbCrLf
	'strXmlBodyUpdate = str5 & vbtab & "end if" & VbCrLf
	
strXmlTail = strXmlTail &  str4 & "envelope = envelope & " & chr(34) & "</sys:" & strAction & ">" & chr(34) & VbCrLf
strXmlTail = strXmlTail &  str3 & "envelope = envelope & " & chr(34) & "</soapenv:Body>" & chr(34) & VbCrLf
strXmlTail = strXmlTail &  str2 & "envelope = envelope & " & chr(34) & "</soapenv:Envelope>" & chr(34) & VbCrLf

strXmlTail = strXmlTail &  str1 & "wsReq.send envelope" & VbCrLf
strXmlTail = strXmlTail &  str1 & "" & VbCrLf
strXmlTail = strXmlTail &  str1 & "'wscript.echo Envelope" & VbCrLf
strXmlTail = strXmlTail &  str1 & strTable & "insert = wsReq.responsetext" & VbCrLf
strXmlTail = strXmlTail &  "end function"
strXmlTail = strXmlTail &  VbCrLf
strXmlTail = strXmlTail &  VbCrLf



strXmlInsert = strXmlHead & strXmlBody & strXmlTail
strXmlUpdate = strXMLUpdateHead & strXmlHead & strXmlBody & strXmlBodyUpdate & strXmlTail
strXmlUpdate = replace(strXmlUpdate,"insert","update")
strXmlgetKeys = strXmlHead & strXmlBody & strXmlBodyUpdateKeys & strXmlTail
strXmlgetKeys = replace(strXmlgetKeys,"insert","getKeys")

strFtnSplit = "" & VbCrLf
strFtnSplit = strFtnSplit & "function ftnFilterData(arrData,strVar)"  & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"arrData = Split(ftnData," & chr(34) & "<S>" & chr(34) & ")" & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"strFilter = Filter(arrData,strVar)" & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"for each x in strFilter"  & VbCrLf
strFtnSplit = strFtnSplit &  str2 &	""  & VbCrLf
strFtnSplit = strFtnSplit &  str2 &	"strValue = Split(x," & chr(34) & "<=>" & chr(34) & ")" & VbCrLf
strFtnSplit = strFtnSplit &  str2 &	"strRan = " & chr(34) & "True" & chr(34) & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"next"  & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"" & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"if strRan = " & chr(34) & "True" & chr(34) & "then" & VbCrLf
strFtnSplit = strFtnSplit &  str2 &	"ftnFilterData = strValue(1)" & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"else" & VbCrLf
strFtnSplit = strFtnSplit &  str2 &	"ftnFilterData = " & chr(34) & chr(34) & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"end if" & VbCrLf
strFtnSplit = strFtnSplit &  str1 &	"strRan = " & chr(34) & chr(34) & VbCrLf
strFtnSplit = strFtnSplit &	"end function" & VbCrLf

strlist = strXmlgetKeys & VbCrLf
strlist = strlist & strXmlUpdate & VbCrLf
strlist = strlist & strXmlInsert & VbCrLf
strlist = strlist & strFtnSplit

sname = strTable & ".vbs"



Set fso = CreateObject("Scripting.FileSystemObject")





If fso.FileExists(sname) Then
    'you delete if you find it'
    fso.DeleteFile sname, True
End If
'you always write it anyway.'
Set spoFile = fso.CreateTextFile(sname, True)
spoFile.WriteLine(strlist)
Set objFolderItem = Nothing
Set objFolder = Nothing
Set objApplication = Nothing
Set fso = Nothing
spoFile.Close



