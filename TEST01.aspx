<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- saved from url=(0052)http://aymkdn.github.io/ExcelPlus/demo/example2.html -->
<%
response.Buffer = true
session.Codepage =65001
response.Charset = "utf-8"
%>

<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import NameSpace = "System.IO"%>
<%@ Import NameSpace = "System.Web.Configuration" %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>


<br>
第1題
<hr>

 <form action="TEST01.aspx" name="TheForm" id="TheForm">
                           
   <input name="SRC" id="SRC" type="text" style="width:50%;" autofocus class="form-control" placeholder="請輸入字串" value="flipped class room is important">
   <button type="submit">送出</button>
                         
</form>
<%
if Request("SRC") <> "" then

Dim SRC as string = Request("SRC")
 
 Dim   strtxt = Split(SRC," ")
    For i = 0 To UBound(strtxt)
        response.Write (""& StrReverse(strtxt(i)) &" ")
    Next

end if 
%>
<br><br>

<br>
第2題
<hr>
<form action="TEST01.aspx" name="TheForm2" id="TheForm2">
                           
   <input name="SRC2" id="SRC2" type="text" style="width:50%;" autofocus class="form-control" placeholder="請輸入數字" value="15">
   <button type="submit">送出</button>
                         
</form>

<%

if Request("SRC2") <> "" then

Dim SRC2 as Integer = Request("SRC2")
Dim txt as string = ""
Dim i as Integer 

  For i = 1 To SRC2
	if i MOD 15 = 0 then
	    txt = ""& txt &" "& i &""
	else if i MOD 3 = 0 then
	    response.Write (" ")
	else if i MOD 5 = 0 then	
	    response.Write (" ")
	else
		txt = ""& txt &" "& i &""
	end if 
  Next
  
  Dim strtxt2 = Split(txt," ")
	response.Write (UBound(strtxt2))
	
end if

%>



<br><br>

<br>
第3題
<hr>
先選擇標籤為鉛筆或原子筆的袋子，如取出鉛筆，則標籤不是鉛筆的袋子內就是鉛筆，如取出原子筆，則標籤不是原子筆的袋子內則是原子筆。

<br><br>

<br>
第4題
<hr>
原本三個人拿出出300元，，但是套餐特價750元，照道理要各退50元給這三位共150元<br>
但實際上服務生退90元，等三人總付810元，扣掉服務生60元，<br>
所以810-60=750，此為實際上三人應付出的價格。


</body>
</html>

