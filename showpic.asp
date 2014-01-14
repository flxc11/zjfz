<!--#include file="Config/conn.asp" -->
<!--#include file="Include/Class_Function.asp" -->
<%
	ID=ReplaceBadChar(Trim(Request("ID")))
	if ID="" or IsNumeric(ID)=false then
	twscript("参数错误")
	end if
	ClassID=GetClassID(ID,"shopinfo")
%>
<%
	Set Rsaa=Server.CreateObject("Adodb.RecordSet")
	Sqlaa = "Select * From shopinfo where ID="&ID&""
	Rsaa.Open Sqlaa,Conn,1,1
	if not (Rsaa.Eof or Rsaa.Bof) then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="css/Global.css"/>
<link rel="stylesheet" type="text/css" href="stylus/style.css"/>
<link rel="stylesheet" type="text/css" href="css/tabs.css"/>
<script type="text/javascript" src="js/jquery180min.js"></script>
<script type="text/javascript" src="js/add.js"></script>
<script type="text/javascript" src="js/tab.js"></script>
<script type="text/javascript" src="js/msclass.js"></script>
<%=SiteKeysTitle("")%>
</head>

<body>
<div class="wrap">
    <!--#include file="top.asp" -->
    <div class="position">
    当前位置:<a href="index.asp">首页</a><%=GetNavPath("shopclass",ClassID)%>
    </div>
    <div class="container clearfix" style="margin-bottom: 20px;">
        <div class="news-cnt">
            <h1 class="title"><%=Rsaa("shopname")%></h1>
            <h2 class="share">发布时间：<%=Right(Year(Rsaa("PostTime")),4)&"-"&Right("0"&Month(Rsaa("PostTime")),2)&"-"&Right("0"&Day(Rsaa("PostTime")),2)%> 浏览次数：<%=Rsaa("ShopClick")%>次</h2>
            <div class="cnt"><%=Rsaa("ShopContent")%></div>
        </div>
    </div>
</div>
<%
				end if
				Rsaa.Close:Set Rsaa=Nothing
			%>
<!--#include file="bottom.asp" -->
</body>
</html>