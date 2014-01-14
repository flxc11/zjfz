<!--#include file="Config/conn.asp" -->
<!--#include file="Include/Class_Function.asp" -->
<%
	ClassID=ReplaceBadChar(Trim(Request("ClassID")))
	if ClassID="" or Isnumeric(classid)=false then
	ClassID=1
	end if
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
    当前位置:<a href="index.asp">首页</a> > 站内搜索
    </div>
    <div class="container clearfix" style="margin-bottom: 20px;">
        <div class="con-r">
            <h2><span>站内搜索</span></h2>
            <div class="news-list">
                <ul>
                    <%
						Keyword=Request("Keyword")
						Set Rs=Server.CreateObject("Adodb.RecordSet")
						Sql = "Select * From newsInfo Where newstitle like '%"&Keyword&"%' order by newsorder desc"
						Rs.Open Sql,Conn,1,1
						Dim Page
						Page=Request("Page")                            
						PageSize = 10                                    
						Rs.PageSize = PageSize               
						Total=Rs.RecordCount               
						PGNum=Rs.PageCount               
						If Page="" Or clng(Page)<1 Then Page=1               
						If Clng(Page) > PGNum Then Page=PGNum               
						If PGNum>0 Then Rs.AbsolutePage=Page                         
						i=0
						Do While Not Rs.Eof And i<Rs.PageSize
					%>
                    <li><span><%=Right(Year(Rs("PostTime")),4)&"-"&Right("0"&Month(Rs("PostTime")),2)&"-"&Right("0"&Day(Rs("PostTime")),2)%></span><a href="shownews.asp?ID=<%=Rs("ID")%>" target="_blank" title="<%=Rs("NewsTitle")%>"><%=Rs("NewsTitle")%></a></li>
                    <%
						i=i+1
						Rs.MoveNext
						Loop
					%>
                </ul>
            </div>
            <div class="NewsPage"><%=GetPage1("Where newstitle like '%"&Keyword&"%'","newsinfo",10,0)%></div>
        </div>
        <div class="con-l">
            <div class="column">
                <h2>站内搜索</h2>
            </div>
            <!--#include file="left.asp" -->
        </div>
    </div>
</div>
<!--#include file="bottom.asp" -->
</body>
</html>