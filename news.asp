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
    当前位置:<a href="index.asp">首页</a><%=GetNavPath("newsclass",ClassID)%>
    </div>
    <div class="container clearfix" style="margin-bottom: 20px;">
        <div class="con-r">
            <h2><span><%=GetSubNavName("newsClass",ClassID)%></span></h2>
            <%
				ID=ReplaceBadChar(Trim(Request("ID")))
				if ID="" or IsNumeric(ID)=false then
					Sqlaa = "select top 1 * from newsinfo where ClassID="&ClassID&" order by id desc "
				else
					Sqlaa = "Select * From newsinfo where ID="&ID&""
				end if
				Set Rsa1=Server.CreateObject("Adodb.RecordSet")
				Rsa1.Open Sqlaa,Conn,1,3
				Rsa1("newsClick") = Rsa1("newsClick") +1
				Rsa1.Update
				if not (Rsa1.Eof or Rsa1.Bof) then
			%>
            <div class="container clearfix" style="margin-bottom: 20px;">
                <div class="news-cnt">
                    <h1 class="title"><%=Rsa1("newstitle")%></h1>
                    <h2 id="share">发布时间：<%=Right(Year(Rsa1("PostTime")),4)&"-"&Right("0"&Month(Rsa1("PostTime")),2)&"-"&Right("0"&Day(Rsa1("PostTime")),2)%> 浏览次数：<%=Rsa1("newsClick")%>次</h2>
                    <div class="cnt"><%=Rsa1("newsContent")%></div>
                </div>
    		</div>
            <%
				end if
				Rsa1.Close:Set Rsa1=Nothing
			%>
        </div>
        <div class="con-l">
            <div class="column">
                <h2><%=GetSubNavName("newsClass",ClassID)%></h2>
                <ul>
                	<%
						Set Rs=Server.CreateObject("Adodb.RecordSet")
						Sql = "Select * From newsInfo Where ClassID In ("&ClassID&GetAllChild("newsClass",ClassID)&") order by newsOrder desc"
						Rs.Open Sql,Conn,1,1
						do while not Rs.eof
					%>
                	<li><a href="news.asp?ID=<%=Rs("ID")%>&ClassID=<%=Rs("ClassID")%>" title="<%=Rs("NewsTitle")%>"><%=GetNewsTitle(Rs("ID"),10)%></a></li>
                    <%
						Rs.MoveNext
						Loop
						Rs.Close:Set Rs=Nothing
					%>
                </ul>
            </div>
            <!--#include file="left.asp" -->
        </div>
    </div>
</div>
<!--#include file="bottom.asp" -->
</body>
</html>