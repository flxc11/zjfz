<!--#include file="Config/conn.asp" -->
<!--#include file="Include/Class_Function.asp" -->
<%
	ClassID=ReplaceBadChar(Trim(Request("ClassID")))
	if ClassID="" or Isnumeric(classid)=false then
	ClassID=1
	end if
	Set Rs2 = Server.CreateObject("Adodb.RecordSet")
		Sql2 = "select * from User_PageCategory where ID="&ClassID
		Rs2.open sql2,conn,1,1
		NavParent = Rs2("User_NavParent")
		Rs2.Close:Set Rs2=Nothing
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
    当前位置:<a href="index.asp">首页</a><%=GetPageNavPath("User_PageCategory",ClassID)%>
    </div>
    <div class="container clearfix" style="margin-bottom: 20px;">
        <div class="con-r">
            <h2><span><%=GetPageNavName("User_PageCategory",ClassID)%></span></h2>
            <div class="cnt"><%=GetPageContent1("SiteExplain",ClassID,"NavContent")%></div>
        </div>
        <div class="con-l">
            <div class="column">
                <h2><%=GetPageNavName("User_PageCategory",ClassID)%></h2>
                <ul>
                	<%
						Set Rsr = Server.CreateObject("Adodb.RecordSet")
						Sqlr = "select * from User_PageCategory where User_NavParent="&NavParent&" and User_NavLevel=2 order by User_NavOrder desc"
						Rsr.open Sqlr,Conn,1,1
						do while not Rsr.Eof
					%>
                	<li><a href="about.asp?ClassID=<%=Rsr("ID")%>"><%=Rsr("User_NavTitle")%></a></li>
                    <%
						Rsr.MoveNext
						Loop
						Rsr.Close:Set Rsr=Nothing
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