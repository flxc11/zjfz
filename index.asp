<!--#include file="Config/conn.asp" -->
<!--#include file="Include/Class_Function.asp" -->
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
<%=SiteKeysTitle("首页")%>
</head>

<body>
<div class="wrap">
    <!--#include file="top.asp" -->
    <div class="banner">
    <object id='FlashID' classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' width='960' height='300'><param value='flash/index.swf' name='movie' /><param value='high' name='quality' /><param name='wmode' value='Opaque'><param value='opaque' name='wmode' /><param value='6.0.65.0' name='swfversion' /><embed id='EmbedID' width='960' height='300' type='application/x-shockwave-flash' src='flash/index.swf'></embed></object>
    </div>
    <div class="container">
        <div class="xgzj">
            <div class="cnvp-tab-nav T clearfix">
                <span class="sp-xgzj">相关专家</span>
                <a href="javascript:void(0);">法律顾问</a>
                <a href="javascript:void(0);">医学专家</a>
                <a href="javascript:void(0);" class="index_tabshover">法医学专家</a>
            </div>
            <div class="cnvp-tab-panle cnvp-tabs-hide C">
                <div id="hottitle2" class="hot">
                    <ul id="ulid2">
                    	<%
							Set Rs=Server.CreateObject("Adodb.RecordSet")
							Sql = "select * from shopinfo where ClassID=4 order by shoporder desc"
							Rs.Open Sql,Conn,1,1
							i=0
							do while not Rs.eof and i<15
						%>
                        <li><a href="showpic.asp?ID=<%=Rs("ID")%>" target="_blank" title="<%=Rs("ShopName")%>"><img src="<%=Rs("ShopSPic")%>" alt=""/></a><a href="showpic.asp?ID=<%=Rs("ID")%>" target="_blank" title="<%=Rs("ShopName")%>"><%=Rs("ShopName")%><br><%=Rs("ShopPara")%></a></li>
                        <%
							i=i+1
							Rs.MoveNext
							loop
							Rs.Close:Set Rs=Nothing
						%>
                    </ul>
                </div>
            </div>
            <div class="cnvp-tab-panle cnvp-tabs-hide C">
                <div id="hottitle" class="hot">
                    <ul id="ulid">
                        <%
							Set Rs=Server.CreateObject("Adodb.RecordSet")
							Sql = "select * from shopinfo where ClassID=2 order by shoporder desc"
							Rs.Open Sql,Conn,1,1
							i=0
							do while not Rs.eof and i<15
						%>
                        <li><a href="showpic.asp?ID=<%=Rs("ID")%>" target="_blank" title="<%=Rs("ShopName")%>"><img src="<%=Rs("ShopSPic")%>" alt=""/></a><a href="showpic.asp?ID=<%=Rs("ID")%>" target="_blank" title="<%=Rs("ShopName")%>"><%=Rs("ShopName")%><br><%=Rs("ShopPara")%></a></li>
                        <%
							i=i+1
							Rs.MoveNext
							loop
							Rs.Close:Set Rs=Nothing
						%>
                    </ul>
                </div>
            </div>
            <div class="cnvp-tab-panle  C">
                <div id="hottitle1" class="hot">
                    <ul id="ulid1">
                        <%
							Set Rs=Server.CreateObject("Adodb.RecordSet")
							Sql = "select * from shopinfo where ClassID=1 order by shoporder desc"
							Rs.Open Sql,Conn,1,1
							i=0
							do while not Rs.eof and i<15
						%>
                        <li><a href="showpic.asp?ID=<%=Rs("ID")%>" target="_blank" title="<%=Rs("ShopName")%>"><img src="<%=Rs("ShopSPic")%>" alt=""/></a><a href="showpic.asp?ID=<%=Rs("ID")%>" target="_blank" title="<%=Rs("ShopName")%>"><%=Rs("ShopName")%><br><%=Rs("ShopPara")%></a></li>
                        <%
							i=i+1
							Rs.MoveNext
							loop
							Rs.Close:Set Rs=Nothing
						%>
                    </ul>
                </div>
            </div>
            <SCRIPT language=javascript>
            new Marquee(["hottitle","ulid"],2,2,958,210,20,0,0);
            new Marquee(["hottitle1","ulid1"],2,2,958,210,20,0,0);
            new Marquee(["hottitle2","ulid2"],2,2,958,210,20,0,0);
            </SCRIPT>
        </div>
    </div>
    <div class="container clearfix" style="margin-top: 25px;">
        <div class="r-1">
            <h2 class="clearfix h2">
                <span class="more"><a href="news.asp?ClassID=3">更多>></a></span>
                <span class="col-name"><i class="icon1"></i>相关动态</span>
            </h2>
            <ul class="ul">
                
                <%
					Set Rs=Server.CreateObject("Adodb.RecordSet")
					Sql = "select * from newsinfo where ClassID=3 order by newsorder desc"
					Rs.Open Sql,Conn,1,1
					i=0
					do while not Rs.eof and i<6
				%>
				<li><a href="news.asp?ID=<%=Rs("ID")%>" target="_blank" title="<%=Rs("NewsTitle")%>"><%=GetNewsTitle(Rs("ID"),15)%></a></li>
				<%
					i=i+1
					Rs.MoveNext
					loop
					Rs.Close:Set Rs=Nothing
				%>
            </ul>
        </div>
        <div class="m-1">
            <h2 class="clearfix h2">
                <span class="more"><a href="news.asp?ClassID=2">更多>></a></span>
                <span class="col-name"><i class="icon2"></i>相关咨询</span>
            </h2>
            <div class="div-2 clearfix">
                <span class="p1"><img src="images/p2.jpg" alt=""/></span>
                <ul class="u2">
                    <%
					Set Rs=Server.CreateObject("Adodb.RecordSet")
					Sql = "select * from newsinfo where ClassID=2 order by newsorder desc"
					Rs.Open Sql,Conn,1,1
					i=0
					do while not Rs.eof and i<6
				%>
				<li><a href="news.asp?ID=<%=Rs("ID")%>&ClassID=<%=Rs("ClassID")%>" target="_blank" title="<%=Rs("NewsTitle")%>"><%=GetNewsTitle(Rs("ID"),20)%></a></li>
				<%
					i=i+1
					Rs.MoveNext
					loop
					Rs.Close:Set Rs=Nothing
				%>
                </ul>
            </div>
        </div>
        <div class="l-1">
        	<h2><span class="more"><a href="news.asp?ClassID=1">更多>></a></span></h2>
        	<ul>
                <%
					Set Rs=Server.CreateObject("Adodb.RecordSet")
					Sql = "select * from newsinfo where ClassID=1 order by newsorder desc"
					Rs.Open Sql,Conn,1,1
					i=1
					do while not Rs.eof and i<6
				%>
				<li class="li<%=i%>"><a href="news.asp?ID=<%=Rs("ID")%>&ClassID=<%=Rs("ClassID")%>" target="_blank" title="<%=Rs("NewsTitle")%>"><%=GetNewsTitle(Rs("ID"),15)%></a></li>
				<%
					i=i+1
					Rs.MoveNext
					loop
					Rs.Close:Set Rs=Nothing
				%>
            </ul>
        </div>
    </div>
    <div class="container clearfix" style="margin-top: 25px;">
        <div class="r-1">
            <h2 class="clearfix h2">
                <span class="more"><a href="news.asp?ClassID=5">更多>></a></span>
                <span class="col-name"><i class="icon1"></i>政策法规</span>
            </h2>
            <ul class="ul">
                <%
					Set Rs=Server.CreateObject("Adodb.RecordSet")
					Sql = "select * from newsinfo where ClassID=5 order by newsorder desc"
					Rs.Open Sql,Conn,1,1
					i=0
					do while not Rs.eof and i<6
				%>
				<li><a href="news.asp?ID=<%=Rs("ID")%>&ClassID=<%=Rs("ClassID")%>" target="_blank" title="<%=Rs("NewsTitle")%>"><%=GetNewsTitle(Rs("ID"),15)%></a></li>
				<%
					i=i+1
					Rs.MoveNext
					loop
					Rs.Close:Set Rs=Nothing
				%>
            </ul>
        </div>
        <div class="m-1">
            <h2 class="clearfix h2">
                <span class="more"><a href="news.asp?ClassID=4">更多>></a></span>
                <span class="col-name"><i class="icon2"></i>专家论丛</span>
            </h2>
            <div class="div-2 clearfix">
                <span class="p1"><img src="images/p3.jpg" alt=""/></span>
                <ul class="u2">
                    <%
					Set Rs=Server.CreateObject("Adodb.RecordSet")
					Sql = "select * from newsinfo where ClassID=4 order by newsorder desc"
					Rs.Open Sql,Conn,1,1
					i=0
					do while not Rs.eof and i<6
				%>
				<li><a href="news.asp?ID=<%=Rs("ID")%>&ClassID=<%=Rs("ClassID")%>" target="_blank" title="<%=Rs("NewsTitle")%>"><%=GetNewsTitle(Rs("ID"),20)%></a></li>
				<%
					i=i+1
					Rs.MoveNext
					loop
					Rs.Close:Set Rs=Nothing
				%>
                </ul>
            </div>
        </div>
        <div class="l-1">
            <div class="contact">
                <h2 class="contact-h2">联系我们</h2>
                <blockquote class="bk">
                浙江省天平鉴定辅助技术研究院<br>
                地址：杭州市万塘路252号计量<br />
				      大厦605室<br>
                联系人：华潭跃  谢淑萍<br>
                邮编：310030<br>
                电话：0571-87358618<br>
                传真：0571-87358617
                </blockquote>
            </div>
        </div>
    </div>
    <div class="friend">
        <h2>友情链接</h2>
        <ul>
            <li><a href="http://www.moj.gov.cn/" target="_blank">中国司法部</a></li>
            <li><a href="http://www.chinacourt.org/index.shtml" target="_blank">中国法院网</a></li>
            <li><a href="http://www.zjsft.gov.cn/" target="_blank">浙江省司法厅</a></li>
            <li><a href="http://www.zjcourt.cn/" target="_blank">浙江省高级人民法院</a></li>
            <li><a href="http://www.hzcourt.cn/" target="_blank">杭州市中级人民法院</a></li>
            <li><a href="http://www.zy91.com/" target="_blank">浙江大学医学院附属第一医院</a></li>
            <li><a href="http://www.z2hospital.com/" target="_blank">浙江大学医学院附属第二医院</a></li>
        </ul>
    </div>
</div>
<!--#include file="bottom.asp" -->
</body>
</html>