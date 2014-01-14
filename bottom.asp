<div class="footer">
    <div class="infooter">
        <div class="nav" style="margin: 0 auto;width: 1000px;border-top: none;border-bottom: 1px solid #285e97;">
            <ul>
                <li><a href="index.asp">网站首页</a></li>
                <li><a href="about.asp?ClassID=1">机构简介</a></li>
                <li><a href="about.asp?ClassID=2">业务范围</a></li>
                <li><a href="piclist.asp?ClassID=1">法医学专家</a></li>
                <li><a href="piclist.asp?ClassID=2">医学专家</a></li>
                <li><a href="piclist.asp?ClassID=3">其他专家</a></li>
                <li><a href="news.asp?ClassID=2">相关咨询</a></li>
                <li><a href="news.asp?ClassID=3">相关动态</a></li>
                <li><a href="news.asp?ClassID=4">专家论丛</a></li>
                <li><a href="news.asp?ClassID=5">政策法规</a></li>
                <li style="background: none;"><a href="about.asp?ClassID=3">加入我们</a></li>
                </ul>
        </div>
        <div class="foot-tools clearfix">
            <div class="tj">网站访问量统计  :  123456</div>
            <div class="copy">版权所有：浙江省天平鉴定辅助技术研究院主办　苏ICP备13032904号 </div>
        </div>
    </div>
</div>
<script language=JavaScript>
    lastScrollY=0;
    function heartBeat(){
        var diffY;
        if (document.documentElement && document.documentElement.scrollTop)
            diffY = document.documentElement.scrollTop;
        else if (document.body)
            diffY = document.body.scrollTop
        else
        {/*Netscape stuff*/}
        percent=.1*(diffY-lastScrollY);
        if(percent>0)percent=Math.ceil(percent);
        else percent=Math.floor(percent);
        document.getElementById("full").style.top=parseInt(document.getElementById("full").style.top)+percent+"px";
        lastScrollY=lastScrollY+percent;
    }
    suspendcode="<div id=\"full\" align=\"center\" style='right:8px;POSITION:absolute;TOP:100px;z-index:100;background-image:url(images/qqback.gif);width:118px;height:200px;line-height: 15px;'><div style='margin-top:80px;line-height: 15px;'></div><a href='#Top' style='line-height: 15px;'><b style='line-height: 15px;'>在线交谈</b></a><br><a target=_blank style='line-height: 15px;' href=http://wpa.qq.com/msgrd?v=3&uin=1916740513&site=qq&menu=yes><img border=0 src=http://wpa.qq.com/pa?p=2:1916740513:51 alt=点击这里给我发消息 title=点击这里给我发消息/></a><br><div style='margin-top:15px'></div><a href='#Top'><b>咨询热线</b></a><br><font color=#CC0000><b>0571-87358618</b></font><br><b>传真:</b></a><br><font color=#CC0000><b>0571-87358617</b></font></div>"
    document.write(suspendcode);
    window.setInterval("heartBeat()",1);
</script>