// JavaScript Document
function initEcAd() {
document.all.AdLayer1.style.posTop = -200;
document.all.AdLayer1.style.visibility = 'visible'
document.all.AdLayer2.style.posTop = -200;
document.all.AdLayer2.style.visibility = 'visible'
MoveLeftLayer('AdLayer1');
MoveRightLayer('AdLayer2');
}
function MoveLeftLayer(layerName) {
var x = 5;
var y = 100;// 左侧广告距离页首高度
var diff = (document.documentElement.scrollTop + y - document.all.AdLayer1.style.posTop)*.40;
var y = document.documentElement.scrollTop + y - diff;
eval("document.all." + layerName + ".style.posTop = parseInt(y)");
eval("document.all." + layerName + ".style.posLeft = x");
setTimeout("MoveLeftLayer('AdLayer1');", 20);
}
function MoveRightLayer(layerName) {
var x = 5;
var y = 100;// 右侧广告距离页首高度
var diff = (document.documentElement.scrollTop + y - document.all.AdLayer2.style.posTop)*.40;
var y = document.documentElement.scrollTop + y - diff;
eval("document.all." + layerName + ".style.posTop = y");
eval("document.all." + layerName + ".style.posRight = x");
setTimeout("MoveRightLayer('AdLayer2');", 20);
}
function imgclose(){
	document.getElementById("AdLayer1").style.display="none";
	document.getElementById("AdLayer2").style.display="none";
}

document.write("<div id=AdLayer1 style='position: absolute;visibility:hidden;z-index:1'></div>");
document.write("<div id=AdLayer2 style='position: absolute;visibility:hidden;z-index:1'><a href='http://www.netpolice.gov.cn/wj/index.asp' target='_blank'><img src=images/imgrightpf.jpg border='0' usemap='#Map'></a><br/><a href='javascript:imgclose()'>关闭</a></div><map name='Map' id='Map'><area shape='rect' coords='7,33,73,99' href='/Col/Col7/Index.aspx' target='_blank' /><area shape='rect' coords='7,99,72,164' href='/Col/Col29/Index.aspx' target='_blank' /><area shape='rect' coords='6,163,72,229' href='/03534636/GuestBook.aspx?Action=Write' /><area shape='rect' coords='6,229,72,296' href='/Col/Col8/Index.aspx' target='_blank' /><area shape='rect' coords='9,296,71,364' href='/Col/Col47/Index.aspx' target='_blank' /><area shape='rect' coords='7,366,70,434' href='https://www.zj96596.com/' target='_blank' /></map>");
initEcAd()