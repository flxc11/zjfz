function Flash(url,w,h,s){
	if (s==1){
	document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="'+w+'" height="'+h+'"> ');
	document.write('<param name="movie" value="' + url + '">');
	document.write('<param name="quality" value="high"> ');
	document.write('<param name="wmode" value="transparent"> ');
	document.write('<param name="menu" value="false"> ');
	document.write('<embed src="' + url + '" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="'+w+'" height="'+h+'" wmode="transparent"></embed> ');
	document.write('</object> ');
	}
	else{
	document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="'+w+'" height="'+h+'"> ');
	document.write('<param name="movie" value="' + url + '">');
	document.write('<param name="quality" value="high"> ');
	document.write('<param name="menu" value="false"> ');
	document.write('<embed src="' + url + '" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="'+w+'" height="'+h+'"></embed> ');
	document.write('</object> ');
	}
}
function showPicup(ID,Pic){	
	ID.setAttribute("src","images/"+Pic+".png");	
	}
var index
function Navshow(pic1,pic2,tiltes,url,id){
	if(id==index){		
		document.write("<img src='images/"+pic2+".png' title='"+tiltes+"' />")
		}else{			
		document.write("<a href='"+url+"' title='"+tiltes+"'><img onmouseover=showPicup(this,'"+pic2+"') onmouseout=showPicup(this,'"+pic1+"') src='images/"+pic1+".png' border='0'/></a>")	
		}
}
function chkguetform(){
var Emailfiter=/^([a-zA-Z0-9_-])+@([a-zA-Z0-9_-])+(\.[a-zA-Z0-9_-])+/;
  if(jQuery('#name').val()=='')
  {
    alert('请填写名字！');
	jQuery('#name').focus();
	return false;
  }
  else if(jQuery('#phone').val()=='')
  {
    alert('联系电话不能为空！');
	jQuery('#phone').focus();
	return false;
  }
   else if(jQuery('#address').val()=='')
  {
    alert('地址不能为空！');
	jQuery('#address').focus();
	return false;
  }
   else if(jQuery('#email').val()=='')
  {
    alert('邮件地址不能为空！');
	jQuery('#email').focus();
	return false;
  }
  else if(!Emailfiter.test(jQuery('#email').val()))
  {
    alert('邮件地址格式不正确！');
	jQuery('#email').focus();
	return false;
  }
   else if(jQuery('#title').val()=='')
  {
    alert('标题不能为空！');
	jQuery('#title').focus();
	return false;
  }
 else if(jQuery('#content').val()=='')
  {
    alert('内容不能为空！');
	$('#content').focus();
	return false;
  } 
  else
  {
  jQuery('#content').val() = jQuery('#content').val().replace( /<script.*?>(.|\s|\r|\r\n)*?<\/script>/gim, "" );
  return true;
  }
}
function chkguetform2(){
var Emailfiter=/^([a-zA-Z0-9_-])+@([a-zA-Z0-9_-])+(\.[a-zA-Z0-9_-])+/;
  if(jQuery('#name').val()=='')
  {
    alert('请填写名字！');
	jQuery('#name').focus();
	return false;
  }
  else if(jQuery('#phone').val()=='')
  {
    alert('联系电话不能为空！');
	jQuery('#phone').focus();
	return false;
  }
   else if(jQuery('#address').val()=='')
  {
    alert('地址不能为空！');
	jQuery('#address').focus();
	return false;
  }
   else if(jQuery('#email').val()=='')
  {
    alert('邮件地址不能为空！');
	jQuery('#email').focus();
	return false;
  }
  else if(!Emailfiter.test(jQuery('#email').val()))
  {
    alert('邮件地址格式不正确！');
	jQuery('#email').focus();
	return false;
  }
   else if(jQuery('#title').val()=='')
  {
    alert('产品数量不能为空！');
	jQuery('#title').focus();
	return false;
  }
 else if(jQuery('#content').val()=='')
  {
    alert('内容不能为空！');
	$('#content').focus();
	return false;
  } 
  else
  {
  jQuery('#content').val() = jQuery('#content').val().replace( /<script.*?>(.|\s|\r|\r\n)*?<\/script>/gim, "" );
  return true;
  }
}



function tab1(a)
{
	if(a==1)
	{
		document.getElementById("tab21").className="TypeSelect";
		document.getElementById("tab22").className="Type";
		document.getElementById("f21").className="f_w";
		document.getElementById("f22").className="f_g";
		document.getElementById("tzgg1").style.display = "";
		document.getElementById("tzgg2").style.display = "none";
	}
	if(a==2)
	{
		document.getElementById("tab21").className="Type";
		document.getElementById("tab22").className="TypeSelect";
		document.getElementById("f21").className="f_g";
		document.getElementById("f22").className="f_w";
		document.getElementById("tzgg1").style.display = "none";
		document.getElementById("tzgg2").style.display = "";
	}
}




function TabNav(a)
{
	if(a==1)
	{
		document.getElementById("tab21").className="TypeSelect";
		document.getElementById("tab22").className="Type";
		document.getElementById("tzgg1").style.display = "";
		document.getElementById("tzgg2").style.display = "none";
	}
	if(a==2)
	{
		document.getElementById("tab21").className="Type";
		document.getElementById("tab22").className="TypeSelect";
		document.getElementById("tzgg1").style.display = "none";
		document.getElementById("tzgg2").style.display = "";
	}
}




function tab2(a)
{
	if(a==1)
	{
		document.getElementById("tab21").className="TypeSelect";
		document.getElementById("tab22").className="Type";
		document.getElementById("tab23").className="Type";
		document.getElementById("tab24").className="Type";
		document.getElementById("tab25").className="Type";
		document.getElementById("tab26").className="Type";
		document.getElementById("f21").className="f_on";
		document.getElementById("f22").className="f_off";
		document.getElementById("f23").className="f_off";
		document.getElementById("f24").className="f_off";
		document.getElementById("f25").className="f_off";
		document.getElementById("f26").className="f_off";
		document.getElementById("tzgg1").style.display = "";
		document.getElementById("tzgg2").style.display = "none";
		document.getElementById("tzgg3").style.display = "none";
		document.getElementById("tzgg4").style.display = "none";
		document.getElementById("tzgg5").style.display = "none";
		document.getElementById("tzgg6").style.display = "none";
	}
	if(a==2)
	{
		document.getElementById("tab21").className="Type";
		document.getElementById("tab22").className="TypeSelect";
		document.getElementById("tab23").className="Type";
		document.getElementById("tab24").className="Type";
		document.getElementById("tab25").className="Type";
		document.getElementById("tab26").className="Type";
		document.getElementById("f21").className="f_off";
		document.getElementById("f22").className="f_on";
		document.getElementById("f23").className="f_off";
		document.getElementById("f24").className="f_off";
		document.getElementById("f25").className="f_off";
		document.getElementById("f26").className="f_off";
		document.getElementById("tzgg1").style.display = "none";
		document.getElementById("tzgg2").style.display = "";
		document.getElementById("tzgg3").style.display = "none";
		document.getElementById("tzgg4").style.display = "none";
		document.getElementById("tzgg5").style.display = "none";
		document.getElementById("tzgg6").style.display = "none";
	}
	if(a==3)
	{
		document.getElementById("tab21").className="Type";
		document.getElementById("tab22").className="Type";
		document.getElementById("tab23").className="TypeSelect";
		document.getElementById("tab24").className="Type";
		document.getElementById("tab25").className="Type";
		document.getElementById("tab26").className="Type";
		document.getElementById("f21").className="f_off";
		document.getElementById("f22").className="f_off";
		document.getElementById("f23").className="f_on";
		document.getElementById("f24").className="f_off";
		document.getElementById("f25").className="f_off";
		document.getElementById("f26").className="f_off";
		document.getElementById("tzgg1").style.display = "none";
		document.getElementById("tzgg2").style.display = "none";
		document.getElementById("tzgg3").style.display = "";
		document.getElementById("tzgg4").style.display = "none";
		document.getElementById("tzgg5").style.display = "none";
		document.getElementById("tzgg6").style.display = "none";
	}
	if(a==4)
	{
		document.getElementById("tab21").className="Type";
		document.getElementById("tab22").className="Type";
		document.getElementById("tab23").className="Type";
		document.getElementById("tab24").className="TypeSelect";
		document.getElementById("tab25").className="Type";
		document.getElementById("tab26").className="Type";
		document.getElementById("f21").className="f_off";
		document.getElementById("f22").className="f_off";
		document.getElementById("f23").className="f_off";
		document.getElementById("f24").className="f_on";
		document.getElementById("f25").className="f_off";
		document.getElementById("f26").className="f_off";
		document.getElementById("tzgg1").style.display = "none";
		document.getElementById("tzgg2").style.display = "none";
		document.getElementById("tzgg3").style.display = "none";
		document.getElementById("tzgg4").style.display = "";
		document.getElementById("tzgg5").style.display = "none";
		document.getElementById("tzgg6").style.display = "none";
	}
	if(a==5)
	{
		document.getElementById("tab21").className="Type";
		document.getElementById("tab22").className="Type";
		document.getElementById("tab23").className="Type";
		document.getElementById("tab24").className="Type";
		document.getElementById("tab25").className="TypeSelect";
		document.getElementById("tab26").className="Type";
		document.getElementById("f21").className="f_off";
		document.getElementById("f22").className="f_off";
		document.getElementById("f23").className="f_off";
		document.getElementById("f24").className="f_off";
		document.getElementById("f25").className="f_on";
		document.getElementById("f26").className="f_off";
		document.getElementById("tzgg1").style.display = "none";
		document.getElementById("tzgg2").style.display = "none";
		document.getElementById("tzgg3").style.display = "none";
		document.getElementById("tzgg4").style.display = "none";
		document.getElementById("tzgg5").style.display = "";
		document.getElementById("tzgg6").style.display = "none";
	}
	if(a==6)
	{
		document.getElementById("tab21").className="Type";
		document.getElementById("tab22").className="Type";
		document.getElementById("tab23").className="Type";
		document.getElementById("tab24").className="Type";
		document.getElementById("tab25").className="Type";
		document.getElementById("tab26").className="TypeSelect";
		document.getElementById("f21").className="f_off";
		document.getElementById("f22").className="f_off";
		document.getElementById("f23").className="f_off";
		document.getElementById("f24").className="f_off";
		document.getElementById("f25").className="f_off";
		document.getElementById("f26").className="f_on";
		document.getElementById("tzgg1").style.display = "none";
		document.getElementById("tzgg2").style.display = "none";
		document.getElementById("tzgg3").style.display = "none";
		document.getElementById("tzgg4").style.display = "none";
		document.getElementById("tzgg5").style.display = "none";
		document.getElementById("tzgg6").style.display = "";
	}
}