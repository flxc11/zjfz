// JavaScript Document
$(function(){
	$('ul.topnav li div').each(function(i,v){
    	$(v).children("ul").children("li:last").remove();
	});
	$('ul.topnav li div:eq(6)').attr( 'style', 'left:-100px;');
	$('ul.topnav li').mouseenter(function() { 
		if($(this).find('div.subnav ul li').length>0)
			$(this).find('div.subnav').parent().children('a').first().removeClass('top');
			$(this).find('div.subnav').parent().children('a').first().addClass('tophover');
			$(this).find('div.subnav').slideDown('fast').show();
			$(this).hover(function() {	
				if($(this).find('div.subnav ul li').length>0)	
					$(this).find('div.subnav').parent().children('a').first().removeClass('top');
					$(this).find('div.subnav').parent().children('a').first().addClass('tophover');
					$(this).find('div.subnav').slideDown('fast').show(); 
				}, 
				function(){
					$(this).find('div.subnav').parent().children('a').first().removeClass('tophover');			
					$(this).find('div.subnav').parent().children('a').first().addClass('top');
					$(this).find('div.subnav').slideUp('fast');});		
				}).hover(function() { 			
						$(this).addClass('subhover');
						}, function(){
							$(this).removeClass('subhover');
							}
			);
})