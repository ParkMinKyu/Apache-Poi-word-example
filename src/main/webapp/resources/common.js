var modal = false;
$(document).bind("ajaxSend",function(){
	if(!modal){
		modal = true;
	}
}).bind("ajaxComplete",function(){
	if(modal){
		modal = false;
	}
});