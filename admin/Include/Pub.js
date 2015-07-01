function $Get(Obj){return document.getElementById(Obj);}
function openw(url,name,w,h){
    var url;                             //转向网页的地址;
    var name;                            //网页名称，可为空;
    var iTop = (window.screen.availHeight-30-h)/2;   //获得窗口的垂直位置     
    var iLeft = (window.screen.availWidth-10-w)/2;    //获得窗口的水平位置     
    window.open(url,name,'height='+ h +',,innerHeight='+ h +',width='+ w +',innerWidth='+ w +',top='+iTop+',left='+iLeft+',status=no,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=0,titlebar=no');
}
