function $Get(Obj){return document.getElementById(Obj);}
function openw(url,name,w,h){
    var url;                             //ת����ҳ�ĵ�ַ;
    var name;                            //��ҳ���ƣ���Ϊ��;
    var iTop = (window.screen.availHeight-30-h)/2;   //��ô��ڵĴ�ֱλ��     
    var iLeft = (window.screen.availWidth-10-w)/2;    //��ô��ڵ�ˮƽλ��     
    window.open(url,name,'height='+ h +',,innerHeight='+ h +',width='+ w +',innerWidth='+ w +',top='+iTop+',left='+iLeft+',status=no,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=0,titlebar=no');
}
