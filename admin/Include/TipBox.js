/**
 * Corp 仿Discuz!论坛评分发帖弹出提示
 * @author jtaosj
 * @date 2010-11-2
 */
var x = window.x||{};
x.creat = function(t,b,c,d,text) {
    this.t=t;
    this.b=b;
    this.c=c;
    this.d=d;
    this.op=1;
    this.div=document.createElement("div");
    this.div.style.height="42px";
    this.div.style.width="234px";
    this.div.style.position="absolute";
    this.div.style.left="50%";
    this.div.style.marginLeft="-100px";
    this.div.style.marginTop="-20px";
    //this.div.innerText=text;
    this.div.style.lineHeight=this.div.style.height 
    this.div.style.top=(this.b+"%");
    
    var leftDiv = document.createElement("div");
    leftDiv.style.height = "42px";
    leftDiv.style.width="10px";
    leftDiv.style.background="url(../images/TipBox.gif) -9px -9px no-repeat";
    leftDiv.style.styleFloat="left";
    
    var mDiv = document.createElement("div");
    mDiv.style.height = "42px";
    mDiv.style.width="214px";
    mDiv.style.background="url(../images/TipBox.gif) 0 -65px repeat-x";
    mDiv.style.styleFloat="left";
    
    var txtDiv=document.createElement("div");
    txtDiv.style.height = "42px";
    txtDiv.style.fontSize = "15";
    txtDiv.style.textAlign="center";
    txtDiv.style.fontWeight="bold"; 

    //txtDiv.style.padding="10px 0 0 0";
    txtDiv.style.color = "#fff";
    txtDiv.style.background="url(../images/TipBox.gif) 0 -121px repeat-x";    
    txtDiv.innerText = text;
    
    mDiv.appendChild(txtDiv);
    
    var rightDiv = document.createElement("div");
    rightDiv.style.height = "42px";
    rightDiv.style.width="10px";
    rightDiv.style.background="url(../images/TipBox.gif) -43px -9px no-repeat";
    rightDiv.style.styleFloat="left";
    
    this.div.appendChild(leftDiv);
    this.div.appendChild(mDiv);
    this.div.appendChild(rightDiv);
    
    document.body.appendChild(this.div);
    this.run();
}
x.creat.prototype = {
    run:function(){
        var me=this;
        //this.div.style.top=-this.c*(this.t/this.d)*(this.t/this.d)+this.b+"%";
        this.div.style.top=-this.c*(this.t/this.d)+this.b+"%";
        this.t++;
        this.q=setTimeout(function(){me.run()},25)
        if(this.t==this.d){
            clearTimeout(me.q);
            setTimeout(function(){me.alpha();},1000);
        }
    },
    alpha:function(){
        var me=this;
        if("\v"=="v"){
            this.div.style.filter="progid:DXImageTransform.Microsoft.Alpha(opacity="+this.op*100+")";
            this.div.style.filter="alpha(opacity="+this.op*100+")";
        ;}
        else{this.div.style.opacity=this.op}
        this.op-=0.02;
        this.w=setTimeout(function(){me.alpha()},25)
        if(this.op<=0){
            clearTimeout(this.w);
            document.body.removeChild(me.div);
        }
    }
}

