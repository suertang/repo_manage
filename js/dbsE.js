
var MyDB=function(p){
	
	this.StrPath=p?p:"Q:\\801_ESD1\\html\\Links.xlsx";
	this.getpath=function(){return this.StrPath}
	this.Conn = new ActiveXObject("ADODB.Connection");
	this.Rst = new ActiveXObject("ADODB.recordset");
	this.open = function(){
		if(!this.Conn.State){
			this.Conn.open("Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;IMEX=1;';Data Source=" + this.StrPath + ";")
		}
	}
	this.close=function(){
		if(this.Rst.State){
			this.Rst.close()
		}
		if(this.Conn.State){
			this.Conn.close()
		}
	}
	this.showMsg=function(){
		alert(this.StrPath)
	}
	this.exe=function(sql){
		this.open();
		this.Rst=this.Conn.Execute(sql);
		return
	}
	this.printTable = function(){
	var t="";
	if(this.Rst.EOF){
		t+="No records found!";
	}
	else{
		t+="<table class='table table-bordered' id='datatable' ><thead><tr>";
		for(i=0;i<this.Rst.Fields.count;i++){
			
				t+="<th >" + this.Rst.Fields(i).name.replace(/[\u4e00-\u9fa5_]/g,"");
			
		}
		t+="</tr></thead><tbody>"
			while(!this.Rst.EOF){
				
					t+="<tr>";
				
				for(i=0;i<this.Rst.Fields.count;i++){
					td=this.Rst.Fields(i);
					if(td.value!==null){

						if(typeof(this.Rst.Fields(i).value) == "date"){
							//h=new Date('#'+this.Rst.Fields(i).value+'#')
							var d=new Date(this.Rst.Fields(i).value+"")
							td=d.toLocaleDateString();
						}
					}
					else{
					td="";
					}
					//td=td.replace(/[\u4e00-\u9fa5]/g,"");
					if(this.Rst.Fields(i).name.replace(/[\u4e00-\u9fa5_]/g,"")=="ID"){
						t+="<td line='" + td +"'>" + td;
					}
					else{
						t+="<td>" + td;
					}
					
				}
				this.Rst.movenext();
			}
		t+="</tbody></table>";
	}
	return t;
	}
	this.showInput=function(table,id){
		this.open();
		this.exe("select * from [" + table + "$A2:BB2000] where [ID] = " + id + "");
		myfrom="";
		if(!this.Rst.EOF){
			
			var myfrom="<div class='modal-header'><H2>Detailed</H2></div><div class='modal-body'><form class='form-horizontal'>";
			for(i=1;i<this.Rst.Fields.count;i++){
				colname=this.Rst.Fields(i).name.replace(/[\u4e00-\u9fa5_]/g,"");
				myfrom += "<div class='control-group'><label for='"+ colname +"' class='control-label'>" + colname + "</label><div class='controls'><input id='"+ colname+ "' type='text' value='" + this.Rst.Fields(i).value +"'></div></div>"
			}
			myfrom += "</form></div>"
			myfrom += "<div class='modal-footer'><button class='btn' data-dismiss='modal' aria-hidden='true'>Close</button><button class='btn btn-primary'>Save changes</button></div>"			
		}
		return myfrom;
		this.close();
	}
}


$(document).ready(function(){
	////// todo: add navbar
		$("body").append($("<div id='modal' class='modal fade'></div>"));
		$(".nav").append(function(){
				html="";
				for(i in mJSON){
				 html+="<li class='dropdown'><a class='dropdown-toggle' data-toggle='dropdown' href='#'>" + mJSON[i].name + "<b class='caret'></b></a></li>"
				}
				return $(html);
			}).end().find(".dropdown")
		.append(function(){
			var txt=$(this).text();			
			return $(makedropdown(txt))
		})
		.find('ul.dropdown-menu li a').attr('comp',function(){
			return $(this).parents('.dropdown').children(":first").text()
		})
		.filter("[table]").click(function(){
			var html="";
			//todo:init database
			var comp =  $(this).parents('.dropdown').children(":first").text()
			var table = $(this).attr("table");
			var dbpath = getpath(comp);
			b=new MyDB(dbpath);
			sql=getsql(comp,table);
			if(!sql){
				sql="select * from [" + table + "$A1:I600] "
			}
			try{
				b.exe(sql);
			}
			catch(e){
				$("#placeholder").html(e.description+sql);
			}
			//html+="<div id=extrafilter>Serch in result: <input type=text name=filter id=filter placeholder='Search'><span id='err'></span></div>";
			//todo:show table 
			html+=b.printTable()
			$("#placeholder").html(html);
			b.close();
			
			//todo:show navigate
			$("#nav").html("<ul class=breadcrumb >You are here : <li>Home<span class='divider'>&raquo;</span></li><li>Overview Partlist<span class='divider'>&raquo;</span></li><li>" + comp + "<span class='divider'>&raquo;</span></li><li>" + $(this).text() + "</li> <span class='divider'>|</span><a href=# class='label label-info'>edit</a></ul>")
			$("#nav").find('a').attr("href",dbpath);
			//todo:apply make filter avaiable
			$(".active").removeClass("active");
			$(this).parents(".dropdown").addClass("active");
			
			$("[disabled]").removeAttr("disabled");
			$(this).attr("disabled","true");
			
			addlisten();
		})
});
function getpath(comp){
	var ret=""
	for (i in mJSON){
		if(comp==mJSON[i].name){
			ret=mJSON[i].path;
			break;
		}
		
	}
	return ret;
}
function getsql(comp,button){
	var ret=""
	for (i in mJSON){
		if(comp==mJSON[i].name){
			for (j in mJSON[i].buttonlist){
				if(button===j){
					ret=mJSON[i].buttonlist[j][1];
					break;
				}
			}
		}
	}
	return ret;
}

function makedropdown(comp){
	var listr="<ul class='dropdown-menu' >";
	for(i in mJSON){
		if(mJSON[i].name==comp.replace(/[\n\r\t]/g,"")){
			for(j in mJSON[i].buttonlist){
				listr += "<li><a href='#' table='" + j +"'>" + mJSON[i].buttonlist[j][0]+ "</a></li>"
			}
			for(k in mJSON[i].links){
				listr += "<li><a href='" + mJSON[i].links[k] + "'>" + k + "</a></li>"
			}
		}
	}	
	listr += "</ul>"	
	return listr;
}
	//defaultsql:"select `Name`,`PN` as [Part No], COUNT(`PN`) as [Good condition],`Flowrate(cm^3/30s)` as FLOW from [Injector$A1:J600] where [Status]='OK' ",

function addlisten(){
	
	if($("table thead tr").text().indexOf("ID")!==-1){
	$("table tbody tr").dblclick(function(){
		$("#modal").html(b.showInput($("[disabled]").attr("table"),$(this).find("td[line]").text()));
		$("#modal").modal("show");
	})
	}
	$('#datatable').DataTable();
	$("#filter").keyup(function(){		
		$("table tbody tr")
			.hide()
			//.filter(":contains(';"+($(this).val())+"')")
			.filter(function(){
			if(!$("#filter").val()){return true}
			try{
			re=new RegExp('<td>'+($("#filter").val().toLowerCase()),'g');
			return re.test(this.innerHTML.toLowerCase());// || re1.test(this.innerHTML);
			}
			catch(e){
				$('err').text("<span>" + e.description + "</span>");
			return true;}
			})
			//return (new RegExp('\s'+n,'i')).test();})
			.show();		
	}).keyup();
	
	
//"select * from [Database$A1:O300] where [ID] = 1"
	
	
/***************
	$('table tbody tr td')
		.dblclick(function(){
			if( !$(this).is('.flag') ){
				$(this).addClass('flag').css('height',$(this).innerHeight()-4)
					.html('<input type=\'text\' value="' + $(this).text() +'" />')
					.find('input').addClass('input').focus()
					.blur(function(){
						$(this).parent().html($(this).val() || "").removeClass('flag');
					});					
			}
		});
***************/

}

var mJSON={
	"Search":{
		name:"Database search",
		path:"Q:/801_ESD1/007_Experience_Sharing/104_Customer project database/202_Database/Customer Project Database_V2.6_2016.xlsm",
		buttonlist:{
			"Database":["All Test data","select `ID`, `Test No`, `TB`, `Testing Type`,`Customer`, `Engine`, `Emission`, `Pump Type`,`SAP No`  from [Database$A2:BB3000] "]
		}
	},

	"Reports":{
		name:"Online reports",
		path:"",
		buttonlist:{
		},
		links:{
			"link1":"#",
			"link2":"#",
			"link3":"#",
			"link4":"#"
		}
	}	
}

	
