
var MyDB=function(p){
	
	this.StrPath=p;
	this.getpath=function(){return this.StrPath}

	this.open = function(wr){
		if(wr!==0){
			imexstr="IMEX=0";
		}
		else{
			imexstr="IMEX=2"
		}
			this.Conn = new ActiveXObject("ADODB.Connection");
			this.Rst = new ActiveXObject("ADODB.recordset");
			if (this.Conn.state==1)
				this.Conn.close()
			this.Conn.open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + this.StrPath + "';Extended Properties='Excel 12.0;HDR=yes;IMEX=2'")

	}
	this.close=function(){
		if(this.Rst.State){
			this.Rst.close()
			this.Rst=null;
		}
		if(this.Conn.State){
			this.Conn.close()
			this.Conn=null;
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
		this.exe("select * from [" + table + "$A1:O300] where [ID] = " + id + "");
		myfrom="";
		if(!this.Rst.EOF){
			
			var myfrom="<div class='modal-header'><H2>Detailed</H2></div><div class='modal-body'><form class='form-horizontal'>";
			for(i=1;i<this.Rst.Fields.count;i++){
				colname=this.Rst.Fields(i).name;//replace(/[\u4e00-\u9fa5_]/g,"");
				myfrom += "<div class='control-group'><label for='"+ colname +"' class='control-label'>" + colname + "</label><div class='controls'><span class='info' id='"+ colname+ "'>" + preprocessdata(this.Rst.Fields(i).value) + "</span></div></div>"
			}
			myfrom += "</form></div>"
			myfrom += "<div class='modal-footer'><div class='warning'>非管理员请勿编辑</div><button class='btn' data-dismiss='modal' aria-hidden='true'>Close</button><button id='edit' class='btn btn-warning' title='Admin Only please!'>EDIT</button></div>"			
		}
		return myfrom;
		this.close();
	}
	
	this.updateData=function(table,id){
		
		sql="update [" + table + "$A1:O300] set ";
		inputs=$("form input[chd]");
		for(i=0;i<inputs.length;i++){
			myid=$("form input[chd]:eq("+ i +")").attr('id');
			myvalue=$("form input[chd]:eq("+ i +")").val();
				if(myid.toLowerCase().indexOf('date')!=-1){
					r=new RegExp(String.fromCharCode(8206),'g')
					myvalue=myvalue.replace(r,'')
					sql += "[" + jQuery.trim(myid) + "] = #" + myvalue + "# ," //dateconvert(myvalue)
				}else{
					sql += "[" + jQuery.trim(myid) + "] = '" + myvalue +"' ," 
				}
			
		}
		if(inputs.length!==0){
			sql=sql.slice(0,sql.length-1)
			sql += " where [ID] = " + id;
			//alert(sql);
			//return false;
			this.open(true);
			try{
			this.Conn.Execute(sql);
			}catch(e){
				this.close()
				alert("Sorry, you met a problem says:"+e.description +"\n\n Update not successful Data is unchanged!, Please contact Designer!\n\n the SQL is:" +sql);
				return
			}
			this.close();
			alert("Success!"+sql);
		}
		//bugfix: if user didn't change anything
		else{
			alert("You didn't change anything!")
		}
		
		}
}

//"update [Pump$A1:O300] set `Update date` = #‎2016-‎11‎-‎3‎#  where [ID] = 1"
function addslashes(str){
	return (str + '').replace(/[\\"']/g,'\\$&').replace(/\u0000/g,'\\0');	
}
function preprocessdata(v){
	//todo change null to "" chang date to date string
	re="";
	if(v==null){
		return "";
	}
	if(typeof(v)=="date"){
	var d=new Date(v+"")
		return d.toLocaleDateString();
	}
	return v;
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
				$("#placeholder").html("<ul><li><h2>错误：</h2><li>你所请求的<a href='"+ dbpath +"'> EXCEL文件 </a>正在被独占打开.点击查看只读版本</li><li>错误描述：" + e.description + "</li></ul>");
				
				b.close()
			}
			//html+="<div id=extrafilter>Serch in result: <input type=text name=filter id=filter placeholder='Search'><span id='err'></span></div>";
			//todo:show table 
			html+=b.printTable()
			$("#placeholder").html(html);
			b.close();
			
			//todo:show navigate
			$("#nav").html("<ul class=breadcrumb >You are here : <li>Home<span class='divider'>&raquo;</span></li><li>Overview Partlist<span class='divider'>&raquo;</span></li><li>" + comp + "<span class='divider'>&raquo;</span></li><li>" + $(this).text() + "</li><li class='label label-warning'>No update for non-admin<li></ul>")
			//$("#nav").find('a').attr("href",dbpath);
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
	$("table thead tr").append($("<th>detail</th>"))
	$("table tbody tr").append($("<td><a line>More</a></td>"))
	//$("td[line]").html(function(){
	//return $("<a href='#' title='ShowAll' line='"+$(this).attr("line")+"'>"+$(this).attr("line")+"</a>")
	//})
	$("a[line]").click(function(){
		var mytable=$("[disabled]").attr("table");
		var myid=$(this).parents("tr").find("td[line]").attr("line");
		$("#modal").html(b.showInput(mytable,myid));
		$("#modal").modal("show");
		$("#modal span.info").addClass("uneditable-input");
		
		
		$("#edit").click(function(){
			$("#modal")
			.find(".uneditable-input")
			.parent()
			.html(function(){
				return "<input id='" +$(this).find("span.info").attr("id")+ "' type='text' value='" + $(this).text()+"'>"
			})
			.removeClass("uneditable-input")
			$(this).attr("ID","sav").text("SaveChanges").removeClass("btn-warning").addClass("btn-danger")
			
			$("#sav").click(function(){
			if($("form input[chd]").length!=0){
				searchStr=$("input[type='search']").val()
				b.updateData(mytable,myid);
				$("[disabled]").removeAttr("disabled").click()
				//$("input[type='search']").val(searchStr).keydown()
			}
			$("#modal").modal("hide");
			})
			$("#modal form :text").change(function(){
			$(this).attr('chd',1)
			$(this).parents('.control-group').addClass("warning")
		})
		
		})
		
	})
	}
	
	$('#datatable').DataTable();

	
	
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
***************/

}

var mJSON={
	"CRS":{
		name:"CRS components",
		path:"Q:/801_ESD1/000_Laboratory/105_Lab_Management/200_ESD1_Overview_Partlist/CRS_Component/totalCRS.xlsx",
		buttonlist:{
			"Pump":["Pump","select `ID`,`Name`, PN as [Part No], SN as [Serial No], [ZP type] as Feedpump, `Geo# Vol(mm^3/rev)` as [GEO VOL], `Position`, `Status` from [Pump$A1:Q600] "],
			"ECU":["ECU","select `ID`,`Type` as [Hardware Ver], [Part number] as [Part No], [Series Number] as [Series No], `Position`, `Status` from [ECU$A1:J600] "],
			"Injector":["Injector","select FIRST(`ID`) as [ID],`Name`,`PN` as [Part No], COUNT(`PN`) as [Good condition],`Flowrate(cm^3/30s)` as FLOW from [Injector$A1:J600] where [Status]='OK' GROUP BY `Name`, `PN`, `Flowrate(cm^3/30s)` "],
			"Rail":["Rail",""],
			"WH":["Wire harness","select `ID`,`Name` as [Name], [Part number] as [Part No], [Series Number] as [Series No], `Position`, `Status` from [WH$A1:J600] "],
			"Spare":["Spare parts",""]
		}
	},
	"Fixtures":{
		name:"Fixtures",
		path:"Q:/801_ESD1/000_Laboratory/105_Lab_Management/200_ESD1_Overview_Partlist/Fixture/Fixture.xlsx",
		buttonlist:{
			"Adaptor":["Adaptor",""],
			"Holder":["Holder",""],
			"InjAdp":["Injector adaptor",""],
			"InjClp":["Injector clamp",""],
			"ShaftAdpNut":["Shaft adaptor & Nut",""],
			"Spannpratze":["Spannpratze",""],
			"PumpHold_Flange":["PumpHold&Flange",""]
		}
	},
	"Consumable":{
		name:"Consumable",
		path:"Q:/801_ESD1/000_Laboratory/105_Lab_Management/200_ESD1_Overview_Partlist/Consumable/TotalConsumable.xlsm",
		buttonlist:{
			"Nut_Screw_Bolt":["Nut screw & bolt",""],
			"Tie-in":["Tie-in",""],
			"Other":["Other",""],
			"Pipe":["Pipe",""]
		}
	},
	"Tools":{
		name:"Tools",
		path:"Q:/801_ESD1/000_Laboratory/105_Lab_Management/200_ESD1_Overview_Partlist/Tools/TotalTool.xlsm",
		buttonlist:{
			"TorqueWrench":["Torque Wrench",""],
			"SitecSlice":["Safyte Slice(SETEC)",""],
			"ParticleCounter":["Particle Counter",""],
			"Oscilloscope":["Oscilloscope",""],
			"HDA":["HDA",""],
			"EVI":["EVI",""],
			"EMI":["EMI",""],
			"ETAS":["ETAS",""],
			"DigitalCaliper":["Digital Caliper",""],
			"CurrentProbe":["Current Probe",""]
		}
	},
	"Sensors":{
		name:"Sensors",
		path:"Q:/801_ESD1/000_Laboratory/105_Lab_Management/200_ESD1_Overview_Partlist/Sensor/TotalSensor.xlsm",
		buttonlist:{
			"Torque":["Torque sensor",""],
			"WIKA":["WIKA pressure sensor",""],
			"SkinThermo":["Skin temperature sensor",""],
			"PLU":["PLU",""],
			"NeedleLift":["Needle lift sensor",""],
			"FlowMeter":["Natec Flowmeter",""],
			"NatecPressue":["Natec Pressure sensor",""],
			"Massflow":["Mass flowmeter",""],
			"KistlerAmp":["Kistler amplifier",""],
			"Kistler":["Kistler pressure sensor",""],
			"FuelThermo":["Fuel termperature sensor",""]
		}
	},	
	"Others":{
		name:"Spare, HPP etc.",
		path:"",
		buttonlist:{
		},
		links:{
			"TB_Spares":"file:///Q:/801_ESD1/000_Laboratory/105_Lab_Management/200_ESD1_Overview_Partlist/Spares.xls",
			"Gatepass":"file:///Q:/801_ESD1/000_Laboratory/105_Lab_Management/200_ESD1_Overview_Partlist/Gate%20pass%20for%20objects/list.xls",
			"HPP infomation":"file:///Q:/801_ESD1/000_Laboratory/104_Consumer_Goods/201_Consummable/362_Pipe/402_HPP/Order_form",
			"New HPP request":"file:///Q:/801_ESD1/006_Process/106_working%20sheet"
		}
	}	
}

	
