ALERT_EACH_RESULT = false;
ALERT_EACH_RESULT_FNC = function(){if(arguments.length >0 ){ALERT_EACH_RESULT=arguments[0]} return ALERT_EACH_RESULT;}
STOP = false;
STOP_FNC = function(){if(arguments.length >0 ){STOP=arguments[0]} return STOP;}
ALLERT_MESSAGE = null; //MOSTRA NO MESSAGE BOX UM TEXTO PERSONALIZADO
ALLERT_MESSAGE_FNC = function(){if(arguments.length >0 ){ALLERT_MESSAGE=arguments[0]} return ALLERT_MESSAGE;}

Array.prototype.includes = function(obj) {
    var i = this.length;
    while (i--) {
        if (this[i] === obj) {
            return true;
        }
    }
    return false;
}


function xls2xlsx(from,to){
  var appExcel = new ActiveXObject("Excel.Application")
  appExcel.DisplayAlerts = false
  var Workbook = appExcel.Workbooks.open(from)
  appExcel.visible = false;
  Workbook.SaveAs(to,51, null, null, null, null, null, null, null, null, null, 1)
  Workbook.Close()
}


Array.prototype.contains = Array.prototype.includes;
String.prototype.includes = Array.prototype.includes;