<%
Dim str_FINALSEMANA
Dim str_FERIADOS
Dim str_DIA_HOJE
Dim str_DATA_INICIAL
Dim str_NAO_HABILITAR
If Request.Form("FINALSEMANA") <> "" Then
	str_FINALSEMANA = Request.Form("FINALSEMANA")
ElseIf Request.QueryString("FINALSEMANA") <> "" Then
	str_FINALSEMANA = Request.QueryString("FINALSEMANA")
End If
If Request.Form("MENORIGUAL") <> "" Then
	str_MENORIGUAL = Request.Form("MENORIGUAL")
ElseIf Request.QueryString("MENORIGUAL") <> "" Then
	str_MENORIGUAL = Request.QueryString("MENORIGUAL")
End If
If Request.Form("MAIORIGUAL") <> "" Then
	str_MAIORIGUAL = Request.Form("MAIORIGUAL")
ElseIf Request.QueryString("MAIORIGUAL") <> "" Then
	str_MAIORIGUAL = Request.QueryString("MAIORIGUAL")
End If
str_FERIADOS  = Request.Form("FERIADOS")
str_NAO_HABILITAR = Request.QueryString("pNAO_HABILITAR")
'response.Write("<p>str_NAO_HABILITAR = " & str_NAO_HABILITAR)
'response.Write("<p>str_MENORIGUAL = " & str_MENORIGUAL)
'response.Write("<p>str_MAIORIGUAL = " & str_MAIORIGUAL)
'response.End()
If Request.QueryString("DATAINICIAL") <> "" Then
	str_DATA_INICIAL = Year(CDate(Request.QueryString("DATAINICIAL"))) & "," & Month(CDate(Request.QueryString("DATAINICIAL")))-1 & "," & Day(CDate(Request.QueryString("DATAINICIAL")))
Else
	str_DATA_INICIAL = "1900,1,1"
End If
If Request.QueryString("DATAFINAL") <> "" Then
	str_DATA_FINAL = Year(CDate(Request.QueryString("DATAFINAL"))) & "," & Month(CDate(Request.QueryString("DATAFINAL")))-1 & "," & Day(CDate(Request.QueryString("DATAFINAL")))
Else
	str_DATA_FINAL = "2040,1,1"
End If
str_DIA_HOJE = Year(Date()) & "," & Month(Date())-1 & "," & Day(Date())
'response.Write(str_DIA_HOJE)
'response.Write("str_DATA_INICIAL = " & str_DATA_INICIAL)
'response.Write("str_DATA_FINAL = " & str_DATA_FINAL)
'response.End()
%>
<HTML>
<HEAD>
  <TITLE>Calend�rio</TITLE>
  <link rel="stylesheet" href="bkb_ib.css" type="text/css">
</HEAD>
<SCRIPT LANGUAGE="JavaScript">
calDateFormat    = "DD/MM/YY";
topBackground    = "#F3F6F7";
bottomBackground = "#F3F6F7";
tableBGColor     = "#F3F6F7";
cellColor        = "#F3F6F7";
headingCellColor = "#DDE8EB";
headingTextColor = "#394F9A";
dateColor        = "#394F9A";
nodateColor      = "#C0C0C0";
focusColor       = "#CC0000";
hoverColor       = "#0000FF";
headingFontStyle = "bold 8pt verdana, arial";
bottomBorder  = false;
tableBorder   = 0;
var isNav = false;
var isIE  = false;
var calDateField2 = '';
var totos;
var today;
if (navigator.appName == "Netscape") {
    isNav = true;
    fontStyle        = "8pt verdana, arial";
}
else {
    isIE = true;
    fontStyle        = "8pt verdana, arial";
}
calendarStyle =
        '<STYLE type="text/css">' +
        '<!-- '+
        'TD.heading2 { text-decoration: none; color:' + dateColor + '; font: ' + fontStyle + '; }' +
        'TD.heading { text-decoration: none; border-bottom: 1px solid #FFFFFF; color:' + dateColor + '; font: ' + headingFontStyle + '; }' +
        'TD.heading3 { text-decoration: none; color:' + nodateColor + '; font: ' + fontStyle + '; }' +
        'A.focusDay:link { color: ' + focusColor + '; text-decoration: none; font: ' + fontStyle + '; }' +
        'A.focusDay:hover { color: ' + focusColor + '; text-decoration: underline; font: ' + fontStyle + '; }' +
        'A.focusDay:visited { color: ' + focusColor + '; text-decoration: none; font: ' + fontStyle + '; }' +
        'A.weekday:link { color: ' + dateColor + '; text-decoration: none; font: ' + fontStyle + '; }' +
        'A.weekday:hover { color: ' + hoverColor + '; font: ' + fontStyle + '; }' +
        'A.weekday:visited { color: ' + dateColor + '; text-decoration: none; font: ' + fontStyle + '; }' +
        '-->' +
        '</STYLE>';
document.write(calendarStyle);
dt_inicial = new Date(<%=str_DATA_INICIAL%>);
dt_final = new Date(<%=str_DATA_FINAL%>);
today = new Date(<%=str_DIA_HOJE%>);
menorigual = '<%=str_MENORIGUAL%>';
maiorigual = '<%=str_MAIORIGUAL%>';
//menorigual = 'MAIORIGUAL';
//menorigual = 'TODOS';
//maiorigual = 'MENORIGUAL';
//maiorigual = 'TODOS';
feriados = '<%=str_FERIADOS%>';
finalsemana = '<%=str_FINALSEMANA%>';
limiteano = '';
jNAO_HABILITAR = '<%=str_NAO_HABILITAR%>';
//today = new Date(2005,11,14);
//maiormenorigual = '';
//maiormenorigual = 'MAIORIGUAL';
//maiormenorigual = 'MENORIGUAL';
//feriados = 'TRUE';
//finalsemana = 'TRUE';
//limiteano = '';
//today = new Date(2005,10,14);
//maiormenorigual = 'MAIORIGUAL';
//feriados = 'FALSE';
//finalsemana = 'FALSE';
//limiteano = '';
auxres = '';
if (auxres != '')
{
	if (auxres < 1 || auxres > 12)
		intThisMonth = today.getMonth();
	else
		intThisMonth = auxres - 1;
}
else
	intThisMonth = today.getMonth();
auxres = '';
if (auxres != '')
	intThisYear = auxres;
else
	intThisYear = today.getFullYear();;
selectedLanguage = navigator.language;
buildCalParts();
setDateField();
function setDateField(){
    setInitialDate();
    calDocTop = buildTopCalFrame();
    calDocBottom = buildBottomCalFrame();
}
function setDateField2(dateField2){
    calDateField2 = dateField2;
}
function setInitialDate(){
    calDate = new Date(2005,01,01);
    atualDay = today.getDate();
    atualMonth = today.getMonth();
    atualYear = today.getFullYear();
    
    calDate.setMonth(intThisMonth);
	calDate.setFullYear(intThisYear);
	
    calDay  = calDate.getDate();
    calMonth = calDate.getMonth();
    calDate.setDate(1);
}
function buildTopCalFrame(){
		var calDoc =
        '<table border=0 cellspacing=0 cellpadding=1 width=100% bgcolor="#DDE8EB">' +
		'<FORM NAME="calControl" onSubmit="parent.setYear();return false;">' +
		'<tr><td>' +
        '<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0 align=center width="100%" bgcolor="#F3F6F7">' +
        '<TR>';
		calDoc +=         
				'<td align="right" width="20"><a href="javascript:;" onClick="parent.setPreviousMonth()" NAME="previousMonth">' +  '<img src="bt_seta_calendario_esq.gif" border="0" title="M�s anterior"></a></td>';
		calDoc += 
		'<td align="center">' + getMonthSelect() + '</td><td align="center" width="50"><INPUT TYPE="TEXT" NAME="year" VALUE="' + calDate.getFullYear() + '" SIZE="3" MAXLENGTH="4" class="forms" style="width:40px"></td>';
		calDoc +=		
				'<td width="20"><a href="javascript:;" NAME="nextMonth" onClick="parent.setNextMonth()">' +
				'<img src="bt_seta_calendario_dir.gif" border="0" title="M�s seguinte"></a></td>';
		calDoc += 
        '</TR>' +       
        '</TABLE>' +
        '</TD>' +
        '</TR>' +
        '</FORM>' +
        '</TABLE>'
    return calDoc;
}
function buildBottomCalFrame(){
    var calDoc = calendarBegin;
	//alert('buildBottomCalFrame')
	
	//alert('calDate 11 = ' + calDate)
	
    month   = calDate.getMonth();
    year    = calDate.getFullYear();
    day     = calDay;                                  
    var i   = 0;
    var days = getDaysInMonth();
    if (day > days){
        day = days;
    }
    var firstOfMonth = new Date (year, month, 1);
	//alert('firstOfMonth = ' + firstOfMonth)
	
    var startingPos  = firstOfMonth.getDay();
	//alert('startingPos = ' + startingPos)

    days += startingPos;
    var columnCount = 0;
    for (i = 0; i < startingPos; i++){
		//alert('i = ' + i)
		//alert('startingPos = ' + startingPos)
		//alert('calDoc = ' + calDoc)

        calDoc += blankCell;
		columnCount++;
    }

	//alert('calDoc = ' + calDoc)

    var currentDay = 0;
    var currentMonth = 0;
    var dayType    = "weekday";
    for (i = startingPos; i < days; i++){
		var paddingChar = "&nbsp;";
        if (i-startingPos+1 < 10){
            padding = "&nbsp;&nbsp;";
        }
        else {
            padding = "&nbsp;";
        }
        currentDay = i-startingPos+1;
        if ((currentDay == atualDay) && (month == atualMonth)){
            dayType = "focusDay";
        }
        else {
            dayType = "weekDay";
        }
	var seldate = currentDay;
	var selmonth = month;
	var selyear = year;
   
    //nResultado = atualYear * 10000 + ( atualMonth + 1 ) * 100 + atualDay;
	
    iniDay = dt_inicial.getDate();
    iniMonth = dt_inicial.getMonth();
    iniYear = dt_inicial.getFullYear();
    finDay = dt_final.getDate();
    finMonth = dt_final.getMonth();
    finYear = dt_final.getFullYear();
    nResultado = iniYear * 10000 + ( iniMonth + 1 ) * 100 + iniDay;
	//alert('nResultado = ' + nResultado)
    fResultado = finYear * 10000 + ( finMonth + 1 ) * 100 + finDay;
	//alert('fResultado = ' + fResultado)
    sResultado = selyear * 10000 + ( selmonth + 1 ) * 100 + seldate;
	//alert('sResultado = ' + sResultado)
    
    exibirData = "SIM";
	if (menorigual != "TODOS")
	{
		if (sResultado < nResultado && menorigual == "MAIORIGUAL")
			exibirData = "N�O";
			//alert('menor MAIORIGUAL')
		else
		{
			if (sResultado > nResultado && menorigual == "MENORIGUAL")
				{
				exibirData = "N�O";
				//alert('menor MENORIGUAL')
				}
			else
			{
				if (sResultado <= nResultado && menorigual == "MAIOR")
					exibirData = "N�O";
				else
				{
					if (sResultado >= nResultado && menorigual == "MENOR")
						exibirData = "N�O";
					else
					{
						if (sResultado != nResultado && menorigual == "IGUAL")
							exibirData = "N�O";
					}
				}
			}
		}
	}
	
	if (maiorigual != "TODOS")
	{
		if (sResultado < fResultado && maiorigual == "MAIORIGUAL")
			exibirData = "N�O";
		else
		{
			if (sResultado > fResultado && maiorigual == "MENORIGUAL")
				exibirData = "N�O";
			else
			{
				if (sResultado <= fResultado && maiorigual == "MAIOR")
					exibirData = "N�O";
				else
				{
					if (sResultado >= fResultado && maiorigual == "MENOR")
						exibirData = "N�O";
					else
					{
						if (sResultado != fResultado && maiorigual == "IGUAL")
							exibirData = "N�O";
					}
				}
			}
	
			if (limiteano == "")
			{
				anteriorData = (atualYear - 1) * 10000 + ( atualMonth + 1 ) * 100 + atualDay;
				posteriorData = (atualYear + 1) * 10000 + ( atualMonth + 1 ) * 100 + atualDay;
			}
			else
			{
				anteriorData = (atualYear - limiteano) * 10000 + ( atualMonth + 1 ) * 100 + atualDay;
				posteriorData = (atualYear + parseInt(limiteano)) * 10000 + ( atualMonth + 1 ) * 100 + atualDay;
			}
			if (sResultado > posteriorData || sResultado < anteriorData)
				exibirData = "N�O";
		}
	}
	
	if (feriados == "FALSE")
	{
		var arrayFeriadosMoveis = new Array();
										
		// Feriados M�veis
		arrayFeriadosMoveis[0] = "7/3/2000";
		arrayFeriadosMoveis[1] = "21/4/2000";
		arrayFeriadosMoveis[2] = "22/6/2000";
										
		arrayFeriadosMoveis[3] = "26/2/2001";
		arrayFeriadosMoveis[4] = "27/2/2001";
		arrayFeriadosMoveis[5] = "13/4/2001";
		arrayFeriadosMoveis[6] = "14/6/2001";
										
		arrayFeriadosMoveis[7] = "11/2/2002";
		arrayFeriadosMoveis[8] = "12/2/2002";
		arrayFeriadosMoveis[9] = "29/3/2002";
		arrayFeriadosMoveis[10] = "30/5/2002";
		
		arrayFeriadosMoveis[11] = "3/3/2003";
		arrayFeriadosMoveis[12] = "4/3/2003";
		
		arrayFeriadosMoveis[13] = "23/2/2004";
		arrayFeriadosMoveis[14] = "24/2/2004";
		
		arrayFeriadosMoveis[15] = "18/4/2003";
										
		var arrayFeriadosFixos = new Array();
										
		// Feriados Fixos
		arrayFeriadosFixos[0] = "1/1";
		arrayFeriadosFixos[1] = "21/4";
		arrayFeriadosFixos[2] = "1/5";
		arrayFeriadosFixos[3] = "7/9";
		arrayFeriadosFixos[4] = "12/10";
		arrayFeriadosFixos[5] = "2/11";
		arrayFeriadosFixos[6] = "15/11";
		arrayFeriadosFixos[7] = "25/12";
		arrayFeriadosFixos[8] = "31/12";
		for (count = 0; count <= 15; count++)
		{if ((currentDay + "/" + (month + 1) + "/" + year) == arrayFeriadosMoveis[count])exibirData = "N�O";}
									
		if (exibirData != "N�O"){				
			for (count = 0; count <= 8; count++)
			{if ((currentDay + "/" + (month + 1)) == arrayFeriadosFixos[count])exibirData = "N�O";}}
	}
							
	if (finalsemana == "FALSE")
	{
		var dtfs = new Date (year, month, currentDay);
		if (dtfs.getDay() == 0 || dtfs.getDay() == 6)
			exibirData = "N�O";
	}
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//'''''  Trata data n�ohabilitada
	if(jNAO_HABILITAR != ""){
		//alert(jNAO_HABILITAR)
		//alert(jNAO_HABILITAR.indexOf(currentDay))
		
		//alert(makeTwoDigit(currentDay) + '/' + makeTwoDigit(month + 1) + '/' + makeTwoDigit(year))
		if(jNAO_HABILITAR.indexOf(makeTwoDigit(currentDay) + '/' + makeTwoDigit(month + 1) + '/' + makeTwoDigit(year)) > -1){		
			exibirData = "N�O"
		}
	}
//'''''  fim Trata data n�ohabilitada
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	if (exibirData == "N�O"){
		calDoc += '<TD class="heading3" align=center>' +
		      padding + currentDay + paddingChar + '</TD>';
		//alert('currentDay 1 = ' +currentDay )	  
	}else{			
		calDoc += '<TD class="heading2" align=center>' +
		          '<a class="' + dayType + '" href="javascript:parent.returnDate(' + 
		          currentDay + ')">' + padding + currentDay + paddingChar + '</a></TD>';
		//alert('currentDay 2 = ' +currentDay )	 
		}
	//alert('calDoc =' + calDoc)
	columnCount++;
    if (columnCount % 7 == 0){
        calDoc += '</TR><TR>';
    }
    }
    for (i=days; i<42; i++){
        calDoc += blankCell;
	columnCount++;
        if (columnCount % 7 == 0){
            calDoc += '</TR>';
            if (i<41){
                calDoc += '<TR>';
            }
        }
    }
    calDoc += calendarEnd;
	//alert('calDoc =' + calDoc)
    return calDoc;
}
function writeCalendar(){    
    calDocBottom = buildBottomCalFrame();
	//alert('calDocBottom = ' + calDocBottom)
	MM_setTextOfLayer('div1','',calDocBottom);
}
function setToday(){
    calDate = new Date();
    var month = calDate.getMonth();
    var year  = calDate.getFullYear();
    document.calControl.month.selectedIndex = month;
    document.calControl.year.value = year;
    writeCalendar();
}
function setYear(){
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)){
        calDate.setFullYear(year);
        writeCalendar();
    }
    else {
        document.calControl.year.focus();
        document.calControl.year.select();
    }
}
function setCurrentMonth(){
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)) {
        calDate.setFullYear(year);
    }
    
    var month = document.calControl.month.selectedIndex;
    calDate.setMonth(month);
    writeCalendar();
}
function setPreviousYear() {
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year) && year > 1000){
        year--;
        calDate.setFullYear(year);
        document.calControl.year.value = year;
        writeCalendar();
    }
}
function setPreviousMonth(){
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)){
        var month = document.calControl.month.selectedIndex;
        if (month == 0){
            month = 11;
            if (year > 1000){
                year--;
                calDate.setFullYear(year);
                document.calControl.year.value = year;
            }
        }
        else {
            month--;
        }
		//alert('aqui')
        calDate.setMonth(month);
		//alert('calDate = ' + calDate)
        document.calControl.month.selectedIndex = month;
        writeCalendar();
    }
}
function setNextMonth(){
    var year = document.calControl.year.value;
    if (isFourDigitYear(year)){
        var month = document.calControl.month.selectedIndex;
        if (month == 11){
            month = 0;
            year++;
            calDate.setFullYear(year);
            document.calControl.year.value = year;
        }
        else {
            month++;
        }
        calDate.setMonth(month);
        document.calControl.month.selectedIndex = month;
        writeCalendar();
    }
}
function setNextYear(){
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)){
        year++;
        calDate.setFullYear(year);
        document.calControl.year.value = year;
        writeCalendar();
    }
}
function getDaysInMonth(){
    var days;
    var month = calDate.getMonth()+1;
    var year  = calDate.getFullYear();
    if (month==1 || month==3 || month==5 || month==7 || month==8 ||
        month==10 || month==12)  {
        days=31;
    }
    else if (month==4 || month==6 || month==9 || month==11) {
        days=30;
    }
    else if (month==2){
        if (isLeapYear(year)){
            days=29;
        }
        else {
            days=28;
        }
    }
    return (days);
}
function isLeapYear (Year){
    if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0)){
        return (true);
    }
    else {
        return (false);
    }
}
function isFourDigitYear(year){
    if (year.length != 4){
        document.calControl.year.value = calDate.getFullYear();
        document.calControl.year.select();
        document.calControl.year.focus();
    }
    else {
        return true;
    }
}
function getMonthSelect(){
        monthArray = new Array('Janeiro', 'Fevereiro', 'Mar�o', 'Abril', 'Maio', 'Junho',
                               'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro');
    var activeMonth = calDate.getMonth();
    monthSelect = '<SELECT NAME="month" onChange="setCurrentMonth()" class="combo">';
    for (i in monthArray){
        if (i == activeMonth){
            monthSelect += '<OPTION SELECTED>' + monthArray[i];
        }
        else {
            monthSelect += '<OPTION>' + monthArray[i];
        }
    }
    monthSelect += '</SELECT><img src="xin.gif" border="0" width="10">';
    return monthSelect;
}
function createWeekdayList(){
    weekdayList  = new Array('S�bado', 'Domingo', 'Segunda-feira', 'Ter�a-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira')
    weekdayArray = new Array('D', 'S', 'T', 'Q', 'Q', 'S', 'S');
    var weekdays = "<TR BGCOLOR='" + headingCellColor + "'>";
    for (i in weekdayArray) {
        weekdays += "<TD class='heading' align=center>" + weekdayArray[i] + "</TD>";
    }
    weekdays += "</TR>";
    return weekdays;
}
function buildCalParts(){
    weekdays = createWeekdayList();
    blankCell = '<TD align=center>&nbsp;&nbsp;&nbsp;</TD>';
    calendarBegin =
        '<table border=0 cellspacing=0 cellpadding=1 align=center width=100% bgcolor="#DDE8EB">' +
		'<tr><td>';        
        if (isNav){
            calendarBegin += 
                '<TABLE CELLPADDING="0" CELLSPACING="0" BORDER=' + tableBorder + ' ALIGN=CENTER width="100%" BGCOLOR="' + tableBGColor + '"><TR><TD>';
        }
        calendarBegin +=
            "<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=" + tableBorder + " ALIGN=CENTER width=100% BGCOLOR=" + tableBGColor + ">" +
            weekdays +
            "<TR>";
	    calendarEnd = "";
        if (bottomBorder){
            calendarEnd += "<TR></TR>";
        }
        if (isNav){
            calendarEnd += "</TD></TR></TABLE>";
        }
        calendarEnd +=
			'<table border=0 cellspacing=0 cellpadding=0 align=center width="100%" bgcolor="' + tableBGColor + '">' +
			'<TR><TD>' +
			'<table border=0 cellspacing=0 cellpadding=0 align=center width="100%" bgcolor="' + tableBGColor + '">' +
        	'<tr>' +
          	'<TD align="center"><a href="javascript:;" onClick="window.close()">Fechar</a></td>' +
          	'<TD align="center"><a href="javascript:;" onClick="parent.returnDate(0)">Limpar campo</a></td>' +
        	'</tr>' +
      		'</table>' +
			'</TD></TR>' +
			'</TABLE>'
			
			
}
function jsReplace(inString, find, replace){
    var outString = "";
    if (!inString){
        return "";
    }
    if (inString.indexOf(find) != -1){
        t = inString.split(find);
        return (t.join(replace));
    }
    else {
        return inString;
    }
}
function doNothing(){
}
function makeTwoDigit(inValue){
    var numVal = parseInt(inValue, 10);
    if (numVal < 10){
        return("0" + numVal);
    }
    else {
        return numVal;
    }
}
function returnDate(inDay)
{
	calDate.setDate(inDay);
	if(!inDay==0)
	{
		calDate.setDate(inDay);
		var day           = calDate.getDate();
		var month         = calDate.getMonth()+1;
		var year          = calDate.getFullYear();
		var monthString   = monthArray[calDate.getMonth()];
		var monthAbbrev   = monthString.substring(0,3);
		var weekday       = weekdayList[calDate.getDay()];
		var weekdayAbbrev = weekday.substring(0,3);
		outDate = calDateFormat;
		if (calDateFormat.indexOf("DD") != -1){
			day = makeTwoDigit(day);
			outDate = jsReplace(outDate, "DD", day);
		}
		else if (calDateFormat.indexOf("dd") != -1){
			outDate = jsReplace(outDate, "dd", day);
		}
		if (calDateFormat.indexOf("MM") != -1){
			month = makeTwoDigit(month);
			outDate = jsReplace(outDate, "MM", month);
		}
		else if (calDateFormat.indexOf("mm") != -1){
			outDate = jsReplace(outDate, "mm", month);
		}
		if (calDateFormat.indexOf("yyyy") != -1){
			outDate = jsReplace(outDate, "yyyy", year);
		}
		else if (calDateFormat.indexOf("yy") != -1){
			var yearString = "" + year;
			var yearString = yearString.substring(2,4);
			outDate = jsReplace(outDate, "yy", yearString);
		}
		else if (calDateFormat.indexOf("YY") != -1){
			outDate = jsReplace(outDate, "YY", year);
		}
		if (calDateFormat.indexOf("Month") != -1){
			outDate = jsReplace(outDate, "Month", monthString);
		}
		else if (calDateFormat.indexOf("month") != -1){
			outDate = jsReplace(outDate, "month", monthString.toLowerCase());
		}
		else if (calDateFormat.indexOf("MONTH") != -1){
			outDate = jsReplace(outDate, "MONTH", monthString.toUpperCase());
		}
		if (calDateFormat.indexOf("Mon") != -1){
			outDate = jsReplace(outDate, "Mon", monthAbbrev);
		}
		else if (calDateFormat.indexOf("mon") != -1){
			outDate = jsReplace(outDate, "mon", monthAbbrev.toLowerCase());
		}
		else if (calDateFormat.indexOf("MON") != -1){
			outDate = jsReplace(outDate, "MON", monthAbbrev.toUpperCase());
		}
		if (calDateFormat.indexOf("Weekday") != -1){
			outDate = jsReplace(outDate, "Weekday", weekday);
		}
		else if (calDateFormat.indexOf("weekday") != -1){
			outDate = jsReplace(outDate, "weekday", weekday.toLowerCase());
		}
		else if (calDateFormat.indexOf("WEEKDAY") != -1){
			outDate = jsReplace(outDate, "WEEKDAY", weekday.toUpperCase());
		}
		if (calDateFormat.indexOf("Wkdy") != -1){
			outDate = jsReplace(outDate, "Wkdy", weekdayAbbrev);
		}
		else if (calDateFormat.indexOf("wkdy") != -1){
			outDate = jsReplace(outDate, "wkdy", weekdayAbbrev.toLowerCase());
		}
		else if (calDateFormat.indexOf("WKDY") != -1){
			outDate = jsReplace(outDate, "WKDY", weekdayAbbrev.toUpperCase());
		}
	}
	else
		{
		outDate = ""
		}
    window.opener.carrega_valor_calendario(outDate);
	self.close();
}
function MM_findObj(n, d) { 
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}
function MM_setTextOfLayer(objName,x,newText) { 
	var obj = document.getElementById(objName);
	obj.innerHTML = unescape(newText); 	
}
</script>
<body leftmargin=0 topmargin=1 marginwidth=0 marginheight=0>
<script>
document.write(calDocTop)
if (document.all)
	document.write('<div id="div1" style="position:absolute;top:30px;left:1px;width:100%;background-color:#ffffff;padding:0px 0px;border: 0px solid black;">' + calDocBottom + '</div>');
else
	document.write('<layer id="div1" left="1" width="175" top="30" BGCOLOR="#ffffff" visibility="visible">' + calDocBottom + '</layer>');
</script>
</body>
</HTML>