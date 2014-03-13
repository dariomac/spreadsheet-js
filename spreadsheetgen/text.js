String.prototype.isEmailAddress = function(){
    var str = this;
    if(str == null)
        return false;

	var at='@';
    var dot='.';
    var lat=str.indexOf(at);
    var lstr=str.length;
    var ldot=str.indexOf(dot);

    if (lstr == 0)
        return false;

    if (str.indexOf(at)==-1 || str.indexOf(at)==0 || str.indexOf(at)==lstr)
        return false;

    if (str.indexOf(dot)==-1 || str.indexOf(dot)==0 || str.indexOf(dot)==lstr)
        return false;

    if (str.indexOf(at,(lat+1))!=-1)
        return false;

    if (str.substring(lat-1,lat)==dot || str.substring(lat+1,lat+2)==dot)
        return false;

    if (str.indexOf(dot,(lat+2))==-1)
        return false;

    if (str.indexOf(' ')!=-1)
        return false;

    return true;
}

String.prototype.isNumeric = function(){
    var sText = this;
	var validChars = '0123456789,-';
	var isNumber = true;
	var sChar;

	if(sText.length == 0) isNumber = false;

	for(var i=0;i<sText.length && isNumber==true;i++){
		sChar = sText.charAt(i);
		if(validChars.indexOf(sChar)==-1)
			isNumber = false;
	}
	if(isNumber) isNumber = !isNaN(parseFloat(sText));

	return isNumber;
}

//Usage: 'Hello. My name is {0} {1}.'.format('Dario', 'Macchi');
String.prototype.format = function(){
	var str = this;

    for(var i=0;i<arguments.length;i++){
        var re = new RegExp('\\{' + (i) + '\\}','gm');
        str = str.replace(re, arguments[i]);
    }
    return str;
}

//Usage: String.format('Hello. My name is {0} {1}.', 'Dario', 'Macchi');
String.format = function(){
    if( arguments.length == 0 )
        return null;

    var str = arguments[0];

    for(var i=1;i<arguments.length;i++){
        var re = new RegExp('\\{' + (i-1) + '\\}','gm');
        str = str.replace(re, arguments[i]);
    }

    return str;
}

function deleteSpaces(pValue)
{
	while(''+pValor.charAt(0)==' '){
		pValue=pValue.substring(1,pValue.length);
	}
	return (pValue);
}
