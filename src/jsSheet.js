var gEVENT = new Event();

//
//
// .----------------------.
// |SHEET CLASS DEFINITION|
// |                      |
// ·----------------------·
//

function Sheet(RowCount, ColCount, updateDependentCells){
	var _rowCount = RowCount;
	var _colCount = ColCount;
	var _sheet = new Array(RowCount+1);
	//almacena las referencias dependientes de una celda precedente. Si F3=C3*B3, F3 es dependiente de C3 y B3. Estas son precedentes de F3.
	var _precedents = new Hashtable();

	var _cellRefRegExp = /[A-Z]{1,2}[1-9]+[0-9]*(?!\()/gm
	var _rangeRefRegExp = /[A-Z]{1,2}[1-9]+[0-9]*\:[A-Z]{1,2}[1-9]+[0-9]*/gm

	var _isLayoutSuspended = false;
	var _cellsToUpdateAfterResume = new Hashtable();

	var _updateDependentMode = (typeof updateDependentCells == 'undefined')?true:updateDependentCells;

	var _me;

	for(var i=1; i<=RowCount; i++){
		_sheet[i] = new Array(ColCount);

		for (var j=1; j<=ColCount; j++){
			_sheet[i][j]=new Cell();
			_sheet[i][j].ColRef = SheetUtil.GetColumnName(j);
			_sheet[i][j].RowRef = i;
		}
	}

	this.SuspendLayout = function(){
		_isLayoutSuspended = true;
	}

	this.ResumeLayout = function(){
		_isLayoutSuspended = false;

		_cellsToUpdateAfterResume.moveFirst();

		while(_cellsToUpdateAfterResume.next()){
			var cell = _cellsToUpdateAfterResume.getValue();
			OnCellUpdate(cell);
		}

		_cellsToUpdateAfterResume = new Hashtable();
	}

	this.GetValue = function(cellRefStr){
		var cellRef = SplitReference(cellRefStr);

		return _sheet[cellRef['row']][cellRef['col']].Value;
	}

	this.GetFormula = function(cellRefStr){
		var cellRef = SplitReference(cellRefStr);

		return _sheet[cellRef['row']][cellRef['col']].Formula;
	}

	this.SetValue = function(cellRefStr, value, state){
		var cellRef = SplitReference(cellRefStr);

		var cell = _sheet[cellRef['row']][cellRef['col']]

		cell.State = (typeof state == 'undefined')?'':state;
		SetCellValue(cell, value);

		cell.Formula = value.toString().toUpperCase();

		if (CheckFormula(cell))
			SetCellValue(cell, EvaluateFormula(cell));

		if (_updateDependentMode)
			UpdateDependentCells(cell);
	}

	this.UpdateValue = function(cellRefStr){
		var cellRef = SplitReference(cellRefStr);
		var cell = _sheet[cellRef['row']][cellRef['col']]

		SetCellValue(cell, EvaluateFormula(cell));
		UpdateDependentCells(cell);
	}

	this.CellUpdated = function(eventHandlerFunc){
		gEVENT.addListener(this,'cell_updated',eventHandlerFunc);
		_me = this;
	}

	function OnCellUpdate(cell){
		if (cell.State == 'hidden') return;

		gEVENT.fireEvent(null,_me,'cell_updated',cell.Ref() + '=' + cell.Value);
	}

	this.ErrorThrowed = function(eventHandlerFunc){
		gEVENT.addListener(this,'error_throwed',eventHandlerFunc);
		_me = this;
	}

	function OnErrorThrowed(cellRef, errMsg){
		var msg = cellRef + ':' + errMsg;
		gEVENT.fireEvent(null,_me,'error_throwed',msg);
	}

	function SplitReference(ref){
		var iRef = new Array(2);
		iRef['row'] = 0; //indices
		iRef['col'] = 0; //	invalidos

		var first = ref.charAt(0);

		//nos aseguramos que referencia comience con una letra
		if(!isNaN(first)) return;

		var last = ref.charAt(1);

		if(isNaN(last)){
			iRef['row'] = ref.substr(2, ref.length);
			iRef['col'] = SheetUtil.GetColumnIndex(first+last);
		}
		else{
			iRef['row'] = ref.substr(1, ref.length);
			iRef['col'] = SheetUtil.GetColumnIndex(first);
		}

		return iRef;
	}

	function CheckFormula(cell){

		if (cell.Formula.toString().indexOf('\\', 0) != -1){
			SetCellValue(cell, cell.Value.toString().replace('\\', ''));
			cell.Formula = cell.Formula.replace('\\', '');

			OnErrorThrowed(cell.Ref(), 'Character \\ es illegal. You can\'t use it as a value of a cell.')
			return false;
		}

		if (cell.Formula.charAt(0) != '=')
			if (isNaN(cell.Formula) ){
				OnErrorThrowed(cell.Ref(), 'String values are not allowed as cell values.');
				return false;
			}

		return true;
	}

	function UpdateDependentCells(cell){
		//obtengo las celdas dependientes de cell (o sea, celdas para las que cell sea precedente)
		var cellDependents = _precedents.get(cell.Ref());
		if (typeof cellDependents != 'undefined'){
			cellDependents.moveFirst();

			while(cellDependents.next()){
				var dependentRef = SplitReference(cellDependents.getValue());
				var dependentCell = _sheet[dependentRef['row']][dependentRef['col']];

				SetCellValue(dependentCell, EvaluateFormula(dependentCell));
			}

		}
	}

	function EvaluateFormula(cell){
		with(ExcelOp){
			var expandedFormula = ExpandFormula(cell, cell);
			if (expandedFormula == '<<circular>>'){
				OnErrorThrowed(cell.Ref(), 'Circular reference founded in "' + cell.Ref() + '" expansion.');
				return 0;
			}
			else
				return eval(expandedFormula);
		}
	}

	function ExpandFormula(startCell, nextRefCell){
		if (!isNaN(nextRefCell.Formula))
			return nextRefCell.Formula;

		var formulaExpanded = nextRefCell.Formula;

		var ranges = formulaExpanded.match(_rangeRefRegExp);
		if (ranges != null)
			formulaExpanded = ExpandRange(formulaExpanded, ranges);

		var refs = formulaExpanded.match(_cellRefRegExp);

		if (refs != null){
			for(var i=0; i<refs.length; i++){

				AddToPrecedents(startCell.Ref(), refs[i]);

				var cellRef = SplitReference(refs[i]);
				var nextCell = _sheet[cellRef['row']][cellRef['col']];

				if(startCell.IsEqual(nextCell)){
					return '<<circular>>';
					break;
				}

				formulaExpanded = formulaExpanded.replace(nextCell.Ref(), ExpandFormula(startCell, nextCell));
			}
		}

		if (formulaExpanded.indexOf('<<circular>>') != -1)
			return '<<circular>>';
		else
			return formulaExpanded.toString().replace('=','');
	}

	function ExpandRange(formulaToExpand, ranges){
		for(var i=0; i<ranges.length; i++){
			var range = ranges[i].split(':');

			var startCellRef = SplitReference(range[0]);
			var endCellRef = SplitReference(range[1]);

			var startRow, endRow;
			if(startCellRef['row'] <= endCellRef['row']){
				startRow = startCellRef['row']; endRow = endCellRef['row'];
			}
			else{
				startRow = endCellRef['row']; endRow = startCellRef['row'];
			}

			var startCol, endCol;
			if(startCellRef['col'] <= endCellRef['col']){
				startCol = startCellRef['col']; endCol = endCellRef['col'];
			}
			else{
				startCol = endCellRef['col']; endCol = startCellRef['col'];
			}

			var rangeExpansion = '';
			for(var row = startRow; row<=endRow; row++){

				for(var col = startCol; col<=endCol; col++){
					if (rangeExpansion != '') rangeExpansion += ', ';

					rangeExpansion += SheetUtil.GetColumnName(col) + row.toString();
				}
			}
			formulaToExpand = formulaToExpand.replace(ranges[i], rangeExpansion);
		}
		return formulaToExpand;
	}

	function FormulaReplace(formula, oldValue, newValue){
		return eval(formula + '.replace(/' + oldValue + '/g, ' + newValue + ')');
	}

	function AddToPrecedents(dependent, precedent){
		if(typeof _precedents.get(precedent) == 'undefined'){
			_precedents.put(precedent, new Hashtable());
		}

		_precedents.get(precedent).put(dependent, dependent);
	}

	function SetCellValue(cell, value){
		if (cell.Value == value) return;
		cell.Value = value;

		if (_isLayoutSuspended)
			_cellsToUpdateAfterResume.put(cell.Ref(), cell);
		else
			OnCellUpdate(cell);
	}
}

//
//
// .--------------------------.
// |SHEETUTIL CLASS DEFINITION|
// |                          |
// ·--------------------------·
//

function SheetUtil(){ }

SheetUtil.Letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

SheetUtil.GetColumnIndex = function(colNam){
	if(colNam.length == 1){
		return SheetUtil.Letters.indexOf(colNam)+1;
	}
	else{
		first = SheetUtil.Letters.indexOf(colNam.charAt(0))+1;
		return (SheetUtil.Letters.indexOf(colNam.charAt(1))+1) + (SheetUtil.Letters.length*first);
	}
}

SheetUtil.GetColumnName = function(colIdx){
	if(colIdx == 0) return;

	var colNam = '';

	if(colIdx <=SheetUtil.Letters.length)
		colNam = SheetUtil.Letters.charAt(colIdx-1);
	else{
		var first = Math.round(colIdx/SheetUtil.Letters.length,0)-1;
		if (first < 0) first = 0;

		var last = (colIdx%SheetUtil.Letters.length)-1;
		if (last < 0) last = 0;

		colNam = SheetUtil.Letters.charAt(first)+SheetUtil.Letters.charAt(last);
	}

	return colNam.toUpperCase();
}

//
//
// .----------------------.
// |CELL CLASS DEFINITION |
// |                      |
// ·----------------------·
//

function Cell(){
	this.Value=0;
	this.Formula='0';
	this.ColRef='';
	this.RowRef=0;
	this.State='';

	function Pub_Ref(){
		return this.ColRef.toString() + this.RowRef.toString();
	}

	function Pub_IsEqual(cell2){
	    if (cell2 == null || typeof cell2 == 'undefined') return false;

		return this.RowRef==cell2.RowRef && this.ColRef==cell2.ColRef;
	}

	this.Ref = Pub_Ref;
	this.IsEqual = Pub_IsEqual;
}

//
//
// .--------------------------.
// |HASHTABLE CLASS DEFINITION|
// |                          |
// ·--------------------------·
//

function Hashtable(){
	this.hash = new Array();
	this.keys = new Array();

	this.location = 0;
}

Hashtable.prototype.put = function (key, value){
	if (value == null)
		return;

	if (this.hash[key] == null)
		this.keys[this.keys.length] = key;

	this.hash[key] = value;
}

Hashtable.prototype.get = function (key){
	return this.hash[key];
}

Hashtable.prototype.size = function (){
    return this.keys.length;
}

Hashtable.prototype.next = function (){
	if (++this.location < this.keys.length)
		return true;
	else
		return false;
}

Hashtable.prototype.moveFirst = function (){
	try {
		this.location = -1;
	} catch(e) {/*//do nothing here :-)*/}
}

Hashtable.prototype.getValue = function (){
	try {
		return this.hash[this.keys[this.location]];
	} catch(e) {
		return null;
	}
}

//
//
// .-------------.
// |Event Handler|
// |             |
// ·-------------·
//


/**
 * Binds a function to the given object's scope
 *
 * @param {Object} object The object to bind the function to.
 * @return {Function}	Returns the function bound to the object's scope.
 */
Function.prototype.bind = function (object)
{
	var method = this;
	return function ()
	{
		return method.apply(object, arguments);
	};
};

/**
 * Create a new instance of Event.
 *
 * @classDescription	This class creates a new Event.
 * @return {Object}	Returns a new Event object.
 * @constructor
 */
function Event()
{
	this.events = [];
	this.builtinEvts = [];
}

/**
 * Gets the index of the given action for the element
 *
 * @memberOf Event
 * @param {Object} obj The element attached to the action.
 * @param {String} evt The name of the event.
 * @param {Function} action The action to execute upon the event firing.
 * @param {Object} binding The object to scope the action to.
 * @return {Number} Returns an integer.
 */
Event.prototype.getActionIdx = function(obj,evt,action,binding)
{
	if(obj && evt)
	{

		var curel = this.events[obj][evt];
		if(curel)
		{
			var len = curel.length;
			for(var i = len-1;i >= 0;i--)
			{
				if(curel[i].action == action && curel[i].binding == binding)
				{
					return i;
				}
			}
		}
		else
		{
			return -1;
		}
	}
	return -1;
};

/**
 * Adds a listener
 *
 * @memberOf Event
 * @param {Object} obj The element attached to the action.
 * @param {String} evt The name of the event.
 * @param {Function} action The action to execute upon the event firing.
 * @param {Object} binding The object to scope the action to.
 * @return {null} Returns null.
 */
Event.prototype.addListener = function(obj,evt,action,binding)
{
	if(this.events[obj])
	{
		if(this.events[obj][evt])
		{
			if(this.getActionIdx(obj,evt,action,binding) == -1)
			{
				var curevt = this.events[obj][evt];
				curevt[curevt.length] = {action:action,binding:binding};
			}
		}
		else
		{
			this.events[obj][evt] = [];
			this.events[obj][evt][0] = {action:action,binding:binding};
		}
	}
	else
	{
		this.events[obj] = [];
		this.events[obj][evt] = [];
		this.events[obj][evt][0] = {action:action,binding:binding};
	}
};

/**
 * Removes a listener
 *
 * @memberOf Event
 * @param {Object} obj The element attached to the action.
 * @param {String} evt The name of the event.
 * @param {Function} action The action to execute upon the event firing.
 * @param {Object} binding The object to scope the action to.
 * @return {null} Returns null.
 */
Event.prototype.removeListener = function(obj,evt,action,binding)
{
	if(this.events[obj])
	{
		if(this.events[obj][evt])
		{
			var idx = this.actionExists(obj,evt,action,binding);
			if(idx >= 0)
			{
				this.events[obj][evt].splice(idx,1);
			}
		}
	}
};

/**
 * Fires an event
 *
 * @memberOf Event
 * @param e [(event)] A builtin event passthrough
 * @param {Object} obj The element attached to the action.
 * @param {String} evt The name of the event.
 * @param {Object} args The argument attached to the event.
 * @return {null} Returns null.
 */
Event.prototype.fireEvent = function(e,obj,evt,args)
{
	if(!e){e = window.event;}

	if(obj && this.events)
	{
		var evtel = this.events[obj];
		if(evtel)
		{
			var curel = evtel[evt];
			if(curel)
			{
				for(var act in curel)
				{
					var action = curel[act].action;
					if(curel[act].binding)
					{
						action = action.bind(curel[act].binding);
					}
					if (action != null)
						action(e,args);
				}
			}
		}
	}
};

//
//
// .----------------.
// |Excel Operations|
// |                |
// ·----------------·
//
function ExcelOp(){ }

ExcelOp.ROUND = function (num, dec) {
	if (dec == null) dec=0;
	var result = Math.round(num*Math.pow(10,dec))/Math.pow(10,dec);
	return result;
}

ExcelOp.AVG = function(/*...*/){
	if (arguments.length == 0) return 0;
	if (arguments.length == 1 && !IsNaN(arguments[0]))
		return arguments[0];

	return this.SUM.apply(this, arguments)/arguments.length;
}

ExcelOp.SUM = function(/*...*/){
	if (arguments.length == 0) return 0;
	if (arguments.length == 1 && !IsNaN(arguments[0]))
		return arguments[0];

	var sum = 0;
	for(var i=0; i<arguments.length; i++){
		sum += arguments[i];
	}

	return sum;
}

ExcelOp.MAX = function(/*...*/){
	if (arguments.length == 0) return 0;
	if (arguments.length == 1 && !IsNaN(arguments[0]))
		return arguments[0];

	var max = -Number.MAX_VALUE;

	for(var i=0; i<arguments.length; i++){
		if (arguments[i] > max)
			max = arguments[i];
	}
	return max;
}

ExcelOp.MIN = function(/*...*/){
	if (arguments.length == 0) return 0;
	if (arguments.length == 1 && !IsNaN(arguments[0]))
		return arguments[0];

	var min = Number.MAX_VALUE;

	for(var i=0; i<arguments.length; i++){
		if (arguments[i] < min)
			min = arguments[i];
	}
	return min;
}

ExcelOp.COUNT = function(/*..*/){
	if (arguments.length == 0) return 0;
	if (arguments.length == 1 && !IsNaN(arguments[0]))
		return 1;

	return arguments.length;
}

ExcelOp.COUNT.IF = function(/*..*/){
	if (arguments.length == 0) return 0;
	if (arguments.length == 1 && !IsNaN(arguments[0]))
		return 0;

	var counter = 0;
	var condition = arguments[arguments.length-1];

	for(var i=0; i<arguments.length-1; i++){
		if ( eval(arguments[i] + condition) )
			counter++;
	}

	return counter;
}

ExcelOp.SUM.IF = function(/*..*/){
	if (arguments.length == 0) return 0;
	if (arguments.length == 1 && !IsNaN(arguments[0]))
		return 0;

	var sum = 0;
	var condition = arguments[arguments.length-1];

	for(var i=0; i<arguments.length-1; i++){
		if ( eval(arguments[i] + condition) )
			sum += arguments[i];
	}

	return sum;
}

ExcelOp.IF = function(condition, trueAction, falseAction){
	if(condition)
		return trueAction;
	else
		return falseAction;
}

ExcelOp.POW = function(value, pow){
	return Math.pow(value, pow);
}

ExcelOp.LOOKUP = function(search_value, comparingCells, resultCells){
	var comparingPlusResultLength = arguments.length-1;
	if ( this.ISODD((comparingPlusResultLength)/2) )
		OnErrorThrowed('Comparing range and result range MUST be of the same size.');
		//throw 'Comparing range and result range MUST be of the same size.';

	var positionInComparingRange = -1;

	for(var i=1; i<((arguments.length)/2); i++){
		if (arguments[0] == arguments[i]){
			positionInComparingRange = i;
			break;
		}
	}

	return arguments[(comparingPlusResultLength/2)+positionInComparingRange];
}

ExcelOp.ISEVEN = function(value){
	return !(value % 2);
}

ExcelOp.ISODD = function(value){
	return !this.ISEVEN(value);
}
