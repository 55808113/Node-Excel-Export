var sheetFront = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' 
		+ ' <x:sheetPr/><x:sheetViews><x:sheetView tabSelected="1" workbookViewId="0" /></x:sheetViews>' 
		+ ' <x:sheetFormatPr defaultRowHeight="15" />';
var sheetBack = ' <x:pageMargins left="0.75" right="0.75" top="0.75" bottom="0.5" header="0.5" footer="0.75" />'
		+ ' <x:headerFooter /></x:worksheet>';
    
var fs = require('fs');

function Sheet(config, xlsx, shareStrings, convertedShareStrings){
  this.config = config;
  this.xlsx = xlsx;
  this.shareStrings = shareStrings;
  this.convertedShareStrings = convertedShareStrings; 
}

Sheet.prototype.generate = function(){
  var config = this.config, xlsx = this.xlsx;
	var cols = config.cols,
	mergeFields = [],
	data = config.rows,
	colsLength = cols.length,
	rows = "",
	row = "",
	colsWidth = "",
	sheetmergeCells = [],
	styleIndex,
  self = this,
	k;
	config.fileName = 'xl/worksheets/' + (config.sheetname || "sheet").replace(/[*?\]\[\/\/]/g, '') + '.xml';
	if (config.stylesXmlFile) {
		var path = config.stylesXmlFile;
		var styles = null;
		styles = fs.readFileSync(path, 'utf8');
		if (styles) {
			xlsx.file("xl/styles.xml", styles);
		}
	}
	//开始的行号===============
	var START_ROW = 2
	//是否显示标题，显示标题开始行从3开始
	if (config.title){
		var titleRow = 1;
		//中间对齐
		var titleStyleIndex = 1
		row = '<x:row r="' + titleRow + '" spans="1:' + colsLength + '">';
		row += addStringCell(self, getColumnLetter(0 + 1) + titleRow, config.title, titleStyleIndex);
		row += '</x:row>';
		rows += row;
		//合并单元格,列大于1时才合并单元格
		if (colsLength>1){
			sheetmergeCells.push({
				startCell: getColumnLetter(0 + 1) + titleRow,
				endCell: getColumnLetter(colsLength) + titleRow
			})
		}
		START_ROW++
	}
	//如果合并字段存在，判断当前字段是合并字段
	if (config.mergeFields){
		for (i = 0; i < config.mergeFields.length; i++) {
			var mergeCell = config.mergeFields[i];
			var mergeField=undefined,premiseField=undefined;
			//通过合并的字段名称。转换为索引号。
			for (k = 0; k < colsLength; k++) {
				if (mergeCell.mergeField==cols[k].name){
					mergeField = k;
				}
				if (mergeCell.premiseField==cols[k].name){
					premiseField = k;
				}
				if (mergeField!=undefined&&premiseField!=undefined){
					mergeFields.push({
						mergeField: mergeField,
						premiseField: premiseField,
						perMergeValue:undefined,//前合并字段值
						perPremiseValue:undefined,//前主键字段值
						curMergeValue:undefined,//当前合并字段值
						curPremiseValue:undefined,//当前主键字段值
						span:1//跨度默认为1
					})
					break;
				}
			}

		}
	}
	//合并自定义的单元格
	if (config.mergeCells){
		for (i = 0; i < config.mergeCells.length; i++) {
			var mergeCell = config.mergeCells[i];
			var startCell = {
				column: getColumnLetter(mergeCell.startCell.column + 1),
				row: mergeCell.startCell.row + START_ROW
			};
			var endCell = {
				column: getColumnLetter(mergeCell.endCell.column + 1),
				row: mergeCell.endCell.row + START_ROW
			}
			sheetmergeCells.push({
				startCell:startCell.column + startCell.row,
				endCell:endCell.column + endCell.row
			})
		}
	}
	//first row for column caption
	row = '<x:row r="' + (START_ROW-1) + '" spans="1:' + colsLength + '">';
	var colStyleIndex;
	for (k = 0; k < colsLength; k++) {
		colStyleIndex = cols[k].captionStyleIndex || 0;
		row += addStringCell(self, getColumnLetter(k + 1) + 1, cols[k].caption, colStyleIndex);
		if (cols[k].width) {
			colsWidth += '<x:col customWidth = "1" width="' + cols[k].width + '" max="' + (k + 1) + '" min="' + (k + 1) + '"/>';
		}
	}
	row += '</x:row>';
	rows += row;

	//fill in data
	var i, j, r, cellData, currRow, cellType, dataLength = data.length;

	for (i = 0; i < dataLength; i++) {
		r = data[i],
		currRow = i + START_ROW;
		row = '<x:row r="' + currRow + '" spans="1:' + colsLength + '">';
		//合并相同字段时使用===============
		for (var l = 0; l < mergeFields.length; l++) {
			var mergeCell = mergeFields[l];
			var flag = true;
			mergeCell.curMergeValue = r[mergeCell.mergeField];
			if (mergeCell.premiseField!=undefined){
				mergeCell.curPremiseValue = r[mergeCell.premiseField];
				flag = false;
			}
			if (mergeCell.perMergeValue == mergeCell.curMergeValue && (flag || mergeCell.perPremiseValue == mergeCell.curPremiseValue)) {
				mergeCell.span += 1;
			} else {
				var columnLetter = getColumnLetter(mergeCell.mergeField + 1);
				var startCell = (currRow - mergeCell.span);
				var endCell = currRow - 1;
				if (endCell!=startCell){
					sheetmergeCells.push({
						startCell:columnLetter + startCell,
						endCell:columnLetter + endCell
					})
				}
				mergeCell.span = 1;
				mergeCell.perMergeValue = mergeCell.curMergeValue;
				if (!flag) {
					mergeCell.perPremiseValue = mergeCell.curPremiseValue;
				}
			}
		}
		//=========================
		for (j = 0; j < colsLength; j++) {
			var flag = true;
			for (var l = 0; l < mergeFields.length; l++) {
				var mergeCell = mergeFields[l];
				if (mergeCell.mergeField==j&&mergeCell.span>1){
					flag = false;
					break;
				}
			}
			if (!flag) continue;
			styleIndex = null;
			cellData = r[j];
			cellType = cols[j].type;
			if (typeof cols[j].beforeCellWrite === 'function') {
				var e = {
					rowNum: currRow,
					styleIndex: null,
					cellType: cellType
				};
				cellData = cols[j].beforeCellWrite(r, cellData, e);
				styleIndex = e.styleIndex || styleIndex;
				cellType = e.cellType;
				delete e;
			}
			var columnletter = getColumnLetter(j + 1);

			switch (cellType) {
			case 'number':
				row += addNumberCell(columnletter + currRow, cellData, styleIndex);
				break;
			case 'date':
				row += addDateCell(columnletter + currRow, cellData, styleIndex);
				break;
			case 'bool':
				row += addBoolCell(columnletter + currRow, cellData, styleIndex);
				break;
			default:
				row += addStringCell(self, columnletter + currRow, cellData, styleIndex);
			}
		}
		row += '</x:row>';
		rows += row;
	}
	var _sheetFront = sheetFront
	if (colsWidth !== "") {
		_sheetFront += '<x:cols>' + colsWidth + '</x:cols>';
	}
	var _sheetMergeCellString = "";
	if (sheetmergeCells.length>0){
		_sheetMergeCellString +=' <x:mergeCells count="'+sheetmergeCells.length+'">';
		for (var l = 0; l < sheetmergeCells.length; l++) {
			var sheetmergeCell = sheetmergeCells[l];
			_sheetMergeCellString +='<x:mergeCell ref="'+sheetmergeCell.startCell+':'+sheetmergeCell.endCell+'"/>';
		}
		_sheetMergeCellString +='</x:mergeCells>';
	}
	xlsx.file(config.fileName, _sheetFront + '<x:sheetData>' + rows + '</x:sheetData>' + _sheetMergeCellString + sheetBack);
}

module.exports = Sheet;

var startTag = function (obj, tagName, closed){
  var result = "<" + tagName, p;
  for (p in obj){
    result += " " + p + "=" + obj[p];
  }
  if (!closed)
    result += ">";
  else
    result += "/>";
  return result;
};

var endTag = function(tagName){
  return "</" + tagName + ">";
};

var addNumberCell = function(cellRef, value, styleIndex){
  styleIndex = styleIndex || 0;
	if (value===null)
		return "";
	else
		return '<x:c r="'+cellRef+'" s="'+ styleIndex +'" t="n"><x:v>'+value+'</x:v></x:c>';
};

var addDateCell = function(cellRef, value, styleIndex){
  styleIndex = styleIndex || 1;
	if (value===null)
		return "";
	else
		return '<x:c r="'+cellRef+'" s="'+ styleIndex +'" t="n"><x:v>'+value+'</x:v></x:c>';
};

var addBoolCell = function(cellRef, value, styleIndex){
  styleIndex = styleIndex || 0;
	if (value===null)
		return "";
	if (value){
	  value = 1;
	} else
	  value = 0;
	return '<x:c r="'+cellRef+'" s="'+ styleIndex + '" t="b"><x:v>'+value+'</x:v></x:c>';
};


var addStringCell = function(sheet, cellRef, value, styleIndex){
  styleIndex = styleIndex || 0;
	if (value===null)
		return "";
  if (typeof value ==='string'){
    value = value.replace(/&/g, "&amp;").replace(/'/g, "&apos;").replace(/>/g, "&gt;").replace(/</g, "&lt;");
  }
  var i = sheet.shareStrings.get(value, -1);
	if ( i< 0){
    i = sheet.shareStrings.length;
  	sheet.shareStrings.add(value, i);
    sheet.convertedShareStrings += "<x:si><x:t>"+value+"</x:t></x:si>";
	}
	return '<x:c r="'+cellRef+'" s="'+ styleIndex + '" t="s"><x:v>'+i+'</x:v></x:c>';
};


var getColumnLetter = function(col){
  if (col <= 0)
	throw "col must be more than 0";
  var array = new Array();
  while (col > 0)
  {
	var remainder = col % 26;
	col /= 26;
	col = Math.floor(col);
	if(remainder ===0)
	{
		remainder = 26;
		col--;
	}
	array.push(64 + remainder);
  }
  return String.fromCharCode.apply(null, array.reverse());
};
