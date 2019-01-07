// Power BI Matrix Visual to TSV
var headerWait = 200;
var bodyWait = 200;

function getColumnHeaders(pivotTable, scrollFlag, steps, lefts, cells) {
	return new Promise(function(resolve) {
		var maxLeft = -1;
		if (lefts.length > 0) {
			maxLeft = Math.max(...lefts);
		}
		steps.push(maxLeft);
		if (scrollFlag) {pivotTable.find("div.bodyCells").scrollLeft(maxLeft);}
		setTimeout(function(){
			var tmp = pivotTable.find("div.columnHeaders").find("div.pivotTableCellWrap,div.pivotTableCellNoWrap");
			continueFlag = false;
			for (var i = 0; i < tmp.length; i++) {
				var t = parseInt($(tmp[i]).parent().css("top"));
				var l = parseInt($(tmp[i]).parent().css("left"));
				if (l > maxLeft) {
					continueFlag = true;
					lefts.push(l);
					cells.push({top:t, left:l, step:maxLeft, offset:$(tmp[i]).offset().left, value:$(tmp[i]).text().trim()});
				}
			}
			resolve(scrollFlag&&continueFlag);
		}, headerWait);
	})
}

function getRowHeaders(pivotTable, scrollFlag, steps, tops, cells) {
	return new Promise(function(resolve) {
		var maxTop = -1;
		if (tops.length > 0) {
			maxTop = Math.max(...tops);
		}
		steps.push(maxTop);
		if (scrollFlag) {pivotTable.find("div.bodyCells").scrollTop(maxTop);}
		setTimeout(function(){
			var tmp = pivotTable.find("div.rowHeaders").find("div.pivotTableCellWrap,div.pivotTableCellNoWrap");
			continueFlag = false;
			for (var i = 0; i < tmp.length; i++) {
				var pdiv = $(tmp[i]).parent();
				if (pdiv.hasClass("expandableCell")) {pdiv = pdiv.parent();}
				var t = parseInt(pdiv.css("top"));
				var l = parseInt(pdiv.css("left"));
				if (t > maxTop) {
					continueFlag = true;
					tops.push(t);
					var indent = Math.floor(parseInt($(tmp[i]).css("padding-left"))/10);
					var val = $(tmp[i]).text().trim();
					for (var j = 0; j < indent; j++) {val = " "+val;}
					cells.push({top:t, left:l, step:maxTop, offset:pdiv.offset().top, value:val});
				}
			}
			resolve(scrollFlag&&continueFlag);
		}, headerWait);
	})
}

function getBodyCells(pivotTable, rIndex, cIndex, cells) {
	return new Promise(function(resolve) {
		setTimeout(function(){
			var tmp = pivotTable.find("div.bodyCells").find("div.pivotTableCellWrap,div.pivotTableCellNoWrap");
			for (var k = 0; k < tmp.length; k++) {
				var val = $(tmp[k]).text().trim();
				if (val == "") {continue;}
				var offset = $(tmp[k]).offset();
				var i = rIndex[""+offset.top];
				var j = cIndex[""+offset.left];
				if (i >= 0 && j >= 0) {
					cells[i][j] = '"'+val.replace(/"/g, '""')+'"';
				}
			}
			resolve(true);
		}, bodyWait);
	})
}

function getFloatingBodyCells(pivotTable, rIndex, cIndex, cells) {
	return new Promise(function(resolve) {
		setTimeout(function(){
			var tmp = pivotTable.find("div.floatingBodyCells").find("div.pivotTableCellWrap,div.pivotTableCellNoWrap");
			for (var k = 0; k < tmp.length; k++) {
				var val = $(tmp[k]).text().trim();
				if (val == "") {continue;}
				var offset = $(tmp[k]).offset();
				var i = rIndex[""+offset.top];
				var j = cIndex[""+offset.left];
				if (i >= 0 && j >= 0) {
					cells[i][j] = '"'+val.replace(/"/g, '""')+'"';
				}
			}
			resolve(true);
		}, bodyWait);
	})
}

async function getCells(pivotTable) {
	var cornerLefts = [];
	var cornerTops = [];
	var headerLefts = [];
	var headerTops = [];
	var floatingTops = [];
	var cornerCells = [];
	var colHeaderCells = [];
	var rowHeaderCells = [];
	var floatingRowHeaderCells = [];

	var hScrollBar = $(pivotTable.find("div.scroll-bar-div")[0])
	var vScrollBar = $(pivotTable.find("div.scroll-bar-div")[1])
	var hScroll = (hScrollBar.css("visibility") == "visible") ? true : false;
	var vScroll = (vScrollBar.css("visibility") == "visible") ? true : false;
	var hSteps = [];
	var vSteps = [];
	var colHeaderIndex = {};
	var rowHeaderIndex = {};
	var floatingIndex = {};
	var colIndex = {};
	var rowIndex = {};
	var cells = [];

	// Corner
	var corner = pivotTable.find("div.corner");
	var tmp = corner.find("div.pivotTableCellWrap,div.pivotTableCellNoWrap");
	for (var i = 0; i < tmp.length; i++) {
		var l = parseInt($(tmp[i]).parent().css("left"));
		var t = parseInt($(tmp[i]).parent().css("top"));
		cornerLefts.push(l);
		cornerTops.push(t);
		cornerCells.push({top:t, left:l, value:$(tmp[i]).text().trim()});
	}
	cornerLefts = cornerLefts.filter(function (x, i, self) {
		return self.indexOf(x) === i;
	}).sort((a, b) => a - b);
	for (var i = 0; i < cornerLefts.length; i++) {
		colHeaderIndex[""+cornerLefts[i]] = i;
	}
	cornerTops = cornerTops.filter(function (x, i, self) {
		return self.indexOf(x) === i;
	}).sort((a, b) => a - b);
	for (var i = 0; i < cornerTops.length; i++) {
		rowHeaderIndex[""+cornerTops[i]] = i;
	}

	// Column Header
	var flag = true;
	while (flag) {
		flag = await getColumnHeaders(pivotTable, hScroll, hSteps, headerLefts, colHeaderCells);
	}
	headerLefts = headerLefts.filter(function (x, i, self) {
		return self.indexOf(x) === i;
	}).sort((a, b) => a - b);

	// Row Header
	flag = true;
	while (flag) {
		flag = await getRowHeaders(pivotTable, vScroll, vSteps, headerTops, rowHeaderCells);
	}
	headerTops = headerTops.filter(function (x, i, self) {
		return self.indexOf(x) === i;
	}).sort((a, b) => a - b);

	// Floating Row Header
	tmp = pivotTable.find("div.floatingRowHeader").find("div.pivotTableCellWrap,div.pivotTableCellNoWrap");
	for (var i = 0; i < tmp.length; i++) {
		var l = parseInt($(tmp[i]).parent().css("left"));
		var t = parseInt($(tmp[i]).parent().css("top"));
		floatingTops.push(t);
		floatingRowHeaderCells.push({top:t, left:l, offset:$(tmp[i]).offset().top, value:$(tmp[i]).text().trim()});
	}
	floatingTops = floatingTops.filter(function (x, i, self) {
		return self.indexOf(x) === i;
	}).sort((a, b) => a - b);

	// Container for output
	var m = cornerLefts.length+headerLefts.length;
	var n = cornerTops.length+headerTops.length+floatingTops.length;
	for (var i = 0; i < n; i++) {
		cells.push([]);
		for (var j = 0; j < m; j++) {
			cells[i].push(null);
		}
	}

	// Make Header Index & Store Header Value
	for (var k = 0; k < cornerCells.length; k++) {
		var c = cornerCells[k];
		var i = rowHeaderIndex[""+c.top];
		var j = colHeaderIndex[""+c.left];
		cells[i][j] = '"'+c.value.replace(/"/g, '""')+'"';
	}

	for (var i = 0; i < hSteps.length; i++) {
		colIndex[""+hSteps[i]] = {}
	}
	for (var i = 0; i < vSteps.length; i++) {
		rowIndex[""+vSteps[i]] = {}
	}
	for (var k = 0; k < colHeaderCells.length; k++) {
		var c = colHeaderCells[k];
		var i = rowHeaderIndex[""+c.top];
		var j = headerLefts.indexOf(c.left)+cornerLefts.length;
		cells[i][j] = '"'+c.value.replace(/"/g, '""')+'"';
		colIndex[""+c.step][""+c.offset] = j;
	}
	for (var k = 0; k < rowHeaderCells.length; k++) {
		var c = rowHeaderCells[k];
		var i = headerTops.indexOf(c.top)+cornerTops.length;
		var j = colHeaderIndex[""+c.left];
		cells[i][j] = '"'+c.value.replace(/"/g, '""')+'"';
		rowIndex[""+c.step][""+c.offset] = i;
	}
	for (var k = 0; k < floatingRowHeaderCells.length; k++) {
		var c = floatingRowHeaderCells[k];
		var i = floatingTops.indexOf(c.top)+cornerTops.length+headerTops.length;
		var j = colHeaderIndex[""+c.left];
		cells[i][j] = '"'+c.value.replace(/"/g, '""')+'"';
		floatingIndex[""+c.offset] = i;
	}

	// Body Cells
	m = hScroll ? hSteps.length-1 : 1;
	n = vScroll ? vSteps.length-1 : 1;
	for (var i = 0; i < n; i++) {
		if (vScroll) {pivotTable.find("div.bodyCells").scrollTop(vSteps[i]);}
		for (var j = 0; j < m; j++) {
			if (hScroll) {pivotTable.find("div.bodyCells").scrollLeft(hSteps[j]);}
			await getBodyCells(pivotTable, rowIndex[""+vSteps[i]], colIndex[""+hSteps[j]], cells);
		}
	}

	// Floating Body Cells
	for (var j = 0; j < m; j++) {
		if (hScroll) {pivotTable.find("div.bodyCells").scrollLeft(hSteps[j]);}
		await getFloatingBodyCells(pivotTable, floatingIndex, colIndex[""+hSteps[j]], cells);
	}

	return cells;
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
	var response = {result: false, message: "??? Unknown Error"};
	
	var url = location.href;
	var pivotTable = $(document).find("div.pivotTable");
	if (!url.match(/^https:\/\/app.powerbi.com\//)) {
		response.message = "Not PowerBI site.";
		sendResponse(response);
	} else if (!pivotTable || (pivotTable.length == 0)) {
		response.message = "Missing Matrix Visual.";
		sendResponse(response);
	} else {
		getCells(pivotTable).then(cells => {
			var tsv = "";
			for (var i = 0; i < cells.length; i++) {
				tsv += cells[i].join("\t")+"\n";
			}
			response.result = true;
			response.message = tsv;
			sendResponse(response);
		});
	}

	// returns true if asyncronous is needed
	return true;
});
