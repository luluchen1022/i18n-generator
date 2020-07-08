var XLSX = require('xlsx');
var electron = require('electron').remote;
var fs = require('fs')
var resx = require('resx')

var process_wb = (function() {
	var HTMLOUT = document.getElementById('htmlout');
	var XPORT = document.getElementById('xport');

	return function process_wb(wb) {
		XPORT.disabled = false;
		HTMLOUT.innerHTML = "";
		wb.SheetNames.forEach(function(sheetName) {
			var htmlstr = XLSX.utils.sheet_to_html(wb.Sheets[sheetName],{editable:true});
			HTMLOUT.innerHTML += htmlstr;
		});
	};
})();

var do_file = (function() {
	return function do_file(files) {
		var f = files[0];
		var reader = new FileReader();
		reader.onload = function(e) {
			var data = e.target.result;
			data = new Uint8Array(data);
			process_wb(XLSX.read(data, {type: 'array'}));
		};
		reader.readAsArrayBuffer(f);
	};
})();

(function() {
	var drop = document.getElementById('drop');

	function handleDrop(e) {
		e.stopPropagation();
		e.preventDefault();
		do_file(e.dataTransfer.files);
	}

	function handleDragover(e) {
		e.stopPropagation();
		e.preventDefault();
		e.dataTransfer.dropEffect = 'copy';
	}

	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
})();

(function() {
	var readf = document.getElementById('readf');
	async function handleF(/*e*/) {
		var o = await electron.dialog.showOpenDialog({
			title: 'Select a file',
			filters: [{
				name: "Spreadsheets",
				extensions: "xls|xlsx|xlsm|xlsb|xml|xlw|xlc|csv|txt|dif|sylk|slk|prn|ods|fods|uos|dbf|wks|123|wq1|qpw|htm|html".split("|")
			}],
			properties: ['openFile']
		});
		if(o.filePaths.length > 0) process_wb(XLSX.readFile(o.filePaths[0]));
	}
	readf.addEventListener('click', handleF, false);
})();

(function() {
	var xlf = document.getElementById('xlf');
	function handleFile(e) { do_file(e.target.files); }
	xlf.addEventListener('change', handleFile, false);
})();

var export_xlsx = (function() {
	var HTMLOUT = document.getElementById('htmlout');
	var XTENSION = "json|resx".split("|")
	return async function() {
		var wb = XLSX.utils.table_to_book(HTMLOUT);
		var content = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) 
		var o = await electron.dialog.showSaveDialog({
			title: 'Save file as',
			filters: [{
				name: "Spreadsheets",
				extensions: XTENSION
			}]
		});
		console.log(o.filePath);
		const exportResult = (result) => {
			fs.writeFile(o.filePath, result, (err) => {
				if (err) {
					console.error(err)
					return
				}
				electron.dialog.showMessageBox({ message: "Exported data to " + o.filePath, buttons: ["OK"] });
			})
		}
		if (o.filePath.includes('json')) {
			exportResult(JSON.stringify(content))
		} else {
			resx.js2resx(content, (err, resx) => exportResult(resx))
		}
	};
})();
