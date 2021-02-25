var vueApp, allkeys, dataSource = [], ldld;
Vue.component('v-select', VueSelect.VueSelect);
$(function () {
	ldld = new ldLoader({root: "#bean-loader"});
	initVueApp();
});

function initVueApp () {
  vueApp = new Vue({
    el: '#figureApp',
    data: {
    	sheetList: [],
    	options: [],
        placeholder: '请选择需要处理的表格',
        sheet_name_list: [],
        file_number: 0,
        export_excel_name: "导出.xlsx",
        export_sheet_name: "sheet1",
        isLoading: false
    },
    computed: {
    },
    methods: {
    	clear: function() {
    		this.sheetList = [];
    		this.options = [];
    		this.sheet_name_list = [];
    		this.file_number = 0;
    	},
    	importExcelFile: function () {
    		this.clear();
			this.isLoading = true;
            importExcel();
    	},
    	handleSheetData: function (sheet, sheetName, fileName) {
    		let companyCode = fileName.split("-")[0];
    		var sheetdataList = []
    		for (var key in sheet) {
    			let type = sheet[key].t;
    			let value = key[0] === '!' ? sheet[key] : sheet[key].v
    			if (type === "n") {
    				var row = key.replace(/[^a-z]+/ig,"");
    				var column = key.replace(/[^0-9]/ig,""); 
    				sheetdataList.push({"公司": companyCode, "行号": row, "列号": column, "行列": key, "值": value, "表名": sheetName});
    			}
    		}
    		return sheetdataList;
    	},
    	openDownloadDialog: function(url, saveName) {
			if(typeof url == 'object' && url instanceof Blob)
			{
				url = URL.createObjectURL(url); // 创建blob地址
			}
			var aLink = document.createElement('a');
			aLink.href = url;
			aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
			var event;
			if(window.MouseEvent) event = new MouseEvent('click');
			else
			{
				event = document.createEvent('MouseEvents');
				event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
			}
			aLink.dispatchEvent(event);
		},
		sheet2blob: function(sheet, sheetName) {
			sheetName = sheetName || 'sheet1';
			var workbook = {
				SheetNames: [sheetName],
				Sheets: {}
			};
			workbook.Sheets[sheetName] = sheet;
			// 生成excel的配置项
			var wopts = {
				bookType: 'xlsx', // 要生成的文件类型
				bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
				type: 'binary'
			};
			var wbout = XLSX.write(workbook, wopts);
			var blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});
			// 字符串转ArrayBuffer
			function s2ab(s) {
				var buf = new ArrayBuffer(s.length);
				var view = new Uint8Array(buf);
				for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
				return buf;
			}
			return blob;
		},
		selecte_sheet_name: function(values) {
			if (values.length > 0) {
    			$('.export-btn').removeAttr("disabled");
    		} else {
    			$('.export-btn').attr('disabled', true);
    		}
        },
        export_sheet: function() {
        	var sheetdataList = [];
        	for (let sheet of this.sheetList) {
        		let sheetName = sheet.sheetName;
        		let fileName = sheet.fileName;
        		if (this.sheet_name_list.includes(sheetName)) {
	        		let result = this.handleSheetData(sheet.sheet, sheetName, fileName);
	        		sheetdataList.push(...result);
	        	}
        	}
        	
        	let sortedList = sheetdataList.sort(function (cellA, cellB) {
    			return cellA.companyCode < cellB.companyCode;
    		});
    		var ws = XLSX.utils.json_to_sheet(sortedList);
			var blob = this.sheet2blob(ws, this.export_sheet_name);
			this.openDownloadDialog(blob, this.export_excel_name);
        }
    },
    watch: {
    	
    },
    created: function () {

    },
    mounted: function () {
    	
    },
    updated: function () {
      
    }
  });
}

