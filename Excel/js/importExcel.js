function importExcel() {
	let _this = vueApp;
    let inputDOM = _this.$refs.inputer;
    // 通过DOM取文件数据
    _this.files = event.currentTarget.files;
    var rABS = false; //是否将文件读取为二进制字符串
    var fs = _this.files;
    _this.file_number = fs.length;
    if (fs.length == 0) {
    	_this.isLoading = false;
    	return
    }
    var reader = new FileReader();
    //if (!FileReader.prototype.readAsBinaryString) {
	var fileIndex = 0
	var sheetNameList = [];
    let promise = new Promise(function(resolve, reject) {
        FileReader.prototype.readAsBinaryString = function(f) {
            var binary = "";
            var rABS = false; //是否将文件读取为二进制字符串
            var pt = this;
            var wb; //读取完成的数据
            var reader = new FileReader();
            reader.onload = function(e) {
                var bytes = new Uint8Array(reader.result);
                var length = bytes.byteLength;
                for(var i = 0; i < length; i++) {
                    binary += String.fromCharCode(bytes[i]);
                }
                if(rABS) {
                    wb = XLSX.read(btoa(fixdata(binary)), { //手动转化
                        type: 'base64'
                    });
                } else {
                    wb = XLSX.read(binary, {
                        type: 'binary'
                    });
                }
                for (let sheetName of wb.SheetNames) {
                	let sheet = wb.Sheets[sheetName];
                	sheetNameList.push(sheetName);
                	_this.sheetList.push({'fileName': f.name, 'sheetName': sheetName, 'sheet': sheet});
                }
                fileIndex++;
                if (fileIndex == fs.length) {
                	resolve(fileIndex);
                }
            }
            reader.readAsArrayBuffer(f);
        }
	});	

    if(rABS) {
    	for (f of fs) {
    		reader.readAsArrayBuffer(f);
    	}
    } else {
    	for (f of fs) {
    		reader.readAsBinaryString(f);
    	}
    	promise.then(function() {
    		_this.options = Array.from(new Set(sheetNameList));
    		_this.isLoading = false;
    	});
    }
}