// delete a specific row
function clamp_range(range) {
    if(range.e.r >= (1<<20)) range.e.r = (1<<20)-1;
    if(range.e.c >= (1<<14)) range.e.c = (1<<14)-1;
    return range;
}

var crefregex = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)([1-9]\d{0,5}|10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6])(?![_.\(A-Za-z0-9])/g;

/*
    deletes `nrows` rows STARTING WITH `start_row`
    - ws         = worksheet object
    - start_row  = starting row (0-indexed) | default 0
    - nrows      = number of rows to delete | default 1
*/

function delete_rows(ws, start_row, nrows) {
    if(!ws) throw new Error("operation expects a worksheet");
    var dense = Array.isArray(ws);
    if(!nrows) nrows = 1;
    if(!start_row) start_row = 0;

    /* extract original range */
    var range = XLSX.utils.decode_range(ws["!ref"]);
    var R = 0, C = 0;

    var formula_cb = function($0, $1, $2, $3, $4, $5) {
        var _R = XLSX.utils.decode_row($5), _C = XLSX.utils.decode_col($3);
        if(_R >= start_row) {
            _R -= nrows;
            if(_R < start_row) return "#REF!";
        }
        return $1+($2=="$" ? $2+$3 : XLSX.utils.encode_col(_C))+($4=="$" ? $4+$5 : XLSX.utils.encode_row(_R));
    };

    var addr, naddr;
    /* move cells and update formulae */
    if(dense) {
        for(R = start_row + nrows; R <= range.e.r; ++R) {
            if(ws[R]) ws[R].forEach(function(cell) { cell.f = cell.f.replace(crefregex, formula_cb); });
            ws[R-nrows] = ws[R];
        }
        ws.length -= nrows;
        for(R = 0; R < start_row; ++R) {
            if(ws[R]) ws[R].forEach(function(cell) { cell.f = cell.f.replace(crefregex, formula_cb); });
        }
    } else {
        for(R = start_row + nrows; R <= range.e.r; ++R) {
            for(C = range.s.c; C <= range.e.c; ++C) {
                addr = XLSX.utils.encode_cell({r:R, c:C});
                naddr = XLSX.utils.encode_cell({r:R-nrows, c:C});
                if(!ws[addr]) { delete ws[naddr]; continue; }
                if(ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
                ws[naddr] = ws[addr];
            }
        }
        for(R = range.e.r; R > range.e.r - nrows; --R) {
            for(C = range.s.c; C <= range.e.c; ++C) {
                addr = XLSX.utils.encode_cell({r:R, c:C});
                delete ws[addr];
            }
        }
        for(R = 0; R < start_row; ++R) {
            for(C = range.s.c; C <= range.e.c; ++C) {
                addr = XLSX.utils.encode_cell({r:R, c:C});
                if(ws[addr] && ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
            }
        }
    }

    /* write new range */
    range.e.r -= nrows;
    if(range.e.r < range.s.r) range.e.r = range.s.r;
    ws["!ref"] = XLSX.utils.encode_range(clamp_range(range));

    /* merge cells */
    if(ws["!merges"]) ws["!merges"].forEach(function(merge, idx) {
        var mergerange;
        switch(typeof merge) {
            case 'string': mergerange = XLSX.utils.decode_range(merge); break;
            case 'object': mergerange = merge; break;
            default: throw new Error("Unexpected merge ref " + merge);
        }
        if(mergerange.s.r >= start_row) {
            mergerange.s.r = Math.max(mergerange.s.r - nrows, start_row);
            if(mergerange.e.r < start_row + nrows) { delete ws["!merges"][idx]; return; }
        } else if(mergerange.e.r >= start_row) mergerange.e.r = Math.max(mergerange.e.r - nrows, start_row);
        clamp_range(mergerange);
        ws["!merges"][idx] = mergerange;
    });
    if(ws["!merges"]) ws["!merges"] = ws["!merges"].filter(function(x) { return !!x; });

    /* rows */
    if(ws["!rows"]) ws["!rows"].splice(start_row, nrows);
}

var type_list = {
    '':{},
    '':{}
};

function handleFile() {
    const fileInput = document.getElementById('fileInput');
    
    if (fileInput.files.length === 0) {
        console.log('沒有選擇任何檔案');
        return;
    }
   
    var counts={};
    var counts_html = "";
    var excel_table = $("#excel_table");
    var reg_alpha_num = new RegExp("^[ a-zA-Z0-9]+$");
    var col_key=' 檢查部位';
    var jsonData=[];
    var processed_ws;
    var htmlTable;
    var total_files=0;
	var total_counts=0;
    
    // 清空excel table
    excel_table.empty();

    // 建立一個陣列來存放所有檔案的Promise
    const filePromises = [];

    // 逐一處理每個檔案
    for (const file of fileInput.files) {
        const fileReader = new FileReader();
    
        total_files++;
        console.log("read file ["+total_files+"].");

        // 定義當檔案讀取完成時的處理函式
        const filePromise = new Promise((resolve, reject) => {
            fileReader.onload = (event) => {
                const contents = event.target.result;
                
                const workbook = XLSX.read(contents, {type: 'array'});
                //const sheetName = workbook.SheetNames[0];
                //const worksheet = workbook.Sheets[sheetName];
				
				// sheet迴圈
				for(const sheet_name of workbook.SheetNames) {
					if(!(sheet_name.includes("門診") || sheet_name.includes("住診")))
						continue;

					const worksheet = workbook.Sheets[sheet_name];
					
					// 首先移除不需要的前幾項
					delete_rows(worksheet, 0, 3);

					// 先轉為json
					jsonData=jsonData.concat(XLSX.utils.sheet_to_json(worksheet));
					resolve(contents);
				}

                
            };

            // 處理讀取檔案時的錯誤
            fileReader.onerror = (event) => {
                reject(event.target.error);
            };
        });

        // 開始讀取檔案
        fileReader.readAsArrayBuffer(file);
        filePromises.push(filePromise);
      }
    
    // 等待所有檔案都讀取完成
    return Promise.all(filePromises)
        .then((fileContents) => {
            // 處理資料: sorting
            jsonData.sort((a, b) => {           
                    return a[col_key].localeCompare(b[col_key]);
            });
                
                
            // people counting 
            jsonData.forEach(function (o) {
                // TODO: group and type list
                //if (reg_alpha_num.test(o[col_key])){
                if(true){
                    if (!counts.hasOwnProperty(o[col_key])) {
                        counts[o[col_key]] = 0;
                    }
                    counts[o[col_key]] += 1;
                }
            });
            
            // 再轉回sheet
            processed_ws = XLSX.utils.json_to_sheet(jsonData);
            
            // show result of people counting
            counts_html='<div style="border-style: solid;">';
            
            // 針對結果排序
            var ordered_counts = Object.keys(counts).sort().reduce(
              (obj, key) => { 
                obj[key] = counts[key]; 
                return obj;
              }, 
              {}
            );
            
            for (const [key, value] of Object.entries(ordered_counts)) {
              counts_html+='<div>'+key+'：'+value+'</div>';
			  total_counts+=value;
            }
            
            counts_html+='</div>';
			
			// 加上總數
			counts_html = '<div>Total: '+total_counts+'</div>'+counts_html;
			
            excel_table.append($(counts_html));

            // show table
            htmlTable = XLSX.utils.sheet_to_html(processed_ws);
            excel_table.append($(htmlTable));
        })
        .catch((error) => {
            console.error('讀取檔案時發生錯誤:', error);
        });
}