'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
var XLSX = require('xlsx');
var _ = require('lodash');

/** 
 * templates 模板样式
 * example:[
 * {
 *   sheetName:'1.sheet1',
 *   sheetHeader:{
 *     "姓名":{key:"name", type:"String"},
 *     "年龄":{key:"age", type:"Number", check: row=>{return row.age>18}},
 *     "状态":{key:"status", type:"String", require="false"},
 *   }
 * },
 * {
 *   sheetName:'2.sheet2',
 *   sheetHeader:{
 *     "公司":{key:"company", type:"String"},
 *     "合作公司":{key:"relation_company", type:"String"},
 *   }
 * }
 * type取值范围：String,Number,Date,Array,Function,Promise,Symbol,Null,Undefined
 */

/**
 * 从浏览器读取excel
 * @param {*} file 
 * @param {*} templates 
 */
var loadByBrowser = function loadByBrowser(file, templates) {
  var reader = new FileReader();
  return new Promise(function (resolve, reject) {
    reader.onload = function (e) {
      try {
        var bstr = e.target.result;
        var wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        var data = processSheet(wb, templates);
        resolve(data);
      } catch (e) {
        reject(e);
      }
    };
    reader.readAsBinaryString(file);
  });
};

/**
 * 从文件系统直接读取excel
 * @param {*} filePath 
 * @param {*} templates 
 */
var loadByPath = function loadByPath(filePath, templates) {
  if (!fs.existsSync(filePath)) {
    throw new Error('\u627E\u4E0D\u5230\u6587\u4EF6[' + filePath + ']!');
  }
  var workSheetArr = XLSX.readFile(filePath, { cellDates: true });
  return processSheet(workSheetArr, templates);
};

var processSheet = function processSheet(workSheetArr, templates) {
  templates = typeOf(templates) === 'Array' ? templates : [templates];
  var errInf = [];
  var jsonArr = templates.map(function (st) {
    try {
      var _ref = st || {},
          name = _ref.sheetName,
          header = _ref.sheetHeader;
      //对于未指定sheetName的，默认第一张sheet


      name = name || Object.keys(workSheetArr.Sheets)[0];
      var sheet = workSheetArr.Sheets[name];
      if (!sheet) {
        throw new Error('\u627E\u4E0D\u5230\u5DE5\u4F5C\u8868[' + name + ']');
      }
      if (header) {
        //验证表头
        validateHeader(name, getSheetHeader(sheet), _.keys(header).filter(function (k) {
          return !!_.get(header[k], 'require', true);
        }));
      }
      //转成json内容
      var data = XLSX.utils.sheet_to_json(sheet, { raw: true });
      if (!header) {
        console.info('注意：该表格未指定验证header！');
        return { name: name, data: data };
      }

      data.map(function (r, i) {
        //验证内容类型并mapping表头到定义key
        return validateAndMappingKey(name, r, header, i);
      });
      return { name: name, data: data };
    } catch (error) {
      errInf.push(error.message);
    }
  });
  if (!_.isEmpty(errInf)) {
    throw new Error(errInf);
  }
  return jsonArr;
};

var transformJson = function transformJson(data, header) {
  if (!header || !data) {
    return data || [{ '提示': '数据不存在' }];
  }
  return data.map(function (r, rIdx) {
    return _.keys(header).reduce(function (p, c, cIdx) {
      var h = header[c].label;
      p[h] = _.get(r, c, header[c].default);
      if (header[c].format) {
        p[h] = header[c].format(p[h], rIdx, cIdx);
      }
      return p;
    }, {});
  });
};

var downXlsxFromJson = function downXlsxFromJson(data, header) {
  var filename = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : '未命名.xlsx';

  var formatData = transformJson(data, header);
  var ws = XLSX.utils.json_to_sheet(formatData);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'SheetJS');
  /* generate file and send to client */
  XLSX.writeFile(wb, filename);
};

var downXlsxFromTable = function downXlsxFromTable(data) {
  var filename = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : '未命名.xlsx';

  /* convert state to workbook */
  var ws = XLSX.utils.table_to_sheet(data);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'SheetJS');
  /* generate file and send to client */
  XLSX.writeFile(wb, filename);
};

/**
 * 判断对象类型,返回类型有：String,Number,Date,Array,Boolean,Function,Promise,Symbol,Null,Undefined,比js自带的typeof更强
 * 需要注意NaN属于Number类型
 * @param {*} obj
 */
var typeOf = function typeOf(obj) {
  return _.trim(Object.prototype.toString.call(obj), '[]').split(' ')[1];
};

/**
 * 去除特殊字符，如：
 * "\u0000", "\u0001", "\u0002", "\u0003", "\u0004", "\u0005", "\u0006", "\u0007", "\b", "\t", "\n", "\u000b", "\f", "\r", "\u000e", "\u000f", "\u0010", "\u0011", "\u0012", "\u0013", "\u0014", "\u0015", "\u0016", "\u0017", "\u0018", "\u0019", "\u001a", "\u001b", "\u001c", "\u001d", "\u001e", "\u001f", "\"", "\\"
 * 对应ascII码：0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 34, 92
 * @param {*} str
 */
var trimCode = function trimCode(str) {
  return str.split('').filter(function (s, i) {
    var code = s.charCodeAt();
    //保留双引号和反斜杠
    // return code>31 && code !==34 && code !==92;
    return code > 31;
  }).join('');
};

/**
 * 获取excel的表头
 * @param {*} sheet
 */
var getSheetHeader = function getSheetHeader(sheet) {
  var range = XLSX.utils.decode_range(sheet['!ref']);
  if (_.isNil(range)) {
    return [];
  }
  var ha = [];
  for (var i = range.s.c; i <= range.e.c; i++) {
    var ca = XLSX.utils.encode_cell({ c: i, r: range.s.r });
    ha.push(_.get(sheet[ca], 'v', '').split(/\r|\n/)[0]);
  }
  return ha;
};

/**
 * 判断表头是否与预设的一致
 * @param {表名} sheetName
 * @param {实际表头内容} sheetTitles
 * @param {模板预设表头内容} shouldTitles
 */
var validateHeader = function validateHeader(sheetName, sheetTitles, shouldTitles) {
  if (_.xor(_.intersection(sheetTitles, shouldTitles), shouldTitles).length !== 0) {
    throw new Error('\u8868[' + sheetName + ']\u8868\u5934\u9519\u8BEF\uFF0C\u6807\u9898\u4E2D\u5FC5\u987B\u5305\u542B\u5217\u540D' + JSON.stringify(shouldTitles) + ',\u5B9E\u9645\u5217\u540D' + JSON.stringify(sheetTitles) + '\uFF0C\u6CE8\u610F\u53BB\u9664\u7A7A\u683C!');
  }
  return true;
};

/**
 * 原始json数组转换成服务端能识别的json数组，
 * mapper格式：
 * @code
 * {
 *    "姓名": {key: "name", type: "String"},
 *    "性别": {key: "gender", type: "String"},
 *    "出生日期": {key: "birthday", type: "Date"},
 *    "证件类型": {key: "certType", type: "String"},
 *    "证件号": {key: "certNo", type: "String"},
 *    "手机号码": {key: "phone", type: "String"},
 *    "身高（cm）": {key: "height", type: "String", require:false},
 *    "QQ号": {key: "qq", type: "String", require:false},
 *    "微信号": {key: "wechat", type: "String", require:false},
 * }
 * @param {*} name
 * @param {*} row
 * @param {*} header
 * @param {*} index
 */
var validateAndMappingKey = function validateAndMappingKey(name, row, header, index) {
  var funcList = [];
  _.keys(row).map(function (k) {
    // let key = _.get(header[k],'key',k);
    var key = _.get(header[k], 'key');
    if (!key) {
      row[k] = undefined;
      return;
    }
    var type = _.get(header[k], 'type', 'String');
    var _type = typeOf(row[k]);
    if (_type !== type) {
      throw new Error('\u8868[' + name + ']\u6570\u636E\u7C7B\u578B\u9A8C\u8BC1\u5931\u8D25,[\u5217:' + k + ', \u884C:' + (index + 2) + ',\u503C:' + row[k] + '],\u6570\u636E\u7C7B\u578B\u4E0D\u5339\u914D\uFF0C\u5FC5\u987B\u662F[' + type + ']\u7C7B\u578B\u7684\u6570\u636E\uFF0C\u5B9E\u9645\u662F[' + _type + ']\u7C7B\u578B');
    }
    var val = _type === 'String' ? _.trim(row[k]) : row[k];
    if (_.get(header[k], 'require') && !val) {
      throw new Error('\u8868[' + name + ']\u6570\u636E\u6709\u6548\u6027\u9A8C\u8BC1\u5931\u8D25,[\u5217:' + k + ', \u884C:' + (index + 2) + ']\u7F3A\u5C11\u503C\uFF0C\u8BE5\u5355\u5143\u6240\u5728\u5217\u4E3A\u5FC5\u586B\u5217');
    }
    var check = _.get(header[k], 'check');
    if (check && val) {
      funcList.push({ check: check, k: k, val: val });
    }
    delete row[k];
    // row[k] = undefined;
    row[key] = type === 'Date' ? val.getTime() : val;
  });
  if (!_.isEmpty(funcList)) {
    funcList.map(function (_ref2) {
      var check = _ref2.check,
          k = _ref2.k;

      var checkRes = check(row);
      if (typeOf(checkRes) !== 'Boolean') {
        throw new Error('\u8868[' + name + ']\u6570\u636E\u6709\u6548\u6027\u9A8C\u8BC1\u5931\u8D25,[\u5217:' + k + ', \u884C:' + (index + 2) + ']\u503C\u9519\u8BEF\uFF0C' + checkRes);
      }
    });
  }
  return row;
};

exports.default = {
  loadByBrowser: loadByBrowser,
  loadByPath: loadByPath,
  downXlsxFromJson: downXlsxFromJson,
  downXlsxFromTable: downXlsxFromTable
};