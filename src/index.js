const XLSX = require('xlsx');
const _ = require('lodash');

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
const loadByBrowser = (file, templates) => {
  const reader = new FileReader();
  return new Promise((resolve, reject) => {
    reader.onload = (e) => {
      try {
        const bstr = e.target.result;
        const wb = XLSX.read(bstr, {type: 'binary', cellDates: true});
        let data = processSheet(wb, templates);
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
const loadByPath = (filePath, templates) => {
  if (!fs.existsSync(filePath)) {
    throw new Error(`找不到文件[${filePath}]!`);
  }
  const workSheetArr = XLSX.readFile(filePath,{cellDates:true});
  return processSheet(workSheetArr, templates);
};

const processSheet = (workSheetArr, templates) => {
  templates = typeOf(templates) === 'Array' ? templates : [templates];
  let errInf = [];
  let jsonArr =  templates.map(st => {
    try {
      let {sheetName: name, sheetHeader: header} = st || {};
      //对于未指定sheetName的，默认第一张sheet
      name = name || Object.keys(workSheetArr.Sheets)[0];
      let sheet = workSheetArr.Sheets[name];
      if (!sheet) {
        throw new Error(`找不到工作表[${name}]`);
      }
      if (header) {
        //验证表头
        validateHeader(name, getSheetHeader(sheet), _.keys(header).filter(k => {
          return !!_.get(header[k], 'require', true);
        }))
      }
      //转成json内容
      let data = XLSX.utils.sheet_to_json(sheet, {raw: true});
      if (!header) {
        console.info('注意：该表格未指定验证header！');
        return {name, data}
      }

      data.map((r, i) => {
        //验证内容类型并mapping表头到定义key
        return validateAndMappingKey(name, r, header, i);
      });
      return {name, data};
    } catch (error) {
      errInf.push(error.message);
    }
  });
  if (!_.isEmpty(errInf)) {
    throw new Error(errInf);
  }
  return jsonArr;
};

const transformJson = (data, header) => {
  if(!header || !data){
    return data || [{'提示': '数据不存在'}];
  }
  return data.map((r, rIdx) => {
    return _.keys(header).reduce((p, c, cIdx) => {
      let h = header[c].label;
      p[h] = _.get(r, c, header[c].default);
      if(header[c].format){
        p[h] = header[c].format(p[h], rIdx, cIdx, r);
      }
      return p;
    },{});
  });
};

const downXlsxFromJson = (data, header, filename = '未命名.xlsx', sheetname = 'Sheet') => {
  let formatData = transformJson(data, header);
  let ws = XLSX.utils.json_to_sheet(formatData);
  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetname);
  /* generate file and send to client */
  XLSX.writeFile(wb, filename);
};

const downXlsxFromTable = (data, filename = '未命名.xlsx', sheetname = 'Sheet') => {
  /* convert state to workbook */
  let ws = XLSX.utils.table_to_sheet(data);
  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetname);
  /* generate file and send to client */
  XLSX.writeFile(wb, filename);
};

/**
 * 判断对象类型,返回类型有：String,Number,Date,Array,Boolean,Function,Promise,Symbol,Null,Undefined,比js自带的typeof更强
 * 需要注意NaN属于Number类型
 * @param {*} obj
 */
const typeOf = (obj) => {
  return _.trim(Object.prototype.toString.call(obj), '[]').split(' ')[1];
};

/**
 * 去除特殊字符，如：
 * "\u0000", "\u0001", "\u0002", "\u0003", "\u0004", "\u0005", "\u0006", "\u0007", "\b", "\t", "\n", "\u000b", "\f", "\r", "\u000e", "\u000f", "\u0010", "\u0011", "\u0012", "\u0013", "\u0014", "\u0015", "\u0016", "\u0017", "\u0018", "\u0019", "\u001a", "\u001b", "\u001c", "\u001d", "\u001e", "\u001f", "\"", "\\"
 * 对应ascII码：0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 34, 92
 * @param {*} str
 */
const trimCode = (str) => {
  return str.split('').filter((s, i) => {
    let code = s.charCodeAt();
    //保留双引号和反斜杠
    // return code>31 && code !==34 && code !==92;
    return code > 31;
  }).join('');
};

/**
 * 获取excel的表头
 * @param {*} sheet
 */
const getSheetHeader = (sheet) => {
  let range = XLSX.utils.decode_range(sheet['!ref']);
  if (_.isNil(range)) {
    return [];
  }
  let ha = [];
  for (let i = range.s.c; i <= range.e.c; i++) {
    let ca = XLSX.utils.encode_cell({c: i, r: range.s.r});
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
const validateHeader = (sheetName, sheetTitles, shouldTitles) => {
  if (_.xor(_.intersection(sheetTitles, shouldTitles), shouldTitles).length !== 0) {
    throw new Error(`表[${sheetName}]表头错误，标题中必须包含列名${JSON.stringify(shouldTitles)},实际列名${JSON.stringify(sheetTitles)}，注意去除空格!`);
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
const validateAndMappingKey = (name, row, header, index) => {
  let funcList = [];
  _.keys(row).map(k => {
    // let key = _.get(header[k],'key',k);
    let key = _.get(header[k], 'key');
    if (!key) {
      row[k] = undefined;
      return;
    }
    let type = _.get(header[k], 'type', 'String');
    let _type = typeOf(row[k]);
    if (_type !== type) {
      throw new Error(`表[${name}]数据类型验证失败,[列:${k}, 行:${index + 2},值:${row[k]}],数据类型不匹配，必须是[${type}]类型的数据，实际是[${_type}]类型`);
    }
    let val = _type === 'String' ? _.trim(row[k]) : row[k];
    if (_.get(header[k], 'require') && !val) {
      throw new Error(`表[${name}]数据有效性验证失败,[列:${k}, 行:${index + 2}]缺少值，该单元所在列为必填列`);
    }
    let check =  _.get(header[k], 'check');
    if (check && val) {
      funcList.push({check, k, val})
    }
    delete row[k];
    // row[k] = undefined;
    row[key] = type === 'Date' ? val.getTime() : val;
  });
  if (!_.isEmpty(funcList)) {
    funcList.map(({check, k}) => {
      let checkRes = check(row);
      if (typeOf(checkRes) !== 'Boolean') {
        throw new Error(`表[${name}]数据有效性验证失败,[列:${k}, 行:${index + 2}]值错误，${checkRes}`);
      }
    });
  }
  return row;
};

export default {
  loadByBrowser,
  loadByPath,
  downXlsxFromJson,
  downXlsxFromTable
}
