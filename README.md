# xlsx-json

- 基于 `SheetJS/js-xlsx` 改造，把excel的二维表格，转化成`key->value`形式的json数据,不是xlsx例子中的二维数组json。
- 提供了一套模板，用来转化key为后端能识别的key，同时，对表中的各列字段做了验证
- 前端项目：浏览器需要支持es6，async和await。后端项目：nodejs 8+

## 模板格式

支持多sheet验证
example: 一下一张表有2个sheet

```js
[
  {
    sheetName:'1.用户表',
    sheetHeader:{
      "名称":{key:"name", type:"String"},
      "状态":{key:"status", type:"String"},
      "年龄":{key:"age", type:"Number", check: row=>{return age > 18 ? true : '年龄必须大于18岁';}},
      "省":{key:"province", type:"String"},
      "市":{key:"city", type:"String"},
      "区/县":{key:"region", type:"String"},
      "地址":{key:"address", type:"String", require:false},
    }
  },
  {
    sheetName:'2.角色表',
    sheetHeader:{
      "角色名称":{key:"company", type:"String"},
    }
  },
```

如果只有一个sheet，可以省去sheetName，为：

```js
[
  {
    sheetHeader:{
      "名称":{key:"name", type:"String"},
      "状态":{key:"status", type:"String"},
      "年龄":{key:"age", type:"Number", check: row=>{return age > 18 ? true : '年龄必须大于18岁';}},
      "省":{key:"province", type:"String"},
      "市":{key:"city", type:"String"},
      "区/县":{key:"region", type:"String"},
      "地址":{key:"address", type:"String", require:false},
    }
  }
]
```