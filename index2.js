let xlsx = require('xlsx');
let json = [ 
    {"大标题":null},
    {null:"大标题"},
    {null:"大标题"},
    {null:"大标题"},
    { Name: 'name_01', Age: 21, Address: 'address_01' },
    { Name: 'name_02', Age: 22, Address: 'address_02' },
    { Name: 'name_03', Age: 23, Address: 'address_03' },
    { Name: 'name_04', Age: 24, Address: 'address_04' },
    { Name: 'name_05', Age: 25, Address: 'address_05' }, ];

let ss = xlsx.utils.json_to_sheet(json); //通过工具将json转表对象
let keys = Object.keys(ss).sort(); //排序 [需要注意，必须从A1开始]

let ref = keys[1]+':'+keys[keys.length - 1]; //这个是定义一个字符串 也就是表的范围[A1:C5] 

let workbook = { //定义操作文档
    SheetNames:['nodejs-sheetname'], //定义表明
    Sheets:{
        'nodejs-sheetname':Object.assign({},ss,{'!ref':ref}) //表对象[注意表明]
    },
}

xlsx.writeFile(workbook,"./suzhenwei.xls"); //将数据写入文件