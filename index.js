let xlsx = require("xlsx");

let _ = require("lodash");

let workbook = xlsx.readFile("3.xlsx"); //workbook就是xls文档对象
formatData(workbook)
function formatData(workbook) {
  let sheetNames = workbook.SheetNames; //获取表明
  let sheet = workbook.Sheets[sheetNames[1]]; //通过表明得到表对象
  var data = xlsx.utils.sheet_to_json(sheet); //通过工具将表对象的数据读出来并转成json
  var data2 = _.groupBy(data, "prescNo号方处");
  var data3 = [];
  var merges = [];
  s = 1;
  for (var i in data2) {
    var str = [];
    let dataItemArray = data2[i];
    var pre = 0;
    merges.push({
      s: {
        c: 12,
        r: s,
      },
      e: {
        c: 12,
        r: s + dataItemArray.length - 1,
      },
    });
    merges.push({
        s: {
            c: 13,
            r: s,
          },
          e: {
            c: 13,
            r: s + dataItemArray.length - 1,
          },
    })
    s = s + dataItemArray.length;
    for (let j = 0; j < dataItemArray.length; j++) {
      if (dataItemArray[j]["药物名称"]) {
        str.push(dataItemArray[j]["药物名称"]);
      }
      pre+=Number(dataItemArray[j].intrqty)
    }
    dataItemArray[0]['药物名称合并'] = str.join(" ");
    dataItemArray[0]['价格合并'] = pre;
    data3 = _.concat(data3, dataItemArray);
  }

  const ws = xlsx.utils.json_to_sheet(data3);
  console.log(data3[0])
  ws["!merges"] = merges;
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "处理后的数据");
  xlsx.writeFile(wb, `处理后的数据.xlsx`);
}
