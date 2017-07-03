// 由写死的数据生成xlsx文件
let XLSX = require('xlsx');
const dlXlsx = () => {
    var _headers = ['id', 'name', 'age', 'country']
    var _data = [{
        id: '1',
        name: 'test1',
        age: '30',
        country: 'China',
    },
    {
        id: '2',
        name: 'test2',
        age: '20',
        country: 'America',
    },
    {
        id: '3',
        name: 'test3',
        age: '18',
        country: 'UK',
    }];

    var headers = _headers
        
        .map((v, i) => Object.assign({}, { v: v, position: String.fromCharCode(65 + i) + 1 }))
        // 为 _headers 添加对应的单元格位置
        // [ { v: 'id', position: 'A1' },
        //   { v: 'name', position: 'B1' },
        //   { v: 'age', position: 'C1' },
        //   { v: 'country', position: 'D1' },
        .reduce((prev, next) => Object.assign({}, prev, { [next.position]: { v: next.v } }), {});
        // 转换成 worksheet 需要的结构
        // { A1: { v: 'id' },
        //   B1: { v: 'name' },
        //   C1: { v: 'age' },
        //   D1: { v: 'country' },

    var data = _data

        .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65 + j) + (i + 2) })))
        // 匹配 headers 的位置，生成对应的单元格数据
        // [ [ { v: '1', position: 'A2' },
        //     { v: 'test1', position: 'B2' },
        //     { v: '30', position: 'C2' },
        //     { v: 'China', position: 'D2' }],
        //   [ { v: '2', position: 'A3' },
        //     { v: 'test2', position: 'B3' },
        //     { v: '20', position: 'C3' },
        //     { v: 'America', position: 'D3' }],
        //   [ { v: '3', position: 'A4' },
        //     { v: 'test3', position: 'B4' },
        //     { v: '18', position: 'C4' },
        //     { v: 'Unkonw', position: 'D4' }] ]
        .reduce((prev, next) => prev.concat(next))
        // 对刚才的结果进行降维处理（二维数组变成一维数组）
        // [ { v: '1', position: 'A2' },
        //   { v: 'test1', position: 'B2' },
        //   { v: '30', position: 'C2' },
        //   { v: 'China', position: 'D2' },
        //   { v: '2', position: 'A3' },
        //   { v: 'test2', position: 'B3' },
        //   { v: '20', position: 'C3' },
        //   { v: 'America', position: 'D3' },
        //   { v: '3', position: 'A4' },
        //   { v: 'test3', position: 'B4' },
        //   { v: '18', position: 'C4' },
        //   { v: 'Unkonw', position: 'D4' },
        .reduce((prev, next) => Object.assign({}, prev, { [next.position]: { v: next.v } }), {});
        // 转换成 worksheet 需要的结构
        //   { A2: { v: '1' },
        //     B2: { v: 'test1' },
        //     C2: { v: '30' },
        //     D2: { v: 'China' },
        //     A3: { v: '2' },
        //     B3: { v: 'test2' },
        //     C3: { v: '20' },
        //     D3: { v: 'America' },
        //     A4: { v: '3' },
        //     B4: { v: 'test3' },
        //     C4: { v: '18' },
        //     D4: { v: 'Unkonw' }
    // 合并 headers 和 data
    var output = Object.assign({}, headers, data);
    // 获取所有单元格的位置
    var outputPos = Object.keys(output);
    // 计算出范围
    var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];

    // 构建 workbook 对象
    var wb = {
        SheetNames: ['mySheet'],
        Sheets: {
            'mySheet': Object.assign({}, output, { '!ref': ref })
        }
    };
   
    // 导出 Excel
    XLSX.writeFile(wb, 'output.xlsx')
}

module.exports = dlXlsx;