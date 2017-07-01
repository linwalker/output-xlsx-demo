const Koa = require('koa');
const Router = require('koa-router');
const fs = require('fs');
const app = new Koa();
const router = new Router();

let XLSX = require('xlsx');
const dlXlsx = () => {
    var _headers = ['id', 'name', 'age', 'country', 'remark']
    var _data = [{
        id: '1',
        name: 'test1',
        age: '30',
        country: 'China',
        remark: 'hello'
    },
    {
        id: '2',
        name: 'test2',
        age: '20',
        country: 'America',
        remark: 'world'
    },
    {
        id: '3',
        name: 'test3',
        age: '18',
        country: 'Unkonw',
        remark: '???'
    }];

    var headers = _headers
        
        .map((v, i) => Object.assign({}, { v: v, position: String.fromCharCode(65 + i) + 1 }))

        .reduce((prev, next) => Object.assign({}, prev, { [next.position]: { v: next.v } }), {});

    var data = _data

        .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65 + j) + (i + 2) })))
        
        .reduce((prev, next) => prev.concat(next))

        .reduce((prev, next) => Object.assign({}, prev, { [next.position]: { v: next.v } }), {});

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
router.get('/', async (ctx) => {
    ctx.body = 'output-xslx-demo';
})
router.get('/download', async (ctx) => {
    // ctx.body = 'download';
    await dlXlsx();
    ctx.type = '.xlsx';
    ctx.body = fs.readFileSync('output.xlsx');
    fs.unlink('output.xlsx');
    console.log('delete');
})
// 加载路由中间件
app.use(router.routes()).use(router.allowedMethods())
app.listen(5000);