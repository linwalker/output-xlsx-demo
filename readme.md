# output-xlsx-demo
## 简介
node端将数据生成xlsx文件，前端点击下载。

项目源码：<https://github.com/linwalker/output-xlsx-demo>

## 操作
- 下载代码
- node >= 7.8
- npm install
- node index.js
- 访问localhost: 5000

## xlsx
js-xlsx是目前处理js处理excel比较好用的一个库，生成xlsx只要调用`XLSX.writeFile（workbook, filename）`。

workbook是一个对象，结构如下：

```
//workbook
{
    SheetNames: ['sheet1'],
    Sheets: {
        // worksheet
        'sheet1': {
            '!ref': 'A1:D4',// 必须要有这个范围才能输出，否则导出的 excel 会是一个空表
            // cell
            'A1': { ... },
            // cell
            'A2': { ... },
            ...
        }
    }
}
```

## 用法
只要将需要生成xlsx文件的数据，调整成上面workbook对象格式，再调用XSLX.writeFile(workbook, filename)方法就可以。

### node 服务
node服务文件index，一个简单的页面渲染和下载请求处理。

```javascript
const Koa = require('koa');
const Router = require('koa-router');
const fs = require('fs');
const path =require('path');
const views = require('koa-views');
const dlXlsx = require('./dlXlsx.js');
const app = new Koa();
const router = new Router();

//加载模版引擎
app.use(views(path.join(__dirname, './views'), {
    extension: 'html'
}))

//页面渲染
router.get('/', async (ctx) => {
    await ctx.render('main');
})
//下载请求处理
router.get('/download', async (ctx) => {
    //生成xlsx文件
    await dlXlsx();
    //类型
    ctx.type = '.xlsx';
    //请求返回，生成的xlsx文件
    ctx.body = fs.readFileSync('output.xlsx');
    //请求返回后，删除生成的xlsx文件，不删除也行，下次请求回覆盖
    fs.unlink('output.xlsx');
})
// 加载路由中间件
app.use(router.routes()).use(router.allowedMethods())
app.listen(5000);
```
node服务基于koa2框架，node>=7.8，可以直接食用async和await处理异步请求。

### xlsx生成处理
xlsx生成在dlXlsx.js中处理

```javascript
//dlXlsx.js
const XLSX = require('xlsx');
//表头
const _headers = ['id', 'name', 'age', 'country'];
//表格数据
const _data = [
    {
        id: '1',
        name: 'test1',
        age: '30',
        country: 'China',
    },
    ...
];
const dlXlsx = () => {
    const headers = _headers
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

    const data = _data
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
    const output = Object.assign({}, headers, data);
    // 获取所有单元格的位置
    const outputPos = Object.keys(output);
    // 计算出范围
    const ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];

    // 构建 workbook 对象
    const workbook = {
        SheetNames: ['mySheet'],
        Sheets: {
            'mySheet': Object.assign({}, output, { '!ref': ref })
        }
    };
   
    // 导出 Excel
    XLSX.writeFile(workbook, 'output.xlsx')
}
```
## 参考
本文主要参考这篇文章

[在 Node.js 中利用 js-xlsx 处理 Excel 文件](https://segmentfault.com/a/1190000004395728);