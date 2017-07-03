const Koa = require('koa');
const Router = require('koa-router');
const fs = require('fs');
const path =require('path');
const views = require('koa-views');
const dlXlsx = require('./dlXlsx.js');
const app = new Koa();
const router = new Router();

app.use(views(path.join(__dirname, './views'), {
    extension: 'html'
}))


router.get('/', async (ctx) => {
    await ctx.render('main');
})
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