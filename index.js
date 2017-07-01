const Koa = require('koa');
const app = new Koa();

app.use(ctx => {
    ctx.body = 'output-xlsx-demo';
})

app.listen(5000);