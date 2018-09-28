const express = require('express');
const bodyParser = require('body-parser');
const app = express();

app.set('view engine','ejs');
app.use(bodyParser.urlencoded({ extended: true }));

const arr = ['hello', 'world', 'test'];

app.get('/', (req, res) => res.render('index',{arr: arr}));

app.get('/create', (req, res) => res.render('create'));
app.post('/create', (req, res) => {
    console.log(req.body);
});

app.listen(3000, () => console.log('port 3000'));
