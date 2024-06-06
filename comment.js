// Create web server
const http = require('http');
const fs = require('fs');
const url = require('url');
const qs = require('querystring');

const app = http.createServer((request, response) => {
    const _url = request.url;
    const queryData = url.parse(_url, true).query;
    const pathname = url.parse(_url, true).pathname;

    if (pathname === '/'){
        if (queryData.id === undefined){
            fs.readdir('./data', (err, filelist) => {
                var title = 'Welcome';
                var description = 'Hello, Node.js';
                var list = templateList(filelist);
                var html = templateHTML(title, list, `<h2>${title}</h2>${description}`, `<a href="/create">create</a>`);
                response.writeHead(200);
                response.end(html);
            });
        } else {
            fs.readdir('./data', (err, filelist) => {
                fs.readFile(`data/${queryData.id}`, 'utf8', (err, description) => {
                    var title = queryData.id;
                    var list = templateList(filelist);
                    var html = templateHTML(title, list, `<h2>${title}</h2>${description}`, `
                        <a href="/create">create</a>
                        <a href="/update?id=${title}">update</a>
                        <form action="/delete_process" method="post">
                            <input type="hidden" name="id" value="${title}">
                            <input type="submit" value="delete">
                        </form>
                    `);
                    response.writeHead(200);
                    response.end(html);
                });
            });
        }
    } else if (pathname === '/create') {
        fs.readdir('./data', (err, filelist) => {
            var title = 'WEB - create';
            var list = templateList(filelist);
            var html = templateHTML(title, list, `
                <form action="/create_process" method="post">
                    <p><input type="text" name="title" placeholder="title"></p>
                    <p><textarea name="description" placeholder="description"></textarea></p>
                    <p><input type="submit"></p>
                </form>
            `, '');
            response.writeHead(200);
            response.end(html);
        });
    } else if (pathname === '/create_process') {
        var body = '';
        request.on('data',