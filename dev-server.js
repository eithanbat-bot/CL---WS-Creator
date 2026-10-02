const http=require('http'),fs=require('fs'),path=require('path');
const root=__dirname,port=3000;
const mime={'.html':'text/html','.js':'application/javascript','.css':'text/css','.xml':'application/xml','.svg':'image/svg+xml','.png':'image/png'};
http.createServer(function(req,res){let p=decodeURIComponent(req.url.split('?')[0]);if(p==='/')p='/taskpane.html';const f=path.join(root,p);if(!f.startsWith(root)||!fs.existsSync(f)||fs.statSync(f).isDirectory()){res.writeHead(404);return res.end('Not found')}res.writeHead(200,{'Content-Type':mime[path.extname(f)]||'application/octet-stream','Access-Control-Allow-Origin':'*'});fs.createReadStream(f).pipe(res)}).listen(port,function(){console.log('Task pane server: http://localhost:'+port)});
