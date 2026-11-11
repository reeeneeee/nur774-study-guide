const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:384,height:691}});
await p.goto('file://'+process.cwd()+'/ch38_phone.html');await p.waitForLoadState('networkidle');
const tall=await p.$$eval('.card',cs=>cs.map((c,i)=>[i,Math.round(c.scrollHeight)]).filter(([i,h])=>h>692));
console.log('cards taller than 1 page:',JSON.stringify(tall));
await p.pdf({path:'NUR774_Ch38_Cardio_100Q_PHONE.pdf',preferCSSPageSize:true,printBackground:true});await b.close();})();
