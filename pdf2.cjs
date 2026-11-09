const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage();
await p.goto('file://'+process.cwd()+'/phone.html');
const over=await p.$$eval('.card',cs=>cs.map((c,i)=>c.scrollHeight>c.clientHeight+1?i:-1).filter(i=>i>=0));
console.log('overflowing cards:',JSON.stringify(over));
await p.pdf({path:'NUR774_Exam3_Practice_PHONE.pdf',preferCSSPageSize:true,printBackground:true});await b.close();})();
