const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage();
await p.goto('file://'+process.cwd()+'/explain_phone.html');
await p.pdf({path:'NUR774_Murmurs_Circulation_Defects_PHONE.pdf',preferCSSPageSize:true,printBackground:true});await b.close();})();
