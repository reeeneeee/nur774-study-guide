const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage();
await p.goto('file://'+process.cwd()+'/'+(process.argv[2]||'guide.html')+'');
await p.pdf({path:(process.argv[3]||'NUR774_Exam3_Mnemonics.pdf'),format:'Letter',margin:{top:'0.6in',bottom:'0.6in',left:'0.7in',right:'0.7in'}});await b.close();})();
