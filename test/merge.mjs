// Proves merge-by-default: filing July alone then dropping August must land in
// exactly the same place as dropping both at once — and re-dropping a statement
// that is already filed must change nothing.
import { chromium } from 'playwright';
import http from 'node:http'; import fs from 'node:fs'; import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MIME={'.html':'text/html','.css':'text/css','.js':'text/javascript','.mjs':'text/javascript','.json':'application/json','.pdf':'application/pdf'};
const server=http.createServer((q,s)=>{const f=path.join(ROOT,decodeURIComponent(q.url.split('?')[0]).replace(/^\//,''));if(!fs.existsSync(f)||fs.statSync(f).isDirectory())return s.writeHead(404).end();s.writeHead(200,{'Content-Type':MIME[path.extname(f)]||'application/octet-stream'});fs.createReadStream(f).pipe(s);});
await new Promise(r=>server.listen(0,r));
const base=`http://127.0.0.1:${server.address().port}`;
const JUL='samples/Rogers Jul 2026 Statement.pdf', AUG='samples/Rogers Aug 2026 Statement.pdf';

const browser=await chromium.launch({executablePath:'/opt/pw-browsers/chromium-1194/chrome-linux/chrome'});
const page=await browser.newPage();
page.on('pageerror',e=>console.log('[pageerror]',e.message));
await page.goto(base+'/index.html',{waitUntil:'domcontentloaded'});
await page.evaluate(j=>localStorage.setItem('statement-filer.rules',j), fs.readFileSync(path.join(ROOT,'my-rules.json'),'utf8'));
await page.reload({waitUntil:'domcontentloaded'});
await page.waitForFunction(()=>!!window.ExcelJS&&!!window.PDFLib,null,{timeout:30000});

// a stand-in for Drive: whatever the previous run "filed" lives in page memory
await page.evaluate(()=>{ window.__filed=null; });

async function runWith(files, useFiled) {
  await page.evaluate((u)=>{ window.__useFiled=u; }, useFiled);
  return page.evaluate(async ({files, urlbase}) => {
    const pipeline = await import('/js/pipeline.js');
    const rulesMod = await import('/js/rules.js');
    const pdfjsLib = await import('/vendor/pdf.min.mjs');
    pdfjsLib.GlobalWorkerOptions.workerSrc = urlbase + '/vendor/pdf.worker.min.mjs';
    const rules = await rulesMod.loadRules();
    const fileObjs = [];
    for (const f of files) {
      const b = await (await fetch(urlbase + '/' + encodeURI(f))).arrayBuffer();
      fileObjs.push(new File([b], f.split('/').pop(), {type:'application/pdf'}));
    }
    const opts = window.__useFiled && window.__filed
      ? { fetchFiled: async () => window.__filed }
      : {};
    const res = await pipeline.run(fileObjs, rules,
      { pdfjsLib, ExcelJS: window.ExcelJS, PDFLib: window.PDFLib }, () => {}, opts);
    // "file" it: keep the workbook + category pdfs the way Drive would
    const catPdfs = {};
    for (const c of res.categoryPdfs) catPdfs[c.category] = c.bytes.slice().buffer;
    window.__filed = { workbookBytes: res.workbook.bytes.slice().buffer, categoryPdfs: catPdfs };
    return {
      n: res.transactions.length, total: res.parsedTotal, control: res.controlTotal,
      variance: res.variance, summary: res.summary, merged: res.merged,
      statements: res.statements.map(s=>s.label),
      pdfs: res.categoryPdfs.map(c=>`${c.category}:${c.lines}(+${c.carriedPages}p)`),
    };
  }, {files, urlbase: base});
}

console.log('--- A. both statements in one go (the baseline)');
const both = await runWith([JUL,AUG], false);
console.log(`   ${both.n} txns  $${both.total}  variance ${both.variance}  statements: ${both.statements}`);
console.log('   ', JSON.stringify(both.summary));

await page.evaluate(()=>{ window.__filed=null; });
console.log('\n--- B. July alone, filed');
const jul = await runWith([JUL], true);
console.log(`   ${jul.n} txns  $${jul.total}  variance ${jul.variance}  statements: ${jul.statements}`);

console.log('\n--- C. August dropped a month later, merging into the filed July');
const aug = await runWith([AUG], true);
console.log(`   ${aug.n} txns  $${aug.total}  variance ${aug.variance}  statements: ${aug.statements}`);
console.log(`   merged: carried ${aug.merged.carried}, added ${aug.merged.added}, skipped ${JSON.stringify(aug.merged.skipped)}`);
console.log('   ', JSON.stringify(aug.summary));
console.log('    pdfs:', aug.pdfs.join(', '));

console.log('\n--- D. August dropped AGAIN (must be a no-op)');
const again = await runWith([AUG], true);
console.log(`   ${again.n} txns  $${again.total}  variance ${again.variance}`);
console.log(`   merged: carried ${again.merged.carried}, added ${again.merged.added}, skipped ${JSON.stringify(again.merged.skipped)}`);

let bad=0;
const eq=(a,b,l)=>{const ok=Math.abs(a-b)<0.005; console.log(`  ${ok?'ok  ':'FAIL'} ${l}: ${a} vs ${b}`); if(!ok)bad++;};
console.log('\nmerged (C) must equal baseline (A):');
eq(aug.n, both.n, 'transaction count');
eq(aug.total, both.total, 'parsed total');
eq(aug.control, both.control, 'control total');
eq(aug.variance, 0, 'variance');
for (const k of new Set([...Object.keys(both.summary),...Object.keys(aug.summary)])) eq(aug.summary[k]||0, both.summary[k]||0, k);
console.log('\nre-drop (D) must be a no-op:');
eq(again.n, both.n, 'transaction count');
eq(again.total, both.total, 'parsed total');
eq(again.variance, 0, 'variance');
console.log(`  ${again.merged.added===0?'ok  ':'FAIL'} nothing added: ${again.merged.added}`); if(again.merged.added!==0)bad++;

await browser.close(); server.close();
console.log(bad?`\n${bad} check(s) failed`:'\nMerge behaves: a quarter built in two drops equals one built in a single drop.');
process.exit(bad?1:0);
