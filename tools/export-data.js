const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');

const root = path.resolve(__dirname, '..');
const inputs = ['7714_1.xlsx', '7716.xlsx'];
const text = value => String(value ?? '').trim();
const number = value => value == null || text(value) === '' ? NaN : Number(text(value).replace(',', '.'));
const isNumber = value => Number.isFinite(number(value));
const seriesFrom = sheet => {
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, {header:1, defval:null, raw:true})
    .filter(row => isNumber(row[0]) && isNumber(row[1]))
    .map(row => [number(row[0]), number(row[1])]);
};

fs.mkdirSync(path.join(root, 'data'), {recursive:true});
const catalog = [];
for (const fileName of inputs) {
  const workbook = XLSX.read(fs.readFileSync(path.join(root, fileName)), {type:'buffer', dense:true});
  const firstName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstName], {header:1, defval:null, raw:true});
  const headerIndex = rows.findIndex(row => text(row[0]).toLowerCase() === 'порт' && text(row[1]).toLowerCase().includes('глуб'));
  if (headerIndex < 0) throw new Error(`${fileName}: не найдена конструкция`);
  const header = rows[headerIndex], nktIndex = header.findIndex(v => text(v).toLowerCase() === 'нкт');
  let tubingDepth = NaN;
  for (let c=nktIndex+1; c<=Math.min(nktIndex+3,header.length-1); c++) if(isNumber(header[c])) {tubingDepth=number(header[c]);break;}
  const perforations=[];
  for(let r=headerIndex+1;r<rows.length;r++){if(isNumber(rows[r][0])&&isNumber(rows[r][1]))perforations.push({port:number(rows[r][0]),depth:number(rows[r][1])});else if(perforations.length)break;}
  const ports=new Set(perforations.map(p=>p.port)), experiments=[];
  for(let r=headerIndex+perforations.length+1;r<rows.length;r++){
    const id=rows[r][0],rawName=rows[r][1];if(!isNumber(id)||typeof rawName!=='string'||!text(rawName))continue;
    const name=text(rawName),match=name.match(/^\s*(\d+)/)||name.match(/(\d+)\s*$/),candidate=match?+match[1]:null;
    experiments.push({id:number(id),name,ballPort:ports.has(candidate)?candidate:null});
  }
  const sheetMap=new Map();
  workbook.SheetNames.forEach(name=>sheetMap.set(text(name).toLowerCase().replace(/с$/,'c'),name));
  const slug=fileName.replace(/\.xlsx?$/i,''),base=path.join(root,'data',slug),expDir=path.join(base,'experiments');
  fs.mkdirSync(expDir,{recursive:true});
  fs.writeFileSync(path.join(base,'common.json'),JSON.stringify({id:slug,name:slug,sourceFile:fileName,tubingDepth,perforations}));
  fs.writeFileSync(path.join(expDir,'index.json'),JSON.stringify(experiments));
  for(const experiment of experiments){
    const sourceName=sheetMap.get(String(experiment.id));
    const cepstrumName=sheetMap.get(`${experiment.id}c`);
    const payload={...experiment,sourceSheet:sourceName||null,cepstrumSheet:cepstrumName||null,source:seriesFrom(workbook.Sheets[sourceName]),cepstrum:seriesFrom(workbook.Sheets[cepstrumName])};
    fs.writeFileSync(path.join(expDir,`${experiment.id}.json`),JSON.stringify(payload));
  }
  catalog.push({id:slug,name:slug,common:`data/${slug}/common.json`,experiments:`data/${slug}/experiments/index.json`,count:experiments.length});
  console.log(`${slug}: ${perforations.length} перфораций, ${experiments.length} испытаний`);
}
fs.writeFileSync(path.join(root,'data','catalog.json'),JSON.stringify(catalog));
