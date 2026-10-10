import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load('demo/Пример сверки.xlsx'));
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:"'ДОБАВЛЕННЫЕ'!A1:H5",tableMaxRows:5,tableMaxCols:8,maxChars:2600})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
for(const name of ['ДОБАВЛЕННЫЕ','МИНУСОВЫЕ И ИСКЛЮЧЁННЫЕ','ФАКТИЧЕСКИ ВЫПОЛНЕННЫЕ']){
  const image=await wb.render({sheetName:name,range:'A1:J5',scale:1.2});
  await fs.writeFile('qa/'+name+'.png',new Uint8Array(await image.arrayBuffer()));
}
console.log('Workbook views rendered');
