import fs from 'node:fs/promises';
import {demoBytes} from '../src/fixture.mjs';
import {importWorkbook,detach,analyze} from '../src/core.mjs';
await fs.mkdir('generated',{recursive:true});
const source=await demoBytes(),workbook=await importWorkbook(source),result=await detach(workbook,['Plan']);
await fs.writeFile('generated/source.xlsx',source);await fs.writeFile('generated/handoff.xlsx',result.xlsx);await fs.writeFile('generated/ledger.csv',result.ledgerCsv);await fs.writeFile('generated/recipe.json',JSON.stringify(result.recipe,null,2)+'\n');await fs.writeFile('generated/report.txt',result.report);const guard=await demoBytes({backlink:true});await fs.writeFile('generated/guard.xlsx',guard);await fs.writeFile('generated/guard-analysis.json',JSON.stringify(analyze(await importWorkbook(guard),['Plan']),null,2)+'\n');console.log(JSON.stringify({values:result.analysis.values,freezeOccurrences:result.analysis.ledger.length}));
