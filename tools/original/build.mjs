import fs from 'node:fs/promises';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

// Scope: a new applicant workbook only; existing application data is never read.
const out = new URL('.', import.meta.url).pathname;
const wb = Workbook.create();
const names=['Instructions','Personal & contacts','Degrees','Exams','Work experience','Language skills','Internationalization','Training','Other activities','Degree class legend','Exam sector legend','Lists'];
const colors=['334155','2563EB','7C3AED','9333EA','0F766E','C2410C','BE185D','4D7C0F','A16207','64748B','475569','94A3B8'];
const sheets=Object.fromEntries(names.map((n,i)=>{const s=wb.worksheets.add(n);s.tabColor='#'+colors[i];s.showGridLines=false;return[n,s];}));
const meta={sheets:{},tables:[],validations:[],unlocked:[],lists:{},formulaCells:[],inheritedDefaults:[]};
const C=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const text=(h,e='',hint='',req=false)=>({h,e,hint,req,type:'text'});
const num=(h,e=0,hint='',req=false)=>({h,e,hint,req,type:'decimal'});
const date=(h,e='2025-06-30',hint='',req=false)=>({h,e,hint,req,type:'date'});
const list=(h,values,e=values[0],hint='',req=false)=>({h,values,e,hint,req,type:'list'});
const calc=(h,fn,hint='Automatic; do not edit')=>({h,fn,hint,type:'formula',e:'Calculated'});
const yes=['Yes','No'];
const countries='AD AE AF AG AI AL AM AO AQ AR AS AT AU AW AX AZ BA BB BD BE BF BG BH BI BJ BL BM BN BO BQ BR BS BT BV BW BY BZ CA CC CD CF CG CH CI CK CL CM CN CO CR CU CV CW CX CY CZ DE DJ DK DM DO DZ EC EE EG EH ER ES ET FI FJ FK FM FO FR GA GB GD GE GF GG GH GI GL GM GN GP GQ GR GS GT GU GW GY HK HM HN HR HT HU ID IE IL IM IN IO IQ IR IS IT JE JM JO JP KE KG KH KI KM KN KP KR KW KY KZ LA LB LC LI LK LR LS LT LU LV LY MA MC MD ME MF MG MH MK ML MM MN MO MP MQ MR MS MT MU MV MW MX MY MZ NA NC NE NF NG NI NL NO NP NR NU NZ OM PA PE PF PG PH PK PL PM PN PR PS PT PW PY QA RE RO RS RU RW SA SB SC SD SE SG SH SI SJ SK SL SM SN SO SR SS ST SV SX SY SZ TC TD TF TG TH TJ TK TL TM TN TO TR TT TV TW TZ UA UG UM US UY UZ VA VC VE VG VI VN VU WF WS YE YT ZA ZM ZW'.split(' ').map(c=>`${c} — ${new Intl.DisplayNames(['en'],{type:'region'}).of(c)}`).sort((a,b)=>a.slice(5).localeCompare(b.slice(5))).concat(['Other']);
const level=['A1','A2','B1','B2','C1','C2','Not assessed','Other'];
const country=h=>list(h,countries,'IT — Italy','Select country; use Other details if absent.',true);
const evidence=[text('Evidence filename','transcript.pdf','Exact attachment filename; separate files with ;',true),text('Page / section','2–4','Locate the supporting passage.',true)];
const other=text('Other details','','Required when any category is Other.');
const repeatedFields={
 Degrees:['Institution country','Credit system','Hours per credit','Final scale minimum','Final pass threshold','Final scale maximum','Final scale direction','Exam scale minimum','Exam pass threshold','Exam scale maximum','Exam scale direction'],
 Exams:['Degree ID','Grading type','Scale minimum','Pass threshold','Scale maximum','Scale direction'],
 'Work experience':['Country'],Internationalization:['Host country'],Training:['Country'],'Other activities':['Country']
};
const formula=(s,cell,f)=>{s.getRange(cell).formulas=[[f]];meta.formulaCells.push([s.name,cell,f]);};
function heading(s,title,note,end='H'){
 s.getRange('A1:'+end+'4').format.font={name:'Arial',size:11,color:'#172033'};
 s.getRange('A2:'+end+'2').merge();s.getRange('A2').values=[[title]];s.getRange('A2').format.font={name:'Arial',size:16,bold:true,color:'#172033'};
 s.getRange('A3:'+end+'4').merge();s.getRange('A3').values=[[note]];s.getRange('A3:'+end+'4').format.wrapText=true;s.getRange('A3:'+end+'4').format.fill='#DBEAFE';s.getRange('A3:'+end+'4').format.rowHeight=16;
 meta.sheets[s.name]={color:s.tabColor,freeze:0};
}
function validation(sheet,range,rule,description='Use the specified format or choose a listed value.'){
 meta.validations.push({sheet,range,rule,description});
}
function addList(values){
 const key=JSON.stringify(values);if(meta.lists[key])return meta.lists[key].name;
 const col=Object.keys(meta.lists).length;const name='Options'+(col+1);
 sheets.Lists.getRange(`${C(col)}1:${C(col)}${values.length}`).values=values.map(v=>[v]);
 meta.lists[key]={name,ref:`'Lists'!$${C(col)}$1:$${C(col)}$${values.length}`};return name;
}
function table(sheet,name,cols,start,count,prefix,notes=''){
 const s=sheets[sheet],end=start+count-1,last=C(cols.length-1),header=start-1;
 s.getRange(`A${header}:${last}${end}`).format.font={name:'Arial',size:11,color:'#172033'};
 s.getRange(`A${header}:${last}${end}`).format.rowHeight=27;
 s.getRange(`A${start}:${last}${end}`).format.fill='#FFF2A6';
 s.getRange(`A${start}:${last}${end}`).format.verticalAlignment='center';
 s.getRange(`A${start}:${last}${end}`).format.wrapText=true;
 s.getRange(`A${header}:${last}${header}`).values=[cols.map(x=>x.h+(x.req?' *':''))];
 s.getRange(`A${header}:${last}${header}`).format={fill:'#23374D',font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},wrapText:true,horizontalAlignment:'center',verticalAlignment:'center',rowHeight:38};
 const t=s.tables.add(`A${header}:${last}${end}`,true,name);t.style='TableStyleLight1';
 const sample=start-3,hints=start-2;
 s.getRange(`A${sample}:${last}${sample}`).setNumberFormat('@');
 s.getRange(`A${sample}:${last}${sample}`).values=[cols.map(x=>typeof x.e==='string'&&/^0\d+$/.test(x.e)?"'"+x.e:x.e)];
 s.getRange(`A${sample}:${last}${sample}`).format={fill:'#FCE7F3',font:{name:'Arial',size:10,italic:true,color:'#334155'},wrapText:true,rowHeight:28};
 s.getRange(`A${hints}:${last}${hints}`).values=[cols.map(x=>x.hint||'')];
 s.getRange(`A${hints}:${last}${hints}`).format={fill:'#DBEAFE',font:{name:'Arial',size:10,color:'#334155'},wrapText:true,rowHeight:44};
 const map=Object.fromEntries(cols.map((x,i)=>[x.h,C(i)]));
 cols.forEach((x,i)=>{
  const c=C(i),range=`${c}${start}:${c}${end}`;s.getRange(c+':'+c).format.columnWidth=x.h.includes('relevant')||x.h.toLowerCase().includes('description')?40:x.h.includes('cluster')?32:x.h==='ID'?10:x.h==='Check'?38:x.type==='date'?14:x.type==='decimal'?17:22;
  if(x.type==='formula'){
   s.getRange(range).format.fill='#E8EDF2';
   s.getRange(range).formulas=Array.from({length:count},(_,j)=>[x.fn(start+j,map,start,end)]);
   meta.formulaCells.push([sheet,range,'column']);
  }else{
   meta.unlocked.push({sheet,range});
   if(x.type==='list')validation(sheet,range,{type:'list',formula1:x.ref||addList(x.values)},x.hint);
   else if(x.type==='date'){s.getRange(range).setNumberFormat('yyyy-mm-dd');validation(sheet,range,{type:'date',operator:'between',formula1:'DATE(1900,1,1)',formula2:'DATE(2100,12,31)'},'Enter a real Excel date. Display: yyyy-mm-dd.');}
   else if(x.type==='decimal'){s.getRange(range).setNumberFormat('0.##');validation(sheet,range,{type:x.h==='Remaining exams'?'whole':'decimal',operator:'between',formula1:/mark|minimum|threshold|maximum|average|score/i.test(x.h)?-1000000:0,formula2:1000000});}
   else{s.getRange(range).setNumberFormat('@');validation(sheet,range,{type:'textLength',operator:'lessThanOrEqual',formula1:x.h.includes('relevant')?300:250});}
   if((repeatedFields[sheet]||[]).includes(x.h)){
    const target=`${c}${start+1}:${c}${end}`;
    s.getRange(target).formulas=Array.from({length:count-1},(_,j)=>{
     const r=start+j+1,previous=`${c}${r-1}`;
     // Different degrees may have different grading systems. Ask again after a degree change.
     const changedDegree=sheet==='Exams'&&x.h!=='Degree ID'?`$C${r}<>$C${r-1},`:'';
     return [`=IF(OR($B${r}="",${changedDegree}${previous}=""),"",${previous})`];
    });
    meta.inheritedDefaults.push({sheet,range:target,field:x.h});
   }
  }
 });
 s.getRange(`A${header}:${last}${end}`).format.borders={preset:'all',style:'thin',color:'#AAB6C4'};
 const check=cols.findIndex(c=>c.h==='Check');if(check>=0)s.getRange(`${C(check)}${start}:${C(check)}${end}`).conditionalFormats.addCustom(`${C(check)}${start}<>""`,{fill:'#FCE3E3',font:{color:'#9C1C23',bold:true}});
 meta.tables.push({sheet,name,start,end,header,cols:cols.map(x=>({h:x.h,type:x.type,req:x.req})),map});
 meta.sheets[sheet].freeze=9;s.freezePanes.freezeRows(9);
 return map;
}
function id(prefix){return calc('ID',(r,m,start)=>`=IF(B${r}="","","${prefix}"&TEXT(ROW()-${start-1},"000"))`);}
function checkedTable(sheet,name,cols,start,count,extra=[]){
 const original=[...cols];
 cols.push(calc('Check',(r,m)=>{
  const v=h=>`${m[h]}${r}`;
  const terms=original.filter(x=>x.req).map(x=>`IF(${v(x.h)}="","Missing ${x.h}; ","")`);
  for(const x of original){
   if(x.type==='date')terms.push(`IF(AND(${v(x.h)}<>"",NOT(ISNUMBER(${v(x.h)}))),"Invalid date: ${x.h}; ","")`);
   if(x.type==='decimal')terms.push(`IF(AND(${v(x.h)}<>"",NOT(ISNUMBER(${v(x.h)}))),"Invalid number: ${x.h}; ","")`);
   if(x.type==='list')terms.push(`IF(AND(${v(x.h)}<>"",COUNTIF(${x.ref||addList(x.values)},${v(x.h)})=0),"Invalid choice: ${x.h}; ","")`);
  }
  if(m['Other details'])terms.push(`IF(AND(COUNTIF(A${r}:${C(original.length-1)}${r},"Other")>0,${v('Other details')}=""),"Specify Other; ","")`);
  if(m['Start date']){
   const starts=`$${m['Start date']}$${start}:$${m['Start date']}$${start+count-1}`;
   const ends=`$${m['End date']}$${start}:$${m['End date']}$${start+count-1}`;
   const ongoing=`$${m['Ongoing?']}$${start}:$${m['Ongoing?']}$${start+count-1}`;
   const until=`MIN(IF(${v('Ongoing?')}="Yes",${reference},${v('End date')}),${reference})`;
   terms.push(`IF(AND(ISNUMBER(${v('Start date')}),${reference}<>"",COUNTIFS(${starts},"<="&${until},${ends},">="&${v('Start date')})+COUNTIFS(${starts},"<="&${until},${ends},"",${ongoing},"Yes")>1),"Overlapping dates: committee review; ","")`);
  }
  terms.push(...extra.map(f=>f(r,m,v)));
  return `=IF(B${r}="","",${terms.length?terms.join('&'):'""'})`;
 }));
 return table(sheet,name,cols,start,count);
}
function declare(s){s.getRange('A5').values=[['Status *']];s.getRange('B5').values=[['Not started']];s.getRange('B5').format.fill='#FFF2A6';meta.unlocked.push({sheet:s.name,range:'B5'});validation(s.name,'B5',{type:'list',formula1:addList(['Not started','Provided','None'])});s.getRange('C5:H5').merge();s.getRange('C5').values=[['Choose Provided if entering rows; otherwise choose None.']];s.getRange('C5:H5').format.fill='#DBEAFE';}

// Degree-class and discipline dictionaries use the existing committee rubric.
const degreeRows=[
 ['L-16','Economics & business','Administration and organisation sciences',"Scienze dell'amministrazione e dell'organizzazione",'Public administration, organisational processes, public services and management.'],
 ['L-18','Economics & business','Economics and business management sciences',"Scienze dell'economia e della gestione aziendale",'Accounting, business administration, finance, marketing and company management.'],
 ['L-33','Economics & business','Economic sciences','Scienze economiche','Economic theory, markets, economic policy and quantitative economic analysis.'],
 ['L-41','Statistics','Statistics','Statistica','Probability, statistical inference, data analysis and statistical modelling.'],
 ['L-31','Computing & information engineering','Computer science and technologies','Scienze e tecnologie informatiche','Programming, algorithms, software, databases, computing systems and theory.'],
 ['L-8','Computing & information engineering','Information engineering',"Ingegneria dell'informazione",'Computer, electronic, telecommunications and automation engineering.'],
 ['L-9','Industrial engineering','Industrial engineering','Ingegneria industriale','Mechanical, energy, manufacturing, chemical and other industrial engineering.'],
 ['Other','Other','Other qualification','Altro titolo','Specify the official class or qualification in Other details.'],
 ['Unknown','Unknown','Class not known','Classe non nota','Use if the official class is unknown. Do not guess.'],
 ['Not applicable','Not applicable','No Italian degree class','Nessuna classe italiana','Use for a foreign qualification without an official Italian degree class. Select a suggested cluster separately.']
];
const sectorRows=[
 ['SECS-P/07','Business & management','Business administration and accounting','Economia aziendale','Accounting, financial statements, management accounting and company administration.'],
 ['SECS-P/08','Business & management','Management','Economia e gestione delle imprese','Business strategy, marketing, operations and company management.'],
 ['SECS-P/09','Business & management','Corporate finance','Finanza aziendale','Company valuation, investment decisions, financing and capital structure.'],
 ['SECS-P/10','Business & management','Organisation and human resources','Organizzazione aziendale','Organisation design, people management and organisational behaviour.'],
 ['SECS-P/11','Business & management','Financial markets and institutions','Economia degli intermediari finanziari','Banking, financial intermediaries, financial services and markets.'],
 ['INF/01','Computer science','Computer science','Informatica','Algorithms, programming, software, databases, AI and computing foundations.'],
 ['ING-INF/05','Computer science','Information processing systems','Sistemi di elaborazione delle informazioni','Computer systems, software engineering, information systems and computing architecture.'],
 ['SECS-P/01','Economics','Economics','Economia politica','Microeconomics, macroeconomics, economic theory and market behaviour.'],
 ['SECS-P/02','Economics','Economic policy','Politica economica','Economic policy design and analysis, growth and policy interventions.'],
 ['SECS-P/03','Economics','Public economics','Scienza delle finanze','Taxation, public spending, welfare and public-sector economics.'],
 ['SECS-P/05','Economics','Econometrics','Econometria','Statistical estimation, causal inference and quantitative models for economics.'],
 ['SECS-P/12','Economics','Economic history','Storia economica','Historical development of economies, firms, institutions and economic systems.'],
 ['Other','Other','Other discipline','Altro settore','Use for any sector outside this rubric, including newer or foreign codes. Enter its code and name in Other details.'],
 ['Unknown','Unknown','Discipline uncertain','Settore non noto','Use if you cannot identify the discipline. The committee will review it.'],
 ['Not applicable','Not applicable','No classification','Non applicabile','Use only where a disciplinary classification does not apply.']
];
const clusters=rows=>[...new Set(rows.map(x=>x[1]))].filter(x=>!['Unknown','Not applicable'].includes(x));
const degreeGuides={
 'Economics & business':'Administration, accounting, finance, business management, economic theory and markets.',
 'Statistics':'Probability, statistical inference, data analysis and statistical modelling.',
 'Computing & information engineering':'Programming, algorithms, software, databases, computer systems, electronics, telecommunications and automation.',
 'Industrial engineering':'Mechanical, energy, manufacturing, chemical and other industrial engineering.',
 'Other':'No cluster above adequately describes this degree. Explain its field in Other details.'
};
const sectorGuides={
 'Business & management':'Accounting, company administration, strategy, marketing, operations, corporate finance, organisation, human resources, banking and financial institutions.',
 'Computer science':'Programming, algorithms, software engineering, databases, AI, computer architecture and information processing systems.',
 'Economics':'Microeconomics, macroeconomics, economic policy, taxation, public spending, econometrics and economic history.',
 'Other':'No cluster above adequately describes this exam. Explain its subject in Other details.'
};
for(const [sheet,rows,guides] of [['Degree class legend',degreeRows,degreeGuides],['Exam sector legend',sectorRows,sectorGuides]]){
 const s=sheets[sheet];heading(s,sheet+' — read only','If your document states a listed code, select its cluster. Otherwise select the closest cluster from these descriptions. Use Other only if no cluster fits.','E');
 const grouped=clusters(rows).map(cluster=>{
  const members=rows.filter(r=>r[1]===cluster&&r[0]!=='Other');
  return [cluster,members.map(r=>r[0]).join(', '),members.map(r=>r[2]).join('; '),members.map(r=>r[3]).join('; '),guides[cluster]];
 });
 s.getRange('A6:E6').values=[['Cluster to select','Included Italian codes','English names','Italian names','Plain-language guide']];
 s.getRange('A7:E'+(6+grouped.length)).values=grouped;
 s.getRange('A6:E'+(6+grouped.length)).format={font:{name:'Arial',size:11},wrapText:true,verticalAlignment:'center'};
 s.getRange('A6:E6').format={fill:'#23374D',font:{bold:true,color:'#FFFFFF'},rowHeight:30};
 s.getRange('A:A').format.columnWidth=31;s.getRange('B:B').format.columnWidth=25;s.getRange('C:D').format.columnWidth=43;s.getRange('E:E').format.columnWidth=62;
 grouped.forEach((row,i)=>{s.getRange('A'+(7+i)+':E'+(7+i)).format.rowHeight=Math.max(34,Math.ceil(Math.max(row[2].length,row[3].length)/49)*14+8);});
 s.getRange('E7:E'+(6+grouped.length)).format.fill='#DBEAFE';
 s.getRange('A6:E'+(6+grouped.length)).format.borders={preset:'all',style:'thin',color:'#AAB6C4'};
 s.getRange('A'+(9+grouped.length)+':E'+(10+grouped.length)).merge();
 s.getRange('A'+(9+grouped.length)).values=[['Source: scripts/score_degrees/rubric.yaml. Cluster descriptions are guidance for applicant self-classification; the committee verifies the selection.']];
 s.getRange('A'+(9+grouped.length)+':E'+(10+grouped.length)).format.wrapText=true;
 meta.sheets[sheet].freeze=6;
}
const p=sheets['Personal & contacts'];heading(p,'Personal information and contacts','Fill every required yellow cell. Write names in full. Expand national abbreviations, even when your ID uses them (MD → Mohammed). * = required.','D');
const personal=[
 ['First name(s) *','Alex','Write in full; expand national abbreviations, e.g. MD → Mohammed, even if used on your ID.'],['Surname(s) *','Example','Write in full; expand national abbreviations, e.g. MD → Mohammed, even if used on your ID.'],['Birth date *','2001-04-15','Real date; displayed yyyy-mm-dd.'],['Birth place *','Bologna','City / region / country.'],['Citizenship *','IT — Italy','Select your main citizenship.'],['Additional citizenships','','Country names; separate with ;, or leave blank.'],['Home country *','IT — Italy','Country where you have spent most of your life.'],['Residential address *','12 Example Street','Street and building number.'],['City *','Bologna','Current residence.'],['Postal code *','00123','Text; keep leading zeros.'],['Residence country *','IT — Italy','Current residence.'],['Phone *','+39 0123456789','Include international calling code.'],['Personal email *','alex@example.invalid','Email you check regularly.'],['University email','','If already assigned.'],['Student / application number','','If already assigned; text, not a number.'],['Applicant category *','EU candidate','Select the category used in your application.'],['Other category details','','Complete if any selected category is Other.'],['ID document type *','Passport','Select document type.'],['ID document number *','AB0012345','Keep letters and leading zeros.'],['ID issuing country *','IT — Italy','Country that issued the document.'],['ID expiry date *','2030-04-15','Real date; displayed yyyy-mm-dd.'],['ID evidence filename *','identity.pdf','Exact submitted filename.'],['ID page / side *','1–2','Include all relevant sides.'],['Permit / visa applicable? *','No','If applicable, complete the next three fields.'],['Permit / visa expiry','','Leave blank only if not applicable or no expiry is stated.'],['Permit evidence filename','','Required when permit / visa is applicable.'],['Permit page / section','','Required when permit / visa is applicable.'],['Reference date *','2026-11-20','Use the application deadline. All ongoing durations stop on this date.'],['Accuracy confirmed? *','Yes','Confirm that your entries match the submitted evidence.']
];
p.getRange('A6:D6').values=[['Field','Your answer','Expected format / guidance','Example (do not copy)']];
p.getRange('A7:D35').values=personal.map(x=>[x[0],null,x[2],/^0\d+$/.test(x[1])?"'"+x[1]:x[1]]);
p.getRange('A6:D35').format={font:{name:'Arial',size:11,color:'#172033'},rowHeight:28,wrapText:true,verticalAlignment:'center'};p.getRange('A6:D6').format={fill:'#23374D',font:{bold:true,color:'#FFFFFF'},rowHeight:34};
p.getRange('B7:B35').format.fill='#FFF2A6';p.getRange('B7:B35').setNumberFormat('@');meta.unlocked.push({sheet:p.name,range:'B7:B35'});
p.getRange('C7:C35').format.fill='#DBEAFE';p.getRange('D7:D35').format.fill='#FCE7F3';p.getRange('C7:C8').format.rowHeight=36;p.getRange('A:A').format.columnWidth=29;p.getRange('B:B').format.columnWidth=30;p.getRange('C:C').format.columnWidth=64;p.getRange('D:D').format.columnWidth=27;p.tables.add('A6:D35',true,'PersonalDetails');
p.getRange('A6:D35').format.borders={preset:'all',style:'thin',color:'#AAB6C4'};
const reference="'Personal & contacts'!$B$34";
const P=Object.fromEntries(personal.map((x,i)=>[x[0],i+7]));
for(const [label,r] of Object.entries(P)){
 let rule={type:'textLength',operator:'lessThanOrEqual',formula1:250};
 if(/Birth date|expiry|Reference date/.test(label)){p.getRange(`B${r}`).setNumberFormat('yyyy-mm-dd');rule={type:'date',operator:'between',formula1:'DATE(1900,1,1)',formula2:'DATE(2100,12,31)'};}
 else if(['Citizenship *','Home country *','Residence country *','ID issuing country *'].includes(label))rule={type:'list',formula1:addList(countries)};
 else if(label==='Applicant category *')rule={type:'list',formula1:addList(['EU candidate','Non-EU equivalent candidate','Non-EU candidate residing abroad','Other','Unknown'])};
 else if(label==='ID document type *')rule={type:'list',formula1:addList(['Passport','Identity card','Other'])};
 else if(/applicable|confirmed/.test(label))rule={type:'list',formula1:addList(yes)};
 validation(p.name,`B${r}`,rule,personal[r-7][2]);
}
p.getRange('A38:D40').merge();p.getRange('A38:D40').format.wrapText=true;
const missingPersonal=personal.flatMap((x,i)=>x[0].includes('*')?[`IF(B${i+7}="",1,0)`]:[]).join('+');
formula(p,'A38',`=IF(${missingPersonal}>0,"Complete all required personal fields. ","")&IF(AND(COUNTIF(B7:B35,"Other")>0,B23=""),"Explain Other in Other category details. ","")&IF(AND(B30="Yes",OR(B32="",B33="")),"Add permit evidence filename and page. ","")&IF(B35<>"Yes","Confirm accuracy. ","")&IF(AND(ISNUMBER(B27),ISNUMBER(B34),B27<B34),"Identity document expires before reference date. ","")`);
p.getRange('A38:D40').conditionalFormats.addCustom('A38<>""',{fill:'#FCE3E3',font:{color:'#9C1C23'}});
meta.sheets[p.name].freeze=6;

heading(sheets.Degrees,'Degrees','Select a degree cluster using the Degree class legend. A listed official code determines the cluster; otherwise choose the closest description. * = required.','H');declare(sheets.Degrees);
const deg=[id('D'),text('Degree title','BSc Computer Science','Official title; add English translation if useful.',true),list('Qualification level',['Bachelor / first cycle','Master / second cycle','Doctorate','Other']),text('Institution','Example University','Official awarding institution.',true),country('Institution country'),text('Programme student ID','00012345','Your matriculation / programme number.'),date('Enrolment date','2021-09-01','Start of this programme.',true),list('Status',['Graduated','Graduating'],'Graduated','Graduating = not yet awarded.',true),date('Graduation date','2024-07-15','Required only for Graduated.'),date('Expected graduation','2027-07-15','Required only for Graduating.'),num('Formal duration (months)',36,'Official full-time programme length.',true),calc('Actual duration (months)',(r,m)=>`=IF(OR(G${r}="",H${r}<>"Graduated",I${r}="",I${r}<G${r}),"",ROUND((I${r}-G${r})/30.4375,1))`,'Completed only; elapsed days / 30.4375.'),list('Main admission degree?',yes,'Yes','Select Yes for exactly one degree.',true),list('Degree cluster',clusters(degreeRows),'Computing & information engineering','See Degree class legend: use the cluster matching your official code, or the closest description if unlabelled.',true),list('Credit system',['ECTS','Local credits','Other'],'ECTS','Do not invent conversions.',true),num('Hours per credit',25,'Documented student workload, not contact hours only.'),num('Programme credits',180,'Original credit system.',true),num('Credits acquired',180,'Passed / recognised credits only.',true),num('Programme ECTS',180,'Only if documented; blank if unavailable.'),num('ECTS acquired',180,'Only if documented; blank if unavailable.'),num('Final mark / GPA',105,'Awarded degree result; blank until awarded.'),num('Final scale minimum',0,'Lowest possible final result, including failing values.'),num('Final pass threshold',66,'Minimum final result needed to pass.'),num('Final scale maximum',110,'Highest possible numeric final result.'),list('Final scale direction',['Higher is better','Lower is better','Not numeric','Other']),list('Final honours', ['Yes','No','Not applicable'],'No'),num('Weighted exam average',27.5,'Use original exam scale; do not convert.'),num('Exam scale minimum',0),num('Exam pass threshold',18),num('Exam scale maximum',30),list('Exam scale direction',['Higher is better','Lower is better','Not numeric','Other']),num('Remaining exams',0,'Whole number; 0 if none.',true),...evidence,other];
checkedTable('Degrees','Degrees',deg,10,8,[
 (r,m,v)=>`IF(AND(${v('Status')}="Graduated",${v('Graduation date')}=""),"Add graduation date; ","")`,
 (r,m,v)=>`IF(AND(${v('Status')}="Graduating",${v('Expected graduation')}=""),"Add expected graduation; ","")`,
 (r,m,v)=>`IF(AND(${v('Graduation date')}<>"",${v('Graduation date')}<${v('Enrolment date')}),"Graduation before enrolment; ","")`,
 (r,m,v)=>`IF(${v('Credits acquired')}>${v('Programme credits')},"Acquired credits exceed total; ","")`,
 (r,m,v)=>`IF(AND(${v('Credit system')}<>"ECTS",${v('Hours per credit')}=""),"Add documented hours per credit or explain unavailable in Other details; ","")`,
 (r,m,v)=>`IF(AND(${v('Final mark / GPA')}<>"",OR(${v('Final scale minimum')}="",${v('Final pass threshold')}="",${v('Final scale maximum')}="")),"Complete final grading scale; ","")`,
 (r,m,v)=>`IF(AND(${v('Final mark / GPA')}<>"",OR(${v('Final mark / GPA')}<${v('Final scale minimum')},${v('Final mark / GPA')}>${v('Final scale maximum')})),"Final grade outside range; ","")`
]);
const degreeIDs={...list('Degree ID',[],'D001','Select an automatically generated degree ID.',true),ref:'DegreeIDs'};
const exam=[id('EX'),text('Exam title','Algorithms','One row per attempt; copy transcript exactly.',true),degreeIDs,list('Result',['Passed','Failed'],'Passed','Record every attempt shown on the transcript.',true),date('Result date','2024-06-20','Exam / attempt date, including failed attempts.',true),num('Original credits',6,'Credits for this course, in the degree credit system.',true),calc('Credit system',(r,m)=>`=IF(C${r}="","",IFERROR(VLOOKUP(C${r},Degrees!$A$10:$AI$17,15,FALSE),"Unknown degree"))`),calc('Hours per credit',(r,m)=>`=IF(C${r}="","",IFERROR(IF(VLOOKUP(C${r},Degrees!$A$10:$AI$17,16,FALSE)=0,"",VLOOKUP(C${r},Degrees!$A$10:$AI$17,16,FALSE)),""))`),num('Documented ECTS',6,'No self-invented conversion.'),list('Grading type',['Numeric','Pass/fail','Other'],'Numeric','Use Pass/fail if no numeric mark.',true),num('Mark',28,'Blank for pass/fail grading.'),num('Scale minimum',0),num('Pass threshold',18),num('Scale maximum',30),list('Scale direction',['Higher is better','Lower is better','Not numeric','Other']),list('Honours',['Yes','No','Not applicable'],'No'),list('Exam cluster',clusters(sectorRows),'Computer science','See Exam sector legend: use the cluster matching your official SSD, or the closest description if unlabelled.',true),...evidence,other];
heading(sheets.Exams,'Exams','Select a Degree ID from Degrees. Select an exam cluster using the Exam sector legend: match a stated SSD, otherwise choose the closest description.','H');declare(sheets.Exams);
checkedTable('Exams','Exams',exam,10,200,[
 (r,m,v)=>`IF(COUNTIF(Degrees!$A$10:$A$17,${v('Degree ID')})<>1,"Unknown degree ID; ","")`,
 (r,m,v)=>`IF(AND(${v('Grading type')}="Numeric",OR(${v('Mark')}="",${v('Scale minimum')}="",${v('Pass threshold')}="",${v('Scale maximum')}="")),"Complete numeric mark and scale; ","")`,
 (r,m,v)=>`IF(AND(${v('Mark')}<>"",OR(${v('Mark')}<${v('Scale minimum')},${v('Mark')}>${v('Scale maximum')})),"Mark outside range; ","")`
]);

const activityChecks=[
 (r,m,v)=>`IF(AND(${v('Ongoing?')}="No",${v('End date')}=""),"Add end date; ","")`,
 (r,m,v)=>`IF(AND(${v('Ongoing?')}="Yes",${v('End date')}<>""),"Ongoing: clear end date; ","")`,
 (r,m,v)=>`IF(AND(${v('End date')}<>"",${v('End date')}<${v('Start date')}),"End before start; ","")`,
 (r,m,v)=>`IF(${reference}="","Set reference date in Personal; ",IF(${v('Start date')}>${reference},"Starts after reference date; ",""))`
];
const period=[date('Start date','2024-01-01','Real date.',true),date('End date','2024-06-30','Blank only if ongoing.'),list('Ongoing?',yes,'No','Yes: leave End date blank.',true),calc('Duration (months)',(r,m)=>`=IF(OR(${m['Start date']}${r}="",${reference}="",${m['Ongoing?']}${r}="",AND(${m['Ongoing?']}${r}="No",${m['End date']}${r}="")),"",IF(MIN(IF(${m['Ongoing?']}${r}="Yes",${reference},${m['End date']}${r}),${reference})<${m['Start date']}${r},"",ROUND((MIN(IF(${m['Ongoing?']}${r}="Yes",${reference},${m['End date']}${r}),${reference})-${m['Start date']}${r}+1)/30.4375,1)))`,'Calendar duration through reference date; inclusive days / 30.4375. Not an eligibility score.')];
const relevance=text('Role / why relevant','Built a reporting tool used to plan production.','Max 300 characters: your contribution and relevance to computing, business or digital transformation.',true);
const link={...list('Related activity ID',[],'W001','Optional: select the same activity recorded elsewhere. Do not count twice.'),ref:'ActivityIDs'};
const work=[id('W'),text('Position','Junior data analyst','Short role title.',true),text('Employer','Example Ltd','Organisation name.',true),country('Country'),text('City','Bologna','Work location.',true),...period,list('Activity type',['Employment','Internship','Volunteering','Self-employed','Other'],'Employment','Select the relationship.',true),list('Paid?',yes,'Yes','Separate from activity type.',true),list('Schedule',['Full-time','Part-time','Variable','Other'],'Full-time','Specify actual hours below.',true),num('Hours per week',40,'Average weekly workload.',true),relevance,...evidence,other];
heading(sheets['Work experience'],'Work experience','One row per role. Include internships and volunteering; clearly identify their type. Pink rows are examples. * = required.','H');declare(sheets['Work experience']);checkedTable('Work experience','Work',work,10,30,activityChecks);
const languages=[id('L'),list('Language',['English','Italian','French','German','Spanish','Portuguese','Arabic','Chinese','Russian','Turkish','Persian','Other'],'English','Other: specify language in Other details.',true),list('Native?',yes,'No','Native is not a CEFR certificate level.',true),list('Certified CEFR',[...level,'Not stated'],'B2','Use Not assessed if no certification.',true),...['Speaking','Reading','Writing','Listening'].map(h=>list(h+' self-assessment',level,'B2','Your own CEFR assessment.',true)),list('Evidence route',['Language certificate','English-taught qualification','Transcript language exam','Native declaration','Self-assessment only','Other'],'Language certificate','Select the evidence you actually submit.',true),list('Certificate type',['IELTS','TOEFL','Cambridge','University language centre','Other','Not applicable'],'Cambridge'),text('Issuer','Cambridge English','Official certificate / evidence provider.'),text('Certificate ID','ABC00123','If present.'),date('Certificate date','2025-06-01','Award / test date, if applicable.'),date('Certificate expiry','2027-06-01','Blank if none is stated.'),num('Test score',175,'Optional original score.'),text('Score scale','80–230','Original test scale; not a conversion.'),...evidence.map(x=>({...x,req:false})),other];
heading(sheets['Language skills'],'Language skills','One row per language / certificate. Self-assessment does not prove a certified level. Evidence is required for a claimed certificate.','H');declare(sheets['Language skills']);checkedTable('Language skills','Languages',languages,10,20,[
 (r,m,v)=>`IF(AND(${v('Evidence route')}<>"Self-assessment only",${v('Evidence route')}<>"Native declaration",OR(${v('Evidence filename')}="",${v('Page / section')}="")),"Add evidence reference; ","")`,
 (r,m,v)=>`IF(AND(${v('Evidence route')}="Language certificate",OR(${v('Issuer')}="",${v('Certificate date')}="")),"Add certificate issuer and date; ","")`
]);
const abroad=[id('AB'),list('Experience type',['Study exchange','Internship','Work','Research','Volunteering','Other'],'Study exchange','Physical stay outside your home country.',true),text('Host organisation','Example University','Employer / university / host.',true),country('Host country'),text('Host city','Berlin','Physical location.',true),list('Participation',['In person','Hybrid','Remote','Other'],'In person','Remote activity is flagged for committee review.',true),...period,link,relevance,...evidence,other];
heading(sheets.Internationalization,'Internationalization','Home country = country where you spent most of your life (Personal). Exclude stays in that country. Link an existing work/training activity instead of claiming it twice.','H');declare(sheets.Internationalization);checkedTable('Internationalization','Abroad',abroad,10,30,[...activityChecks,
 (r,m,v)=>`IF('Personal & contacts'!$B$13="","Set home country in Personal; ",IF(${v('Host country')}='Personal & contacts'!$B$13,"Exclude home-country experience; ",""))`,
 (r,m,v)=>`IF(${v('Participation')}="Remote","Remote: not a physical stay abroad; ","")`
]);
const training=[id('T'),text('Activity title','Data analytics summer school','Short title.',true),list('Activity type',['Internal internship','External internship','School project','Summer school','Postgraduate course','Professional course','Workshop','Language course','Other'],'Summer school','Select closest category.',true),text('Institution','Example Institute','Provider / host.',true),country('Country'),...period,num('Total hours',48,'Actual documented learning / work hours.',true),list('Completion',['Completed','Ongoing','Not completed','Other'],'Completed','Status at reference date.',true),list('Part of main degree?',yes,'No','Yes if already part of the academic curriculum.',true),link,relevance,...evidence,other];
heading(sheets.Training,'Training and additional education','Include training beyond your main academic programme. Identify curricular activities and link any internship already recorded under Work.','H');declare(sheets.Training);checkedTable('Training','Training',training,10,30,activityChecks);
const others=[id('O'),text('Activity title','Open-source inventory project','Short name.',true),list('Activity type',['Independent project','Open-source contribution','Startup / entrepreneurship','Competition / award','Volunteering','Research activity','Other'],'Independent project','Only activities that do not fit another sheet.',true),text('Organisation / context','Independent','Organisation or Independent.',true),country('Country'),...period,num('Total hours',120,'Documented estimate; explain basis in relevance.'),link,relevance,...evidence,other];
heading(sheets['Other activities'],'Other relevant activities','Use for activities that do not fit the other sheets. Explain your contribution and its relevance. Do not repeat an existing activity.','H');declare(sheets['Other activities']);checkedTable('Other activities','OtherActivities',others,10,30,activityChecks);

// Lists are visible and read-only so applicants can inspect all allowed values.
meta.sheets.Lists={freeze:0};
for(const [key,l] of Object.entries(meta.lists)){const vals=JSON.parse(key);const col=Object.values(meta.lists).indexOf(l);sheets.Lists.getRange(`${C(col)}1:${C(col)}${vals.length}`).values=vals.map(v=>[v]);sheets.Lists.getRange(`${C(col)}:${C(col)}`).format.columnWidth=34;}
meta.sheets.Lists.freeze=0;

const ins=sheets.Instructions;heading(ins,'Applicant evidence workbook','Complete this workbook and submit it together with the documents you reference. No macros are required.','H');
ins.getRange('A:A').format.columnWidth=6;ins.getRange('B:B').format.columnWidth=32;ins.getRange('C:H').format.columnWidth=15;
const steps=[
 ['1','Start with Personal','Enter your legal details, home country and the application deadline as Reference date.'],
 ['2','Edit yellow cells only','Pink = example; blue = guidance; grey = formula. * = required. Rows are compact: increase a row height to read longer entries.'],
 ['3','Choose the section status','On each category sheet choose Provided or None. A blank table is not a declaration of no experience.'],
 ['4','Use the empty rows','Examples are locked above each table. Start in row 10 on every input table. Leave unused rows completely empty.'],
 ['5','Use dropdowns and real dates','Dates display as yyyy-mm-dd. If ISO typing remains text in your Excel locale, enter a local date and check its ISO display.'],
 ['6','Use Other honestly','Choose Other if needed and complete Other details. Unknown means you do not know; Not applicable means it does not apply.'],
 ['7','Attach evidence','Type the exact submitted filename and page / section for each claim. Submit the actual files separately; a filename alone proves nothing.'],
 ['8','Link your records','IDs appear automatically when you enter a title. Select a Degree ID for each exam and Related activity ID for repeated contexts.'],
 ['9','Check carried-forward settings','Starting a row copies selected settings from the previous row. Yellow defaults are editable and stay linked until overwritten. Check each value. New degree: enter exam grading settings again.'],
 ['10','Check before submitting','Review every red Check cell, all required fields, section statuses and evidence attachments. No red warning does not certify completeness.'],
 ['11','Save and submit','Save as .xlsx with your name or application number. Keep the provided row order; filtering is allowed, sorting and row insertion are blocked.'],
 ['12','If you run out of rows','Ask the admissions office for an expanded copy. Capacity: 8 degrees, 200 exams, 30 activities per sheet, 20 language records.']
];
ins.getRange('A6:H6').merge();ins.getRange('A6').values=[['Yellow = input. Blue = guidance. Pink = example. Grey = formula. Red = review.']];ins.getRange('A6:H6').format={fill:'#FFF2A6',font:{bold:true},rowHeight:30};
steps.forEach((x,i)=>{let r=8+i*2;ins.getRange(`A${r}:A${r+1}`).merge();ins.getRange(`B${r}:B${r+1}`).merge();ins.getRange(`C${r}:H${r+1}`).merge();ins.getRange(`A${r}`).values=[[x[0]]];ins.getRange(`B${r}`).values=[[x[1]]];ins.getRange(`C${r}`).values=[[x[2]]];ins.getRange(`A${r}:H${r+1}`).format={font:{name:'Arial',size:11,color:'#172033'},wrapText:true,verticalAlignment:'center',rowHeight:23};ins.getRange(`B${r}`).format.font.bold=true;});
ins.getRange('A34:H35').merge();ins.getRange('A34').values=[['Home country: the country where you have spent most of your life. If tied, explain in Personal → Other category details. Durations are calendar months (days / 30.4375), not awarded points. Overlapping periods and self-classifications require committee review.']];ins.getRange('A34:H35').format={wrapText:true,rowHeight:27,font:{size:10,color:'#475569'}};
ins.getRange('A37:D37').values=[['Section','Declared status','Rows entered','Rows with warnings']];ins.getRange('A37:D37').format={fill:'#23374D',font:{color:'#FFFFFF',bold:true},rowHeight:30};
const categorySheets=['Degrees','Exams','Work experience','Language skills','Internationalization','Training','Other activities'];
categorySheets.forEach((n,i)=>{const r=38+i;const ts=meta.tables.filter(t=>t.sheet===n);ins.getRange(`A${r}:A${r}`).values=[[n]];formula(ins,`B${r}`,`='${n}'!B5`);formula(ins,`C${r}`,'='+ts.map(t=>`COUNTA('${n}'!B${t.start}:B${t.end})`).join('+'));formula(ins,`D${r}`,'='+ts.map(t=>`COUNTIF('${n}'!${C(t.cols.length-1)}${t.start}:${C(t.cols.length-1)}${t.end},"?*")`).join('+'));});
ins.getRange('A38:D44').format={font:{name:'Arial',size:11},rowHeight:30};ins.getRange('A:A').format.columnWidth=26;
ins.getRange('E37:H37').merge();ins.getRange('E37').values=[['Section check']];
categorySheets.forEach((n,i)=>{const r=38+i;ins.getRange(`E${r}:H${r}`).merge();formula(ins,`E${r}`,`=IF(B${r}="Not started","Choose Provided or None",IF(AND(B${r}="None",C${r}>0),"None conflicts with entered rows",IF(AND(B${r}="Provided",C${r}=0),"Add rows or choose None",IF(D${r}>0,"Review red Check cells","Review evidence before submission"))))`);});
ins.getRange('E38:H44').format={font:{name:'Arial',size:11},wrapText:true,rowHeight:30};
ins.getRange('A37:D44').format.borders={preset:'all',style:'thin',color:'#AAB6C4'};
ins.getRange('E37:H44').format.borders={preset:'outside',style:'thin',color:'#AAB6C4'};
ins.getRange('A46:H47').merge();formula(ins,'A46','=IF(COUNTIF(Degrees!M10:M17,"Yes")<>1,"Education: select exactly one main admission degree.","Main admission degree selected. Check its transcript and grading scales.")');ins.getRange('A46:H47').format.wrapText=true;

meta.definedNames=[...Object.values(meta.lists),{name:'DegreeIDs',ref:"'Degrees'!$A$10:$A$17"}];
const listCol=Object.keys(meta.lists).length+1;
let activityRow=1;
for(const n of ['Work experience','Training','Other activities'])for(let r=10;r<=39;r++)formula(sheets.Lists,`${C(listCol)}${activityRow++}`,`=IF('${n}'!A${r}="","",'${n}'!A${r})`);
meta.definedNames.push({name:'ActivityIDs',ref:`'Lists'!$${C(listCol)}$1:$${C(listCol)}$90`});
for(const item of meta.definedNames)wb.names.add(item.name,'='+item.ref);
meta.sheets.Lists.freeze=0;
wb.recalculate();
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:15},maxChars:1500})).ndjson);
// Exercise representative formula dependencies, then restore every sample input.
const assert=(ok,msg)=>{if(!ok)throw new Error(msg);};
const edu=sheets.Degrees,ex=sheets.Exams;
const dm=meta.tables.find(t=>t.sheet==='Degrees').map,em=meta.tables.find(t=>t.sheet==='Exams').map;
edu.getRange('B10').values=[['Test degree']];edu.getRange('G10').values=[[new Date('2021-09-01T00:00:00Z')]];edu.getRange('H10').values=[['Graduated']];edu.getRange('I10').values=[[new Date('2024-09-01T00:00:00Z')]];
edu.getRange(dm['Degree cluster']+'10').values=[['Computing & information engineering']];
edu.getRange(dm['Credit system']+'10').values=[['Local credits']];edu.getRange(dm['Hours per credit']+'10').values=[[25]];
ex.getRange('B10').values=[['Test exam']];ex.getRange('C10').values=[['D001']];ex.getRange(em['Exam cluster']+'10').values=[['Computer science']];
wb.recalculate();
assert(edu.getRange('A10').values[0][0]==='D001','Generated degree ID');
assert(ex.getRange('A10').values[0][0]==='EX001','Generated exam ID');
assert(Math.abs(Number(edu.getRange('L10').values[0][0])-36)<0.2,'Degree duration');
assert(ex.getRange('G10').values[0][0]==='Local credits','Cross-sheet degree link');
assert(ex.getRange('H10').values[0][0]===25,'Cross-sheet credit hours');
assert(!String(ex.getRange(em.Check+'10').values[0][0]).includes('Unknown degree'),'Known degree accepted');
assert(!String(edu.getRange(dm.Check+'10').values[0][0]).includes('Invalid choice: Degree cluster'),'Cluster accepted');
// Input defaults remain blank until a row starts, preserve zero, and allow overrides.
edu.getRange(dm['Final scale minimum']+'10').values=[[0]];
edu.getRange(dm['Final scale maximum']+'10').values=[[110]];
assert(edu.getRange(dm['Final scale maximum']+'11').values[0][0]==='','Unused row default stays blank');
edu.getRange('B11').values=[['Second degree']];edu.getRange('B12').values=[['Third degree']];
assert(edu.getRange(dm['Final scale minimum']+'11').values[0][0]===0,'Zero grading bound is inherited');
assert(edu.getRange(dm['Final scale maximum']+'12').values[0][0]===110,'Defaults propagate through started rows');
const defaultFormula=edu.getRange(dm['Final scale maximum']+'11').formulas;
edu.getRange(dm['Final scale maximum']+'11').values=[[100]];
assert(edu.getRange(dm['Final scale maximum']+'12').values[0][0]===100,'Applicant override propagates');
edu.getRange(dm['Final scale maximum']+'11').formulas=defaultFormula;
ex.getRange(em['Scale minimum']+'10').values=[[0]];ex.getRange(em['Scale maximum']+'10').values=[[30]];
ex.getRange('B11').values=[['Second exam']];
assert(ex.getRange('C11').values[0][0]==='D001','Next exam inherits degree');
assert(ex.getRange(em['Scale maximum']+'11').values[0][0]===30,'Next exam inherits grading range');
const degreeDefault=ex.getRange('C11').formulas;
ex.getRange('C11').values=[['D002']];
assert(ex.getRange(em['Scale maximum']+'11').values[0][0]==='','Degree change clears inherited exam range');
ex.getRange('C11').formulas=degreeDefault;
for(const c of ['B11','B12',dm['Final scale minimum']+'10',dm['Final scale maximum']+'10'])edu.getRange(c).clear({applyTo:'contents'});
for(const c of ['B11',em['Scale minimum']+'10',em['Scale maximum']+'10'])ex.getRange(c).clear({applyTo:'contents'});
for(const cluster of clusters(degreeRows))assert(!['Unknown','Not applicable'].includes(cluster),'Degree options contain clusters only');
for(const cluster of clusters(sectorRows))assert(!['Unknown','Not applicable'].includes(cluster),'Exam options contain clusters only');
for(const c of ['B10','G10','H10','I10',dm['Degree cluster']+'10',dm['Credit system']+'10',dm['Hours per credit']+'10'])edu.getRange(c).clear({applyTo:'contents'});
for(const c of ['B10','C10',em['Exam cluster']+'10'])ex.getRange(c).clear({applyTo:'contents'});
const workSheet=sheets['Work experience'];
p.getRange('B34').values=[[new Date('2026-11-20T00:00:00Z')]];
workSheet.getRange('B10').values=[['Test role']];
workSheet.getRange('F10').values=[[new Date('2026-01-01T00:00:00Z')]];
workSheet.getRange('G10').values=[[new Date('2026-03-31T00:00:00Z')]];
workSheet.getRange('H10').values=[['No']];
assert(Math.abs(Number(workSheet.getRange('I10').values[0][0])-3)<0.1,'Completed activity duration');
workSheet.getRange('H10').values=[['Yes']];workSheet.getRange('G10').clear({applyTo:'contents'});
assert(Number(workSheet.getRange('I10').values[0][0])>10,'Ongoing duration uses reference date');
for(const c of ['B10','F10','G10','H10'])workSheet.getRange(c).clear({applyTo:'contents'});
p.getRange('B34').clear({applyTo:'contents'});
wb.recalculate();
await fs.writeFile(out+'metadata.json',JSON.stringify(meta,null,2));
await (await SpreadsheetFile.exportXlsx(wb)).save(out+'applicant-template.xlsx');
for(const name of names){
 if(name==='Lists')continue;
 const range=name==='Instructions'?'A1:H31':name==='Personal & contacts'?'A1:D16':name==='Degree class legend'?'A1:E10':name==='Exam sector legend'?'A1:E10':name==='Degrees'?'A1:H12':'A1:H11';
 try{const pic=await wb.render({sheetName:name,range,scale:1,format:'png'});await fs.writeFile(out+name.replaceAll(/[^a-zA-Z]/g,'_')+'.png',new Uint8Array(await pic.arrayBuffer()));}catch(e){console.log('Render '+name+': '+e.message);}
}
for(const [name,range] of [['Degrees','M7:S11'],['Exams','N7:U11']]){
 const pic=await wb.render({sheetName:name,range,scale:1,format:'png'});
 await fs.writeFile(out+name+'_clusters.png',new Uint8Array(await pic.arrayBuffer()));
}
console.log('Created workbook and previews.');
