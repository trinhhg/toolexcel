const OFF=4,$=id=>document.getElementById(id),T=$('grid'),CJK=/[\u4e00-\u9fff]/;
const NAMES={q:'Đề',a:'Đáp án',inp:'Ô nhập',res:'Kết quả'};
const L=n=>{let s='';for(n++;n;n=((n-1)/26)|0)s=String.fromCharCode(65+(n-1)%26)+s;return s};
const S={ncol:8,roles:{q:3,a:1},F:null,sub:false,lock:true,sel:null,orig:null,
 rows:[['STT','Tiếng Trung','Pinyin','Tiếng Việt'],['1','你好','nǐ hǎo','xin chào'],['2','谢谢','xièxie','cảm ơn'],['3','学校','xuéxiào','trường học'],['4','老师','lǎoshī','giáo viên'],['5','学生','xuéshēng','học sinh']]};
let drag=false;
const cell=(r,c)=>T.rows[r+1]?.cells[c+1];
const inR=k=>S.F&&k>=1&&k>=S.F.r1&&k<=S.F.r2;
const tag=c=>Object.keys(S.roles).filter(k=>S.roles[k]===c).map(k=>NAMES[k]).join('/');

/* ---------- chấm điểm (mirror công thức Excel) ---------- */
function ok(i,a){const F=S.F,f=s=>{s=s.trim();return F.ci?s.toLowerCase():s};return F.multi?a.split('|').map(f).includes(f(i)):f(i)===f(a)}
function st(r){const F=S.F,i=(r[F.inp]||'').trim();return !i||(F.mode=='exam'&&!S.sub)?'':ok(i,r[F.a]||'')?'ĐÚNG':'SAI'}
function stats(){const F=S.F;let d=0,s=0,t=0,dn=0;for(let k=F.r1;k<=F.r2;k++){const r=S.rows[k],x=st(r);if((r[F.a]||'').trim())t++;if((r[F.inp]||'').trim())dn++;if(x=='ĐÚNG')d++;else if(x=='SAI')s++}return{d,s,t,dn}}
function scoreArr(){const E=[['KẾT QUẢ'],[],[],[]];if(!S.F)return E;const{d,s,t,dn}=stats(),p=Math.min(10,t?Math.round(dn/t*10):0);
 E[1]=['Đúng',d,'Sai',s,'Chưa làm',t-d-s];E[2]=['Tổng',t,'Điểm /10',t?+(d/t*10).toFixed(2):0,'Tỷ lệ',(t?Math.round(d/t*100):0)+'%'];
 E[3]=['Nộp bài','','Đã làm',dn,'Tiến độ','█'.repeat(p)+'░'.repeat(10-p)];return E}
function fml(x){const F=S.F,I=L(F.inp)+x,A=L(F.a)+x,n=s=>{s=`TRIM(${s})`;return F.ci?`LOWER(${s})`:s};
 const c=F.multi?`ISNUMBER(FIND("|"&${n(I)}&"|","|"&SUBSTITUTE(SUBSTITUTE(${n(A)}," |","|"),"| ","|")&"|"))`:F.ci?`${n(I)}=${n(A)}`:`EXACT(${n(I)},${n(A)})`;
 return `IF(${F.mode=='exam'?`OR(${I}="",$B$4<>"Đã nộp")`:`${I}=""`},"",IF(${c},"ĐÚNG","SAI"))`}

/* ---------- vẽ bảng ---------- */
function draw(){
 let h='<thead><tr><th class="corner"></th>';
 for(let c=0;c<S.ncol;c++)h+=`<th data-c="${c}">${L(c)}<i>${tag(c)}</i></th>`;
 h+='</tr></thead><tbody>';
 for(let r=0;r<S.rows.length+OFF;r++){const k=r-OFF;h+=`<tr><th class="rh" data-r="${r}">${r+1}</th>`;
  for(let c=0;c<S.ncol;c++){const ed=r>=OFF&&!(S.F&&c===S.F.res&&inR(k));
   h+=`<td data-r="${r}" data-c="${c}" data-z="${r<OFF?'sc':r==OFF?'hd':''}"${ed?' contenteditable="true"':''}></td>`}
  h+='</tr>'}
 T.innerHTML=h+'</tbody>';
 if(S.F?.mode=='exam'){const td=cell(3,1);td.innerHTML='<select id="sub"><option>Chưa nộp</option><option>Đã nộp</option></select>';
  $('sub').value=S.sub?'Đã nộp':'Chưa nộp';$('sub').onchange=e=>{S.sub=e.target.value=='Đã nộp';paint()}}
 paint();hl();
}
function paint(){
 const sc=scoreArr();
 for(let r=0;r<S.rows.length+OFF;r++){const k=r-OFF,row=k>=0?S.rows[k]:null,x=inR(k)?st(row):null;
  for(let c=0;c<S.ncol;c++){if(r==3&&c==1&&S.F?.mode=='exam')continue;
   const td=cell(r,c);let v=r<OFF?(sc[r][c]??''):(x!==null&&c===S.F.res?x:(row[c]??''));
   if(td.textContent!==String(v))td.textContent=v;
   td.dataset.s=x&&(c===S.F.res||c===S.F.inp)?(x=='ĐÚNG'?'ok':'bad'):''}}
 fx();
}
function hl(){document.querySelectorAll('td.sel').forEach(t=>t.classList.remove('sel'));const s=S.sel;if(!s)return;
 const[a,b]=[Math.min(s.r1,s.r2),Math.max(s.r1,s.r2)],[c,d]=[Math.min(s.c1,s.c2),Math.max(s.c1,s.c2)];
 for(let r=a;r<=b;r++)for(let k=c;k<=d;k++)cell(r,k)?.classList.add('sel');
 if(a>=5){$('r1').value=a+1;$('r2').value=b+1}fx()}
function fx(){const s=S.sel;if(!s)return;const r=s.r1,c=s.c1,k=r-OFF;$('nm').textContent=L(c)+(r+1);
 $('fv').textContent=S.F&&inR(k)&&c===S.F.res?'='+fml(r+1):(r<OFF?'':S.rows[k]?.[c]||'')}
const syncR=()=>{$('r1').value=6;$('r2').value=S.rows.length+4};

/* ---------- chọn ô / cột / hàng ---------- */
T.onmousedown=e=>{const t=e.target.closest('td'),h=e.target.closest('th');
 if(t){S.sel={r1:+t.dataset.r,c1:+t.dataset.c,r2:+t.dataset.r,c2:+t.dataset.c};drag=true}
 else if(h?.dataset.c!=null)S.sel={r1:0,c1:+h.dataset.c,r2:S.rows.length+OFF-1,c2:+h.dataset.c};
 else if(h?.dataset.r!=null)S.sel={r1:+h.dataset.r,c1:0,r2:+h.dataset.r,c2:S.ncol-1};
 hl()};
T.onmouseover=e=>{if(!drag)return;const t=e.target.closest('td');if(t){S.sel.r2=+t.dataset.r;S.sel.c2=+t.dataset.c;hl()}};
document.onmouseup=()=>drag=false;
T.addEventListener('focusin',e=>{const t=e.target.closest('td');if(t&&!drag){S.sel={r1:+t.dataset.r,c1:+t.dataset.c,r2:+t.dataset.r,c2:+t.dataset.c};hl()}});
T.addEventListener('input',e=>{const t=e.target.closest('td');if(!t)return;const k=+t.dataset.r-OFF;
 if(k>=0){S.rows[k][+t.dataset.c]=t.textContent;paint()}});
T.addEventListener('keydown',e=>{if(e.key=='Enter'){e.preventDefault();const t=e.target.closest('td');cell(+t.dataset.r+1,+t.dataset.c)?.focus()}});

/* ---------- quét dữ liệu dán ---------- */
function parse(t){return t.split(/\r?\n/).map(l=>l.trim()).filter(Boolean).map(l=>{
 let p=l.includes('\t')?l.split('\t'):l.includes('|')?l.split('|'):l.includes(';')?l.split(';'):/\s{2,}/.test(l)?l.split(/\s{2,}/):null;
 if(!p){const m=l.match(/^([\u4e00-\u9fff，。？！、…]+)\s+(.+)$/);
  if(m){const w=m[2].split(/\s+/),n=m[1].length*2;let py=[],c=0;while(w.length>1&&c<n){const x=w.shift();py.push(x);c+=x.length}p=[m[1],py.join(' '),w.join(' ')]}
  else p=l.split(/\s+/)}
 return p.map(x=>x.trim())})}
$('bPaste').onclick=()=>$('dlg').hidden=false;$('close').onclick=()=>$('dlg').hidden=true;
$('go').onclick=()=>{const P=parse($('ta').value);if(!P.length)return;
 const n=Math.max(...P.map(r=>r.length)),cols=Array.from({length:n},(_,i)=>P.map(r=>r[i]||'')),
  cj=cols.map(c=>c.filter(x=>CJK.test(x)).length>c.length/2),rest=cols.map((_,i)=>i).filter(i=>!cj[i]),
  hd=cols.map((_,i)=>cj[i]?'Tiếng Trung':rest.length>1&&i==rest[0]?'Pinyin':'Nghĩa');
 S.rows=[['STT',...hd],...P.map((r,k)=>[String(k+1),...cols.map((_,i)=>r[i]||'')])];
 S.roles={a:Math.max(1,cj.indexOf(true)+1),q:n};S.ncol=Math.max(8,n+4);S.F=null;S.orig=null;S.sub=false;
 $('dlg').hidden=true;syncR();draw()};

/* ---------- vai trò cột + áp công thức ---------- */
document.querySelectorAll('[data-role]').forEach(b=>b.onclick=()=>{if(!S.sel)return alert('Bấm vào tiêu đề cột (A, B, C…) trước');
 S.roles[b.dataset.role]=S.sel.c1;S.F=null;draw()});
$('bSwap').onclick=()=>{[S.roles.q,S.roles.a]=[S.roles.a,S.roles.q];S.F=null;draw()};
$('bAddCol').onclick=()=>{S.ncol++;draw()};
const free=()=>{let c=0;S.rows[0].forEach((x,i)=>{if(x)c=i+1});if(c>=S.ncol)S.ncol=c+1;return c};
$('bApply').onclick=()=>{const R=S.roles;if(R.q==null||R.a==null)return alert('Chọn cột Đề và Đáp án trước');
 if(R.inp==null){R.inp=free();S.rows[0][R.inp]='Nhập đáp án'}
 if(R.res==null){R.res=free();S.rows[0][R.res]='Kết quả'}
 const n=S.rows.length;
 S.F={q:R.q,a:R.a,inp:R.inp,res:R.res,r1:Math.max(1,(+$('r1').value||6)-5),r2:Math.min(n-1,(+$('r2').value||n+4)-5),
  mode:$('mode').value,ci:$('ci').checked,multi:$('multi').checked,hide:$('hide').checked};
 draw()};
['mode','ci','multi','hide'].forEach(id=>$(id).onchange=()=>{if(!S.F)return;
 Object.assign(S.F,{mode:$('mode').value,ci:$('ci').checked,multi:$('multi').checked,hide:$('hide').checked});draw()});

/* ---------- trộn nguyên hàng ---------- */
const shuf=a=>{for(let i=a.length-1;i>0;i--){const j=(Math.random()*(i+1))|0;[a[i],a[j]]=[a[j],a[i]]}return a};
const renum=rows=>{const c=rows[0].indexOf('STT');if(c>=0)for(let k=1;k<rows.length;k++)rows[k][c]=String(k)};
const snap=()=>{S.orig??=S.rows.map(r=>[...r])};
$('bShuf').onclick=()=>{snap();S.rows=[S.rows[0],...shuf(S.rows.slice(1))];renum(S.rows);draw()};
$('bRest').onclick=()=>{if(S.orig){S.rows=S.orig.map(r=>[...r]);draw()}};
$('bPick').onclick=()=>{const n=+$('nPick').value;if(n<1)return;snap();
 S.rows=[S.rows[0],...shuf(S.rows.slice(1)).slice(0,n)];renum(S.rows);
 if(S.F){S.F.r1=1;S.F.r2=S.rows.length-1}syncR();draw()};

/* ---------- xuất XLSX (công thức + khoá + tô màu) ---------- */
$('bLock').onclick=e=>{S.lock=!S.lock;e.target.textContent='🔒 Khoá bảng: '+(S.lock?'BẬT':'TẮT');e.target.classList.toggle('on',S.lock)};
$('bExp').onclick=async()=>{
 if(!S.F)return alert('Bấm "Áp công thức" trước');
 const F=S.F,wb=new ExcelJS.Workbook(),nv=Math.max(1,+$('nv').value||1),h=S.rows[0],stt=h.indexOf('STT');let pr=1;
 for(let v=0;v<nv;v++){
  const d=S.rows.slice(1).map(r=>[...r]);if(v)shuf(d);const rows=[h,...d];if(v)renum(rows);
  const R1=nv>1?1:F.r1,R2=nv>1?rows.length-1:F.r2,e1=R1+5,e2=R2+5,rg=c=>`${L(c)}${e1}:${L(c)}${e2}`;
  const ws=wb.addWorksheet('Đề '+(v+1),{views:[{state:'frozen',ySplit:5}]}),f=(a,s)=>ws.getCell(a).value={formula:s};
  for(let c=0;c<S.ncol;c++)ws.getColumn(c+1).width=c==stt?7:20;
  rows.forEach((r,k)=>r.forEach((x,c)=>{if(x!==''&&x!=null)ws.getCell(k+5,c+1).value=c==stt&&k?+x:x}));
  for(let k=R1;k<=R2;k++){const x=k+5;f(L(F.res)+x,fml(x));ws.getCell(x,F.inp+1).protection={locked:false}}
  [['A1','KẾT QUẢ'],['A2','Đúng'],['C2','Sai'],['E2','Chưa làm'],['A3','Tổng'],['C3','Điểm /10'],['E3','Tỷ lệ'],['A4','Nộp bài'],['C4','Đã làm'],['E4','Tiến độ']]
   .forEach(([a,t])=>{ws.getCell(a).value=t});
  f('B2',`COUNTIF(${rg(F.res)},"ĐÚNG")`);f('D2',`COUNTIF(${rg(F.res)},"SAI")`);f('F2','B3-B2-D2');
  f('B3',`COUNTA(${rg(F.a)})`);f('D3','IF(B3=0,0,ROUND(B2/B3*10,2))');f('F3','IF(B3=0,0,B2/B3)');ws.getCell('F3').numFmt='0%';
  f('D4',`COUNTA(${rg(F.inp)})`);f('F4','REPT("█",MIN(10,ROUND(D4/MAX(B3,1)*10,0)))&REPT("░",10-MIN(10,ROUND(D4/MAX(B3,1)*10,0)))');
  if(F.mode=='exam'){const b=ws.getCell('B4');b.value='Chưa nộp';b.protection={locked:false};b.dataValidation={type:'list',allowBlank:false,formulae:['"Chưa nộp,Đã nộp"']}}
  else ws.getCell('B4').value='—';
  for(let r=1;r<=4;r++)for(let c=1;c<=6;c++){const x=ws.getCell(r,c);x.font={bold:true};x.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF7F3E3'}}}
  for(let c=1;c<=S.ncol;c++){const x=ws.getCell(5,c);x.font={bold:true};x.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD9E8DD'}}}
  const fill=a=>({fill:{type:'pattern',pattern:'solid',bgColor:{argb:a}}});
  for(const c of [F.inp,F.res])ws.addConditionalFormatting({ref:rg(c),rules:[
   {type:'expression',priority:pr++,formulae:[`$${L(F.res)}${e1}="ĐÚNG"`],style:fill('FFC6EFCE')},
   {type:'expression',priority:pr++,formulae:[`$${L(F.res)}${e1}="SAI"`],style:fill('FFFFC7CE')}]});
  if(F.hide)ws.getColumn(F.a+1).hidden=true;
  if(S.lock)await ws.protect($('pw').value,{selectLockedCells:true,selectUnlockedCells:true});
 }
 const buf=await wb.xlsx.writeBuffer(),a=document.createElement('a');
 a.href=URL.createObjectURL(new Blob([buf]));a.download='De_tu_vung.xlsx';a.click();
};

syncR();draw();
