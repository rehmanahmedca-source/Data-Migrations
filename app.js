const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
let legacy=[], schema=null;
function renderFiles(){
  $('#legacyFiles').innerHTML=legacy.map(f=>`<div class="file-chip">⌁ ${f.name}<span>${(f.size/1024).toFixed(1)} KB</span></div>`).join('');
  $('#schemaFiles').innerHTML=schema?`<div class="file-chip">⌁ ${schema.name}<span>${(schema.size/1024).toFixed(1)} KB</span></div>`:'';
  $('#discoverBtn').disabled=!(legacy.length&&schema);
}
$('#legacyInput').onchange=e=>{legacy=[...e.target.files];renderFiles()}; $('#schemaInput').onchange=e=>{schema=e.target.files[0];renderFiles()};
function drop(id,input){let el=$(id);['dragenter','dragover'].forEach(x=>el.addEventListener(x,e=>{e.preventDefault();el.style.borderColor='#3e765c'}));el.addEventListener('dragleave',()=>el.style.borderColor='');el.addEventListener('drop',e=>{e.preventDefault();el.style.borderColor='';let fs=[...e.dataTransfer.files];if(input==='legacy')legacy=fs;else schema=fs[0];renderFiles()})}drop('#legacyDrop','legacy');drop('#schemaDrop','schema');
function show(id){$$('.content').forEach(x=>x.classList.add('hidden'));$(id).classList.remove('hidden');window.scrollTo(0,0)}
$('#discoverBtn').onclick=()=>{$('#sourceName').textContent=legacy[0].name;$('#targetName').textContent=schema.name;show('#review')};$('#backBtn').onclick=()=>show('#workspace');$('#exportBtn').onclick=()=>{$('#exportCount').textContent=$('#selectedCount').textContent;show('#export')};$('#reviewBack').onclick=()=>show('#review');
$$('.record:not(.head)').forEach(row=>row.onclick=e=>{if(e.target.classList.contains('check'))e.stopPropagation();let c=row.querySelector('.check');c.classList.toggle('checked');updateCount()});
function updateCount(){let n=$$('.record:not(.head) .check.checked').reduce((a,c)=>a+ +c.closest('.record').dataset.count,0);$('#selectedCount').textContent=n.toLocaleString();$('#exportCount').textContent=n.toLocaleString()}
$('#selectAll').onclick=()=>{$$('.record:not(.head) .check').forEach(c=>c.classList.add('checked'));updateCount()};
$$('.format').forEach(b=>b.onclick=()=>{$$('.format').forEach(x=>x.classList.remove('selected'));b.classList.add('selected')});$('#outputName').oninput=e=>$('#folderPreview').textContent=e.target.value||'untitled-migration';$('#finishBtn').onclick=()=>{$('#toast').classList.add('show');setTimeout(()=>$('#toast').classList.remove('show'),350)};
