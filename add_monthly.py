import sys, re
sys.stdout.reconfigure(encoding='utf-8')

with open('c:/Users/admin/Desktop/work/자동화/CNC분석/CNC_일별생산현황_v3.html', encoding='utf-8') as f:
    content = f.read()

# ── 1. 탭 버튼 ──
content = content.replace(
    "onclick=\"switchTab('freq')\">📊 발주 빈도</div>\n</div>",
    "onclick=\"switchTab('freq')\">📊 발주 빈도</div>\n  <div class=\"tab\" onclick=\"switchTab('monthly')\">📅 월별 비교</div>\n</div>"
)
print("1:", "월별 비교" in content)

# ── 2. switchTab 배열 ──
content = content.replace(
    "const tabs=['daily','worker','delivery','order','problem','freq'];",
    "const tabs=['daily','worker','delivery','order','problem','freq','monthly'];"
)
print("2:", ",'monthly']" in content)

# ── 3. renderCurrent ──
content = content.replace(
    "  else if(currentTab==='freq'){renderFreq();}\n}",
    "  else if(currentTab==='freq'){renderFreq();}\n  else if(currentTab==='monthly'){renderMonthly();}\n}"
)
print("3:", "renderMonthly" in content)

# ── 4. chart 변수 선언 ──
content = content.replace(
    "let comboChart=null,dimChart=null,leadChart=null;",
    "let comboChart=null,dimChart=null,leadChart=null,monthlyBarChart=null,monthlyTrendChart=null;"
)
print("4:", "monthlyBarChart=null" in content)

# ── 5. 월별 탭 HTML ──
html_block = """
<!-- ══ 탭: 월별 비교 ══ -->
<div id="tab-monthly" class="tab-content">
<div class="main">
  <div class="card">
    <div class="card-hdr">
      <h3>📅 월별 생산수량 비교 (호기별 누적)</h3>
      <span class="note">막대 위 숫자 = 월 총 생산량</span>
    </div>
    <div class="card-body"><canvas id="monthlyBarChart"></canvas></div>
    <div class="legend-row" id="monthlyLegend"></div>
    <div class="stats-row" id="monthlyStatsRow"></div>
  </div>
  <div class="card">
    <div class="card-hdr">
      <h3>📈 월별 추이 · 전월 대비 증감</h3>
      <span class="note">초록 = 전월 대비 증가 · 빨강 = 감소</span>
    </div>
    <div class="card-body"><canvas id="monthlyTrendChart"></canvas></div>
  </div>
  <div class="card">
    <div class="card-hdr">
      <h3>📋 월별 상세 현황</h3>
      <span class="note">전월 대비 증감율 · 호기별 생산량</span>
    </div>
    <div class="order-table-wrap"><div id="monthlyTableWrap"></div></div>
  </div>
</div>
</div>

"""
content = content.replace('<div class="tooltip-custom" id="tooltip"></div>', html_block + '<div class="tooltip-custom" id="tooltip"></div>')
print("5:", "tab-monthly" in content)

# ── 6. renderMonthly 함수 ──
fn = """
// ═══════════════ 월별 비교 ═══════════════
function renderMonthly(){
  const filtDates=filteredDates();
  const months=[...new Set(filtDates.map(d=>d.slice(0,7)))].sort();
  const machList=MACHINES.filter(m=>activeMachines.has(m));
  if(!months.length) return;

  // 월별 호기별 집계
  const moData={};
  months.forEach(mo=>{
    moData[mo]={};
    machList.forEach(m=>{
      moData[mo][m]=filtDates.filter(d=>d.startsWith(mo)).reduce((s,d)=>s+((D.bar[d]||{})[m]||0),0);
    });
  });
  const moTot=months.map(mo=>machList.reduce((s,m)=>s+(moData[mo][m]||0),0));
  const momPct=moTot.map((v,i)=>i===0?null:moTot[i-1]>0?((v-moTot[i-1])/moTot[i-1]*100).toFixed(1):null);

  // 누적 막대 차트
  const barDs=machList.map(m=>({
    label:m,type:'bar',
    data:months.map(mo=>moData[mo][m]||0),
    backgroundColor:MACH_COLORS[m]+'99',borderColor:MACH_COLORS[m],
    borderWidth:1,stack:'s'
  }));
  const topLbl={id:'moLbl',afterDatasetsDraw(chart){
    const {ctx,scales:{x,y}}=chart;
    months.forEach((mo,i)=>{
      const tot=moTot[i]; if(!tot) return;
      const xPos=x.getPixelForValue(mo),top=y.getPixelForValue(tot);
      ctx.save();ctx.fillStyle='#e0e8f0';ctx.font='bold 10px sans-serif';ctx.textAlign='center';
      ctx.fillText(tot.toLocaleString()+'개',xPos,top-6);ctx.restore();
    });
  }};
  if(monthlyBarChart) monthlyBarChart.destroy();
  monthlyBarChart=new Chart(document.getElementById('monthlyBarChart'),{
    type:'bar',data:{labels:months,datasets:barDs},
    options:{
      responsive:true,animation:false,layout:{padding:{top:22}},
      scales:{
        x:{stacked:true,ticks:{color:'#94a3b8',font:{size:10}},grid:{color:'#1e3a5f55'}},
        y:{stacked:true,ticks:{color:'#94a3b8',font:{size:9}},grid:{color:'#1e3a5f55'},
          title:{display:true,text:'수량',color:'#7dd3fc',font:{size:9}}}
      },
      plugins:{legend:{display:false},tooltip:{callbacks:{
        footer:its=>'합계: '+moTot[its[0].dataIndex].toLocaleString()+'개'
      }}}
    },plugins:[topLbl]
  });
  const ml=document.getElementById('monthlyLegend');ml.innerHTML='';
  machList.forEach(m=>{ml.innerHTML+=`<div class="li"><div class="lsq" style="background:${MACH_COLORS[m]}"></div>${m}</div>`;});

  // 추이 차트
  const uc='#34d399',dc='#f87171';
  const bgC=moTot.map((v,i)=>i===0||v>=moTot[i-1]?uc+'88':dc+'88');
  const bdC=moTot.map((v,i)=>i===0||v>=moTot[i-1]?uc:dc);
  if(monthlyTrendChart) monthlyTrendChart.destroy();
  monthlyTrendChart=new Chart(document.getElementById('monthlyTrendChart'),{
    data:{labels:months,datasets:[
      {label:'생산량',type:'bar',data:moTot,backgroundColor:bgC,borderColor:bdC,borderWidth:1.5},
      {label:'추이',type:'line',data:moTot,borderColor:'#60a5fa',borderWidth:2,
       pointRadius:4,pointBackgroundColor:'#60a5fa',fill:false,tension:0.3}
    ]},
    options:{
      responsive:true,animation:false,
      scales:{
        x:{ticks:{color:'#94a3b8',font:{size:10}},grid:{color:'#1e3a5f55'}},
        y:{ticks:{color:'#94a3b8',font:{size:9}},grid:{color:'#1e3a5f55'}}
      },
      plugins:{legend:{display:false},tooltip:{callbacks:{
        label:it=>{
          if(it.datasetIndex!==0) return '';
          const mom=momPct[it.dataIndex];
          const s=mom===null?'':(parseFloat(mom)>=0?' ▲+':'  ▼')+mom+'%';
          return it.raw.toLocaleString()+'개'+s;
        }
      }}}
    }
  });

  // 통계 요약
  const totAll=moTot.reduce((a,b)=>a+b,0);
  const avgMo=months.length?Math.round(totAll/months.length):0;
  const maxI=moTot.indexOf(Math.max(...moTot));
  const minI=moTot.indexOf(Math.min(...moTot));
  document.getElementById('monthlyStatsRow').innerHTML=
    '<div class="stat"><div class="v">'+months.length+'개월</div><div class="l">분석 기간</div></div>'+
    '<div class="stat"><div class="v">'+totAll.toLocaleString()+'</div><div class="l">전체 생산량</div></div>'+
    '<div class="stat"><div class="v">'+avgMo.toLocaleString()+'</div><div class="l">월 평균</div></div>'+
    '<div class="stat"><div class="v" style="color:#34d399">'+(months[maxI]||'-')+'</div><div class="l">최다 생산월</div></div>'+
    '<div class="stat"><div class="v" style="color:#f87171">'+(months[minI]||'-')+'</div><div class="l">최소 생산월</div></div>';

  // 상세 테이블
  const machHds=machList.map(m=>'<th style="color:'+MACH_COLORS[m]+'">'+m+'</th>').join('');
  const rows=months.map((mo,i)=>{
    const tot=moTot[i];
    const mom=momPct[i];
    const mc=mom===null?'#64748b':parseFloat(mom)>=0?'#34d399':'#f87171';
    const ms=mom===null?'-':(parseFloat(mom)>=0?'▲+':'▼')+mom+'%';
    const days=[...new Set(filtDates.filter(d=>d.startsWith(mo)))].length;
    const avg=days>0?Math.round(tot/days):0;
    const jobs=D.scatter.filter(r=>r.d.startsWith(mo)&&activeMachines.has(r.m)).length;
    const mcs=machList.map(m=>'<td style="text-align:right;color:#94a3b8">'+(moData[mo][m]||0).toLocaleString()+'</td>').join('');
    return '<tr><td style="font-weight:700;color:#7dd3fc">'+mo+'</td>'+
      '<td style="text-align:right;font-weight:700">'+tot.toLocaleString()+'</td>'+
      '<td style="text-align:right">'+jobs.toLocaleString()+'</td>'+
      '<td style="text-align:right;font-weight:700;color:'+mc+'">'+ms+'</td>'+
      '<td style="text-align:right">'+avg.toLocaleString()+'</td>'+
      '<td style="text-align:right">'+days+'</td>'+mcs+'</tr>';
  }).join('');
  document.getElementById('monthlyTableWrap').innerHTML=
    '<table><thead><tr><th>월</th><th>총 생산량</th><th>작업건수</th>'+
    '<th>전월대비</th><th>일평균</th><th>가동일수</th>'+machHds+'</tr></thead>'+
    '<tbody>'+rows+'</tbody></table>';
}
"""

insert_target = "updatePeriodFilter();\nbuildToggles();\nrenderCurrent();"
content = content.replace(insert_target, fn + "\n" + insert_target)
print("6:", "function renderMonthly" in content)

with open('c:/Users/admin/Desktop/work/자동화/CNC분석/CNC_일별생산현황_v3.html', 'w', encoding='utf-8') as f:
    f.write(content)
print("Done")
