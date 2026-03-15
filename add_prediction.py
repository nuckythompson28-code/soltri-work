import sys, re
sys.stdout.reconfigure(encoding='utf-8')

path = 'c:/Users/admin/Desktop/work/자동화/CNC분석/CNC_일별생산현황_v3.html'
with open(path, encoding='utf-8') as f:
    content = f.read()

# ── 1. 월별 탭 HTML에 예측 카드 추가 (bar chart 카드 앞에) ──
pred_card_html = """  <div class="card" id="monthlyPredCard" style="display:none">
    <div class="card-hdr">
      <h3>🔮 이달 생산량 예측 · 전월 대비 달성 전망</h3>
      <span class="note" id="predNote"></span>
    </div>
    <div class="card-body" style="padding:12px 16px">
      <div id="predContent"></div>
    </div>
  </div>
"""

old_bar_card = """  <div class="card">
    <div class="card-hdr">
      <h3>📅 월별 생산수량 비교 (호기별 누적)</h3>"""

new_bar_card = pred_card_html + """  <div class="card">
    <div class="card-hdr">
      <h3>📅 월별 생산수량 비교 (호기별 누적)</h3>"""

if old_bar_card in content:
    content = content.replace(old_bar_card, new_bar_card)
    print("1: 예측 카드 HTML 삽입 OK")
else:
    print("1: FAIL - 대상 문자열 없음")

# ── 2. renderMonthly() 함수 끝 (상세 테이블 렌더 직후)에 예측 로직 추가 ──
pred_js = """
  // ── 이달 예측 ──
  (function(){
    const predCard = document.getElementById('monthlyPredCard');
    const predContent = document.getElementById('predContent');
    const predNote = document.getElementById('predNote');
    if(!predCard) return;

    const lastMo = months[months.length-1];
    const now = new Date();
    const todayMo = now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');
    const isCurrentMo = (lastMo === todayMo);

    if(!isCurrentMo || months.length < 2){
      predCard.style.display='none';
      return;
    }
    predCard.style.display='';

    const prevMoIdx = months.length-2;
    const prevMo = months[prevMoIdx];
    const prevTot = moTot[prevMoIdx];
    const curTot  = moTot[months.length-1];

    // 이달 가동일수 & 일평균
    const curWorkDays = [...new Set(filtDates.filter(d=>d.startsWith(lastMo)))].length;
    const dailyAvg = curWorkDays>0 ? curTot/curWorkDays : 0;

    // 이달 달력 일수 · 오늘 기준 남은 캘린더 일수
    const [cy,cm] = lastMo.split('-').map(Number);
    const totalCalDays = new Date(cy, cm, 0).getDate();
    const todayDay = now.getDate();
    const remainCalDays = totalCalDays - todayDay;

    // 전월 가동 밀도 (가동일수/달력일수)
    const [py,pm] = prevMo.split('-').map(Number);
    const prevCalDays = new Date(py, pm, 0).getDate();
    const prevWorkDays = [...new Set(filtDates.filter(d=>d.startsWith(prevMo)))].length;
    const workRatio = prevWorkDays>0 ? prevWorkDays/prevCalDays : 0.7;

    // 남은 예상 가동일수
    const estRemainDays = Math.max(0, Math.round(remainCalDays * workRatio));

    // 예상 월말 생산량
    const projected = Math.round(curTot + dailyAvg * estRemainDays);
    const gap = projected - prevTot;
    const gapPct = prevTot>0 ? (gap/prevTot*100).toFixed(1) : '0.0';
    const isAhead = gap >= 0;

    // 전월 달성률 (현재까지)
    const progressPct = prevTot>0 ? Math.min(100,(curTot/prevTot*100)).toFixed(1) : '0.0';

    // 전월 동기 생산량 (이달 경과일과 같은 날수 기준)
    const prevFiltDates = filtDates.filter(d=>d.startsWith(prevMo)).sort();
    const prevSameDayTot = prevFiltDates.slice(0,curWorkDays).reduce((s,d)=>s+machList.reduce((ms,m)=>ms+((D.bar[d]||{})[m]||0),0),0);
    const vsSync = curTot - prevSameDayTot;
    const vsSyncPct = prevSameDayTot>0 ? (vsSync/prevSameDayTot*100).toFixed(1) : null;

    // 전월 달성하려면 남은 일평균 필요량
    const needTotal = prevTot - curTot;
    const needPerDay = estRemainDays>0 ? Math.ceil(needTotal/estRemainDays) : null;

    // 색상
    const gc = isAhead ? '#34d399' : '#f87171';
    const sc = vsSync>=0 ? '#34d399' : '#f87171';

    predNote.textContent = todayDay+'일 기준 · 남은 예상 가동 '+estRemainDays+'일';

    predContent.innerHTML =
      // 상단 요약 KPI 행
      '<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:14px">'+
        '<div style="flex:1;min-width:130px;background:#0f2744;border-radius:8px;padding:10px 14px;border-left:3px solid #60a5fa">'+
          '<div style="font-size:11px;color:#94a3b8;margin-bottom:4px">현재까지 생산량</div>'+
          '<div style="font-size:20px;font-weight:700;color:#e0e8f0">'+curTot.toLocaleString()+'<span style="font-size:11px;color:#94a3b8">개</span></div>'+
          '<div style="font-size:10px;color:#94a3b8">일평균 '+Math.round(dailyAvg).toLocaleString()+'개 · '+curWorkDays+'일 가동</div>'+
        '</div>'+
        '<div style="flex:1;min-width:130px;background:#0f2744;border-radius:8px;padding:10px 14px;border-left:3px solid #60a5fa">'+
          '<div style="font-size:11px;color:#94a3b8;margin-bottom:4px">전월 동기 대비</div>'+
          '<div style="font-size:20px;font-weight:700;color:'+sc+'">'+(vsSync>=0?'▲+':'▼')+vsSync.toLocaleString()+'<span style="font-size:11px">개</span></div>'+
          '<div style="font-size:10px;color:#94a3b8">'+(vsSyncPct!==null?(vsSync>=0?'+':'')+vsSyncPct+'%':'')+'</div>'+
        '</div>'+
        '<div style="flex:1;min-width:130px;background:#0f2744;border-radius:8px;padding:10px 14px;border-left:3px solid '+gc+'">'+
          '<div style="font-size:11px;color:#94a3b8;margin-bottom:4px">📌 월말 예상 생산량</div>'+
          '<div style="font-size:20px;font-weight:700;color:'+gc+'">'+projected.toLocaleString()+'<span style="font-size:11px;color:#94a3b8">개</span></div>'+
          '<div style="font-size:10px;color:'+gc+'">'+(isAhead?'▲+':'▼')+Math.abs(gap).toLocaleString()+'개 ('+gapPct+'%)</div>'+
        '</div>'+
        '<div style="flex:1;min-width:130px;background:#0f2744;border-radius:8px;padding:10px 14px;border-left:3px solid #a78bfa">'+
          '<div style="font-size:11px;color:#94a3b8;margin-bottom:4px">전월 달성 목표</div>'+
          '<div style="font-size:20px;font-weight:700;color:#e0e8f0">'+prevTot.toLocaleString()+'<span style="font-size:11px;color:#94a3b8">개</span></div>'+
          '<div style="font-size:10px;color:#94a3b8">전월('+prevMo+') 실적</div>'+
        '</div>'+
      '</div>'+
      // 달성률 프로그레스 바
      '<div style="margin-bottom:12px">'+
        '<div style="display:flex;justify-content:space-between;font-size:11px;color:#94a3b8;margin-bottom:4px">'+
          '<span>전월 대비 달성률</span><span style="color:'+gc+'">'+progressPct+'%</span>'+
        '</div>'+
        '<div style="height:10px;background:#1e3a5f;border-radius:5px;overflow:hidden">'+
          '<div style="height:100%;width:'+Math.min(100,parseFloat(progressPct))+'%;background:linear-gradient(90deg,'+gc+','+gc+'aa);border-radius:5px;transition:width .4s"></div>'+
        '</div>'+
      '</div>'+
      // 예측 달성률 (projected/prev)
      (function(){
        const projPct = prevTot>0 ? Math.min(120,(projected/prevTot*100)).toFixed(1) : '0.0';
        const pw = Math.min(100,parseFloat(projPct));
        return '<div style="margin-bottom:12px">'+
          '<div style="display:flex;justify-content:space-between;font-size:11px;color:#94a3b8;margin-bottom:4px">'+
            '<span>월말 예상 달성률</span><span style="color:'+gc+'">'+projPct+'%</span>'+
          '</div>'+
          '<div style="height:10px;background:#1e3a5f;border-radius:5px;overflow:hidden">'+
            '<div style="height:100%;width:'+pw+'%;background:linear-gradient(90deg,'+(isAhead?'#34d399':'#f87171')+','+(isAhead?'#34d399aa':'#f87171aa')+');border-radius:5px;opacity:0.7"></div>'+
          '</div>'+
        '</div>';
      })()+
      // 하단 메시지
      '<div style="padding:10px 14px;border-radius:8px;background:'+(isAhead?'#052e16':'#2d0c0c')+';border:1px solid '+(isAhead?'#34d399':'#f87171')+';font-size:12px;color:#e0e8f0">'+
        (isAhead
          ? '✅ 현재 페이스 유지 시 전월 대비 <strong style="color:#34d399">+'+Math.abs(gap).toLocaleString()+'개 (▲'+Math.abs(parseFloat(gapPct))+'%)</strong> 초과 달성 예상'
          : '⚠️ 전월 달성을 위해 남은 '+estRemainDays+'일간 일평균 <strong style="color:#f87171">'+(needPerDay!==null?needPerDay.toLocaleString():'-')+'개</strong> 필요'+
            ' (현재 일평균 '+Math.round(dailyAvg).toLocaleString()+'개 → '+(needPerDay!==null&&needPerDay>Math.round(dailyAvg)?'<strong style="color:#f87171">일 '+(needPerDay-Math.round(dailyAvg)).toLocaleString()+'개 증산 필요</strong>':'달성 가능')+')'
        )+
      '</div>';
  })();
"""

# renderMonthly 함수 끝 (닫는 }) 직전에 삽입
old_fn_end = """  document.getElementById('monthlyTableWrap').innerHTML=
    '<table><thead><tr><th>월</th><th>총 생산량</th><th>작업건수</th>'+
    '<th>전월대비</th><th>일평균</th><th>가동일수</th>'+machHds+'</tr></thead>'+
    '<tbody>'+rows+'</tbody></table>';
}"""

new_fn_end = """  document.getElementById('monthlyTableWrap').innerHTML=
    '<table><thead><tr><th>월</th><th>총 생산량</th><th>작업건수</th>'+
    '<th>전월대비</th><th>일평균</th><th>가동일수</th>'+machHds+'</tr></thead>'+
    '<tbody>'+rows+'</tbody></table>';
""" + pred_js + "\n}"

if old_fn_end in content:
    content = content.replace(old_fn_end, new_fn_end)
    print("2: 예측 JS 삽입 OK")
else:
    print("2: FAIL - 대상 문자열 없음")

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print("Done")
