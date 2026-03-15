import sys, re
sys.stdout.reconfigure(encoding='utf-8')

path = 'c:/Users/admin/Desktop/work/자동화/CNC분석/CNC_일별생산현황_v3.html'
with open(path, encoding='utf-8') as f:
    content = f.read()

# ── 1. handleXlsxUpload 리팩토링 ──
# file.name → fileName 치환 (함수 내부만)
old_fn = '''function handleXlsxUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  const st = document.getElementById('uploadStatus');
  st.className = 'upload-status loading';
  st.textContent = '⏳ 파일 읽는 중... ' + file.name;

  const reader = new FileReader();
  reader.onload = function(e){
    try {
      const data = new Uint8Array(e.target.result);'''

new_fn = '''function processXlsxData(data, fileName){
  const st = document.getElementById('uploadStatus');
  st.className = 'upload-status loading';
  st.textContent = '⏳ 파일 읽는 중... ' + fileName;
  try {'''

if old_fn in content:
    content = content.replace(old_fn, new_fn)
    print("1a: 함수 시작부 변환 OK")
else:
    print("1a: FAIL")

# success line: file.name → fileName
old_ok = "st.textContent = `✅ ${file.name} 로드 완료"
new_ok = "st.textContent = `✅ ${fileName} 로드 완료"
if old_ok in content:
    content = content.replace(old_ok, new_ok)
    print("1b: file.name → fileName OK")
else:
    print("1b: FAIL")

# 함수 끝부분 변환: catch + reader wrapper 제거하고 processXlsxData 닫기 + handleXlsxUpload 래퍼 추가
old_end = '''    } catch(err) {
      const st2 = document.getElementById('uploadStatus');
      st2.className='upload-status err';
      st2.textContent='❌ 오류: ' + err.message;
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}'''

new_end = '''  } catch(err) {
    const st2 = document.getElementById('uploadStatus');
    st2.className='upload-status err';
    st2.textContent='❌ 오류: ' + err.message;
    console.error(err);
  }
}

function handleXlsxUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = function(e){
    processXlsxData(new Uint8Array(e.target.result), file.name);
  };
  reader.readAsArrayBuffer(file);
}

// ── 자동 로드 (로컬 서버 모드: ?autoload=파일명) ──
window.addEventListener('load', function(){
  const params = new URLSearchParams(window.location.search);
  const autoFile = params.get('autoload');
  if(!autoFile) return;
  const st = document.getElementById('uploadStatus');
  st.className = 'upload-status loading';
  st.textContent = '⏳ 자동 로딩 중... ' + autoFile;
  fetch(autoFile)
    .then(r => { if(!r.ok) throw new Error('파일 로드 실패 ('+r.status+')'); return r.arrayBuffer(); })
    .then(buf => processXlsxData(new Uint8Array(buf), autoFile))
    .catch(e => {
      st.className = 'upload-status err';
      st.textContent = '⚠️ 자동 로드 실패: ' + e.message;
    });
});'''

if old_end in content:
    content = content.replace(old_end, new_end)
    print("1c: 함수 끝부분 변환 OK")
else:
    print("1c: FAIL")

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print("Done")
