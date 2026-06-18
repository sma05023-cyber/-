<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>교육청 인사/행정 관리 대시보드</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
        :root { --primary-color: #1a73e8; --primary-hover: #1557b0; --bg-color: #f8f9fa; --sidebar-bg: #ffffff; --text-main: #202124; --text-muted: #5f6368; --border-color: #dadce0; --row-green: #e6f4ea; --row-yellow: #fef7e0; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Noto Sans KR', sans-serif; background-color: var(--bg-color); color: var(--text-main); display: flex; height: 100vh; overflow: hidden; }
        .sidebar { width: 260px; background-color: var(--sidebar-bg); border-right: 1px solid var(--border-color); display: flex; flex-direction: column; }
        .sidebar-header { padding: 24px 20px; border-bottom: 1px solid var(--border-color); }
        .sidebar-header h2 { font-size: 18px; font-weight: 700; color: var(--primary-color); display: flex; align-items: center; gap: 8px; }
        .nav-menu { list-style: none; padding: 12px 0; flex-grow: 1; }
        .nav-item { padding: 12px 24px; cursor: pointer; display: flex; align-items: center; gap: 12px; font-size: 15px; font-weight: 500; color: var(--text-muted); transition: background 0.2s; }
        .nav-item:hover { background-color: #f1f3f4; }
        .nav-item.active { color: var(--primary-color); background-color: #e8f0fe; border-radius: 0 24px 24px 0; margin-right: 12px; }
        .main-content { flex-grow: 1; display: flex; flex-direction: column; overflow: hidden; }
        .content-area { padding: 32px; overflow-y: auto; flex-grow: 1; }
        .section { display: none; background: #fff; border-radius: 8px; padding: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border: 1px solid var(--border-color); }
        .section.active { display: block; }
        h3 { margin-bottom: 16px; font-size: 18px; color: var(--text-main); }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 500; margin-bottom: 8px; font-size: 14px; }
        
        input[type="text"], input[type="date"], input[type="file"] { 
            padding: 10px 14px; border: 1px solid var(--border-color); border-radius: 6px; 
            font-family: inherit; font-size: 14px; width: 100%; max-width: 400px; outline: none; 
        }
        input[type="text"]:focus { border-color: var(--primary-color); }
        
        .search-bar-container { display: flex; gap: 12px; margin-bottom: 16px; align-items: center; background: #f1f3f4; padding: 12px; border-radius: 8px; }
        .search-bar-container .material-icons { color: var(--text-muted); }
        
        .btn { background-color: var(--primary-color); color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; display: inline-flex; align-items: center; gap: 8px; transition: background 0.2s; }
        .btn:hover { background-color: var(--primary-hover); }
        table { width: 100%; border-collapse: collapse; margin-top: 16px; text-align: left; }
        th, td { padding: 14px 16px; border-bottom: 1px solid var(--border-color); font-size: 14px; }
        th { background-color: #f8f9fa; font-weight: 700; color: var(--text-muted); position: sticky; top: 0; }
        tr.bg-green { background-color: var(--row-green) !important; }
        tr.bg-yellow { background-color: var(--row-yellow) !important; }
        .badge { display: inline-block; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: 700; }
        .badge-leave { background-color: #fce8e6; color: #d93025; }
        .badge-namesake { background-color: #ea4335; color: #fff; } 
        .badge-type { background-color: #f1f3f4; color: #5f6368; margin-right: 8px; }
        .badge-source { background-color: #e8eaed; color: #3c4043; font-size: 11px; margin-bottom: 4px; display: inline-block; }
        .badge-excel { background-color: #1d6f42; color: #fff; font-size: 11px; margin-bottom: 4px; display: inline-block; }
    </style>
</head>
<body>

    <aside class="sidebar">
        <div class="sidebar-header">
            <span class="material-icons">account_balance</span>
            <h2>인사행정 시스템</h2>
        </div>
        <ul class="nav-menu">
            <li class="nav-item active" onclick="switchTab('upload', this)"><span class="material-icons">upload_file</span> 데이터 업로드</li>
            <li class="nav-item" onclick="switchTab('settings', this)"><span class="material-icons">calendar_today</span> 근무기간 계산 기준일</li>
            <li class="nav-item" onclick="switchTab('dashboard', this)"><span class="material-icons">dashboard</span> 인사 내신/만기자 조회</li>
            <li class="nav-item" onclick="switchTab('export', this)"><span class="material-icons">download</span> 결과 내보내기</li>
        </ul>
    </aside>

    <main class="main-content">
        <div class="content-area">
            
            <section id="upload" class="section active">
                <h3>다중 인사 데이터 업로드</h3>
                <p style="margin-bottom: 16px; color: var(--text-muted); font-size: 14px;">
                    <span class="badge badge-type">✨ 투트랙 지능형 파서</span> <span class="badge badge-type">제자리 승진 보존</span> <span class="badge badge-type">동명이인 완벽 방어</span><br>
                    연도별, 양식별(도/지원청), 수정차수별 엑셀 파일을 모두 던져 넣으셔도 시스템이 알아서 분류하고 추적합니다.<br>
                    업로드가 완료되면 <strong>근무기간 계산 기준일</strong> 설정 화면으로 이동합니다.
                </p>

                <div class="form-group">
                    <label for="dataFile">인사발령 데이터 파일 (다중 선택 가능)</label>
                    <input type="file" id="dataFile" accept=".csv, .xlsx, .xls" multiple>
                </div>
                
                <button class="btn" onclick="handleMultipleFileUpload()">
                    <span class="material-icons">cloud_upload</span> 전체 업로드 및 병합 파싱
                </button>
            </section>

            <section id="settings" class="section">
                <h3>근무기간 계산 기준일</h3>
                <div class="form-group">
                    <label for="baseDate">기준일자 설정</label>
                    <input type="date" id="baseDate">
                </div>
                <button class="btn" onclick="saveBaseDate()"><span class="material-icons">save</span> 설정 저장 및 조회하기</button>
            </section>

            <section id="dashboard" class="section">
                <h3>인사 내신 및 만기 대상자 통합 조회</h3>
                <p style="margin-bottom: 16px; font-size: 14px;">기준일: <strong id="displayBaseDate" style="color: var(--primary-color);"></strong></p>
                
                <div class="search-bar-container">
                    <span class="material-icons">search</span>
                    <input type="text" id="searchInput" placeholder="이름, 지역, 발령지, 직급으로 검색해보세요..." onkeyup="filterTable()">
                </div>

                <table>
                    <thead>
                        <tr><th>출처/발령지</th><th>이름</th><th>직급</th><th>최근 발령일자</th><th>근무기간 (내신 기준)</th><th>비고</th></tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </section>

            <section id="export" class="section">
                <h3>데이터 내보내기</h3>
                <button class="btn" onclick="copyTable()"><span class="material-icons">content_copy</span> 표 전체 복사하기</button>
            </section>

        </div>
    </main>

    <script>
        let baseDate = new Date();
        let parsedData = []; 

        document.addEventListener('DOMContentLoaded', () => {
            document.getElementById('baseDate').value = baseDate.toISOString().split('T')[0];
            updateDisplayBaseDate();
        });

        // 엑셀 및 CSV 하이브리드 읽기 기능
        function readFileAsync(file) {
            return new Promise((resolve, reject) => {
                const ext = file.name.split('.').pop().toLowerCase();
                const reader = new FileReader();

                if (ext === 'xlsx' || ext === 'xls') {
                    reader.readAsArrayBuffer(file);
                    reader.onload = (e) => {
                        try {
                            const data = new Uint8Array(e.target.result);
                            const workbook = XLSX.read(data, {type: 'array'});
                            let allRows = [];
                            
                            workbook.SheetNames.forEach(sheetName => {
                                const worksheet = workbook.Sheets[sheetName];
                                const rows = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ""});
                                allRows = allRows.concat(rows);
                            });
                            
                            resolve({ name: file.name, rows: allRows, fileExt: 'excel' });
                        } catch (err) {
                            reject('엑셀 파일을 읽는 중 오류가 발생했습니다.');
                        }
                    };
                } else {
                    reader.readAsText(file, 'euc-kr');
                    reader.onload = () => {
                        try {
                            const text = reader.result;
                            const rows = parseCSVText(text);
                            resolve({ name: file.name, rows: rows, fileExt: 'csv' });
                        } catch (err) {
                            reject('CSV 파일을 읽는 중 오류가 발생했습니다.');
                        }
                    };
                }
                reader.onerror = error => reject(error);
            });
        }

        function parseCSVText(text) {
            let p = '', row = [], ret = [row], i = 0, r = 0, s = !0, l;
            for (l of text) {
                if ('"' === l) { if (s && l === p) row[i] = (row[i] || "") + l; s = !s; }
                else if (',' === l && s) row[++i] = '';
                else if ('\n' === l && s) {
                    if ('\r' === p && row[i]) row[i] = row[i].slice(0, -1);
                    row = ret[++r] = []; i = 0;
                } else row[i] = (row[i] || "") + l;
                p = l;
            }
            return ret.map(r => r.map(c => c || ""));
        }

        function simplifyRank(rankRaw) {
            if(!rankRaw) return '';
            let rank = rankRaw.replace(/\n/g, '').trim();
            
            if(rank.includes('부이사관')) return '지방부이사관(3급)';
            if(rank.includes('서기관')) return '지방교육행정서기관(4급)';
            if(rank.includes('사무관')) return '지방교육행정사무관(5급)';
            if(rank.includes('주사보')) return '지방교육행정주사보(7급)';
            if(rank.includes('주사') && !rank.includes('보')) return '지방교육행정주사(6급)';
            if(rank.includes('서기보시보')) return '지방교육행정서기보시보'; 
            if(rank.includes('서기보') && !rank.includes('시보')) return '지방교육행정서기보(9급)';
            if(rank.includes('서기') && !rank.includes('보')) return '지방교육행정서기(8급)';
            
            return rank;
        }

        function cleanName(rawName) {
            return String(rawName).replace(/[★*\s]/g, '').trim();
        }

        async function handleMultipleFileUpload() {
            const fileInput = document.getElementById('dataFile');
            if (!fileInput.files.length) return;
            
            let rawExtractedData = [];
            const EXCLUDE_KEYWORDS = ['임명장', '임용장', '교부대상', '대상자 명단', '정년퇴직', '퇴직준비', '명예퇴직', '의원면직', '공로연수'];

            for (let f = 0; f < fileInput.files.length; f++) {
                try {
                    const fileObj = await readFileAsync(fileInput.files[f]);
                    const rows = fileObj.rows;
                    const fileFormat = fileObj.fileExt; 
                    const fileName = fileObj.name;
                    
                    const checkRowsText = rows.slice(0, 50).map(r => r.join(',')).join('\n');
                    
                    let commonAppointDate = "2024-01-01"; 
                    let dateMatch = checkRowsText.match(/20\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\./);
                    if(!dateMatch) {
                        dateMatch = fileName.match(/20\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\./);
                    }
                    if(dateMatch) {
                        const cleanDate = dateMatch[0].replace(/\s/g, '').replace(/\.$/, '').split('.');
                        commonAppointDate = `${cleanDate[0]}-${cleanDate[1].padStart(2,'0')}-${cleanDate[2].padStart(2,'0')}`;
                    }

                    let isDistrict = checkRowsText.includes("현 소 속") && (checkRowsText.replace(/\s/g, '').includes("발령사항") || checkRowsText.replace(/\s/g, '').includes("소속"));
                    let fileTypeStr = isDistrict ? "지원청" : "도교육청";

                    let currentCategory = "";
                    let skipCurrentCategory = false; 

                    if (isDistrict) {
                        for(let i=0; i<rows.length; i++) {
                            const row = rows[i];
                            const denseRow = row.map(c => String(c).trim()).filter(c => c !== '');
                            if (denseRow.length < 3) continue; 
                            
                            let loc = String(row[0] || "").trim();
                            let rowString = denseRow.join('');

                            if(!loc || loc.match(/^[<≪\[▣"▩]/) || loc.replace(/\s/g,'').includes('발령사항') || loc.replace(/\s/g,'').includes('소속') || loc.includes('인사규모') || loc.includes('표시자')) {
                                if(loc.match(/^[<≪\[]/) || rowString.includes('대상자') || rowString.includes('임지지정')) {
                                    currentCategory = loc || rowString;
                                    skipCurrentCategory = EXCLUDE_KEYWORDS.some(k => currentCategory.includes(k));
                                }
                                continue;
                            }
                            if (skipCurrentCategory) continue;

                            let rank = denseRow[1];
                            let name = "";
                            let oldLoc = "";

                            for (let c = 2; c < denseRow.length; c++) {
                                let val = denseRow[c].replace(/[★*\s]/g, '');
                                if (val === '성명' || val.includes('표시자')) continue;
                                
                                if (val.length >= 2 && val.length <= 5 && !val.endsWith('실장') && !val.endsWith('과장') && !val.endsWith('팀장') && !val.endsWith('센터장') && !val.includes('파견')) {
                                    name = val;
                                    oldLoc = denseRow.slice(c + 1).join(' ').replace(/[★*]/g, '').trim();
                                    break;
                                }
                            }

                            if(!name) continue;

                            let isLeave = currentCategory.includes('휴직') || loc.includes('휴직');
                            if(isLeave && !loc) loc = "휴직";

                            rawExtractedData.push({
                                sourceName: fileTypeStr,
                                fileExt: fileFormat,
                                location: loc.replace(/\n/g, ' '),
                                oldLocation: oldLoc.replace(/\n/g, ' '),
                                name: name,
                                rank: simplifyRank(rank),
                                appointDate: commonAppointDate,
                                isLeave: isLeave
                            });
                        }
                    } else {
                        for(let i=0; i<rows.length; i++) {
                            const denseRow = rows[i].map(c => String(c).trim()).filter(c => c !== '');
                            if (denseRow.length === 0) continue;
                            
                            const rowString = denseRow.join('');
                            let firstCellStr = denseRow[0];

                            if (rowString.includes('★') && rowString.includes('표시자')) continue;

                            if(firstCellStr.match(/^[<≪\[]/) || rowString.includes('대상자') || rowString.includes('교부대상')) { 
                                currentCategory = rowString; 
                                skipCurrentCategory = EXCLUDE_KEYWORDS.some(keyword => currentCategory.includes(keyword));
                                continue; 
                            }
                            
                            if (skipCurrentCategory) continue;

                            if(denseRow.length >= 4 && /^\d+$/.test(firstCellStr)) {
                                let rawName = denseRow[1];
                                if(!rawName || rawName === '성명' || rawName.includes('표시자')) continue;
                                let name = cleanName(rawName); 
                                
                                let rank = simplifyRank(denseRow[2]);
                                
                                let locIdx = 3;
                                while (locIdx < denseRow.length && /^\d+$/.test(denseRow[locIdx])) {
                                    locIdx++; 
                                }
                                
                                let location = locIdx < denseRow.length ? denseRow[locIdx] : "";
                                let oldLocation = denseRow.slice(locIdx + 1).join(' ').trim(); 
                                
                                let isLeave = currentCategory.includes('휴직') || location.includes('휴직');
                                
                                rawExtractedData.push({
                                    sourceName: fileTypeStr,
                                    fileExt: fileFormat,
                                    location: location.replace(/\n/g, ' '),
                                    oldLocation: oldLocation.replace(/\n/g, ' '),
                                    name: name,
                                    rank: rank,
                                    appointDate: commonAppointDate,
                                    isLeave: isLeave
                                });
                            }
                        }
                    }
                } catch (error) {
                    console.error("파일 처리 중 오류:", error);
                }
            } 

            rawExtractedData.sort((a, b) => new Date(a.appointDate) - new Date(b.appointDate));

            let finalList = [];

            for (let row of rawExtractedData) {
                let matches = finalList.filter(e => e.name === row.name);

                if (matches.length === 0) {
                    row.isNamesake = false;
                    finalList.push(row);
                } else {
                    let cleanOldRow = row.oldLocation.replace(/[^가-힣]/g, '');
                    let cleanRowLoc = row.location.replace(/[^가-힣]/g, '');
                    let isNewHire = cleanOldRow.includes('신규'); 

                    let targetPerson = null;

                    for (let e of matches) {
                        let cleanOldExisting = e.oldLocation.replace(/[^가-힣]/g, '');
                        let cleanExistingLoc = e.location.replace(/[^가-힣]/g, '');

                        if (e.appointDate === row.appointDate) {
                            let isLinked = false;
                            if (cleanExistingLoc && cleanRowLoc && (cleanExistingLoc.includes(cleanRowLoc) || cleanRowLoc.includes(cleanExistingLoc))) isLinked = true;
                            if (cleanOldExisting && cleanOldRow && (cleanOldExisting.includes(cleanOldRow) || cleanOldRow.includes(cleanOldExisting))) isLinked = true;
                            if (isNewHire && cleanOldExisting.includes('신규')) isLinked = true;
                            if (cleanExistingLoc && cleanOldRow && cleanOldRow.includes(cleanExistingLoc)) isLinked = true;
                            if (cleanRowLoc && cleanOldExisting && cleanOldExisting.includes(cleanRowLoc)) isLinked = true;

                            if (isLinked) {
                                targetPerson = e;
                                break;
                            }
                        } else {
                            if (!cleanOldRow) continue; 
                            if (cleanExistingLoc.includes(cleanOldRow) || cleanOldRow.includes(cleanExistingLoc) || cleanOldRow.includes('휴직') || cleanExistingLoc.includes('휴직')) {
                                targetPerson = e;
                                break;
                            }
                        }
                    }

                    if (targetPerson) {
                        let targetDate = new Date(targetPerson.appointDate);
                        let rowDate = new Date(row.appointDate);

                        if (rowDate > targetDate) {
                            let cleanExistingLoc = targetPerson.location.replace(/[^가-힣]/g, '');
                            let isPromotionInPlace = cleanExistingLoc && cleanRowLoc && (cleanExistingLoc.includes(cleanRowLoc) || cleanRowLoc.includes(cleanExistingLoc));

                            if (isPromotionInPlace) {
                                targetPerson.rank = row.rank; 
                                if (row.location.length > targetPerson.location.length) {
                                    targetPerson.location = row.location;
                                }
                                targetPerson.oldLocation = row.oldLocation;
                            } else {
                                targetPerson.location = row.location;
                                targetPerson.oldLocation = row.oldLocation;
                                targetPerson.rank = row.rank;
                                targetPerson.appointDate = row.appointDate;
                                targetPerson.isLeave = row.isLeave;
                                targetPerson.sourceName = row.sourceName;
                                targetPerson.fileExt = row.fileExt;
                            }
                        } else if (rowDate.getTime() === targetDate.getTime()) {
                            if (targetPerson.sourceName === '도교육청' && row.sourceName === '지원청') {
                                targetPerson.location = row.location;
                                targetPerson.oldLocation = row.oldLocation;
                                targetPerson.rank = row.rank;
                                targetPerson.sourceName = row.sourceName;
                                targetPerson.fileExt = row.fileExt;
                            } 
                        }
                    } else {
                        row.isNamesake = true;
                        matches.forEach(m => m.isNamesake = true);
                        finalList.push(row); 
                    }
                }
            }

            parsedData = finalList;
            
            // 💡 파일 업로드 완료 후 '근무기간 계산 기준일' 탭으로 바로 이동
            switchTab('settings', document.querySelectorAll('.nav-item')[1]);
        }

        function filterTable() {
            const keyword = document.getElementById('searchInput').value.toLowerCase();
            const filteredData = parsedData.filter(emp => 
                emp.name.toLowerCase().includes(keyword) || 
                emp.location.toLowerCase().includes(keyword) ||
                emp.rank.toLowerCase().includes(keyword)
            );
            renderTable(filteredData);
        }

        function saveBaseDate() {
            const inputDate = document.getElementById('baseDate').value;
            if(!inputDate) return alert('날짜를 선택해주세요.');
            baseDate = new Date(inputDate);
            updateDisplayBaseDate();
            renderTable();
            switchTab('dashboard', document.querySelectorAll('.nav-item')[2]);
        }

        function updateDisplayBaseDate() {
            document.getElementById('displayBaseDate').innerText = `${baseDate.getFullYear()}년 ${baseDate.getMonth()+1}월 ${baseDate.getDate()}일`;
        }

        function calculateTenure(startDate, targetDate) {
            let start = new Date(startDate);
            let end = new Date(targetDate);
            let years = end.getFullYear() - start.getFullYear();
            let months = end.getMonth() - start.getMonth();
            let days = end.getDate() - start.getDate();

            if (days < 0) months--;
            if (months < 0) { years--; months += 12; }
            return { years, months, totalMonths: (years * 12) + months };
        }

        function renderTable(dataToRender = parsedData) {
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            if(dataToRender.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding: 40px; color:#5f6368;">조건에 맞는 데이터가 없습니다.</td></tr>';
                return;
            }

            dataToRender.forEach(emp => {
                const tenure = calculateTenure(emp.appointDate, baseDate);
                
                let rowClass = '';
                if (tenure.totalMonths >= 24) rowClass = 'bg-yellow';
                else if (tenure.totalMonths >= 18) rowClass = 'bg-green';

                let remarks = '';
                if(emp.isLeave) remarks += '<span class="badge badge-leave">휴직자</span> ';
                if(emp.isNamesake) remarks += '<span class="badge badge-namesake">동명이인</span> ';

                let extBadge = emp.fileExt === 'excel' 
                    ? `<span class="badge-excel" style="padding:2px 4px; border-radius:4px;">EXCEL</span>` 
                    : `<span style="background:#80868b; color:white; padding:2px 4px; border-radius:4px; font-size:11px;">CSV</span>`;

                const tr = document.createElement('tr');
                if(rowClass) tr.className = rowClass;
                
                tr.innerHTML = `
                    <td>
                        <span class="badge-source">${extBadge} ${emp.sourceName}</span><br>
                        ${emp.location}
                    </td>
                    <td><strong>${emp.name}</strong></td>
                    <td>${emp.rank}</td>
                    <td>${emp.appointDate}</td>
                    <td>${tenure.years}년 ${tenure.months}개월</td>
                    <td>${remarks}</td>
                `;
                tbody.appendChild(tr);
            });
        }

        function switchTab(sectionId, element) {
            document.querySelectorAll('.section').forEach(sec => sec.classList.remove('active'));
            document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
            document.getElementById(sectionId).classList.add('active');
            if(element) element.classList.add('active');
            
            // 💡 상단 헤더 타이틀 변경 로직 제거 완료
            
            if(sectionId === 'dashboard') {
                document.getElementById('searchInput').value = '';
                renderTable();
            }
        }

        function copyTable() {
            const table = document.querySelector('table');
            const range = document.createRange();
            range.selectNode(table);
            window.getSelection().removeAllRanges();
            window.getSelection().addRange(range);
            document.execCommand('copy');
            alert('테이블 전체가 복사되었습니다. 엑셀이나 한글에 바로 붙여넣기 하세요.');
            window.getSelection().removeAllRanges();
        }
    </script>
</body>
</html>
