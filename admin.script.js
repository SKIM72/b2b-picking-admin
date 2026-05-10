// Supabase 클라이언트 설정
const SUPABASE_URL = 'https://lmlpbjosdygnpqcnrwuj.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxtbHBiam9zZHlnbnBxY25yd3VqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTI5MDA2NjgsImV4cCI6MjA2ODQ3NjY2OH0.Jt1Al2Sl44fSlRMAsvRw5cBuKfXcMzeYyzE774stBuQ';
const supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// --- DOM 요소 ---
const contentArea = document.getElementById('content-area');
const navButtons = document.querySelectorAll('nav button');
const logoutButton = document.getElementById('logout-button');
const userEmailDisplay = document.getElementById('current-user-email');
const batchSummaryModal = document.getElementById('batch-summary-modal');


// --- 상태 관리 변수 (필터 유지용 추가) ---
let currentChannelId = localStorage.getItem('adminSelectedChannelId') || null;
let currentSort = { column: 'id', direction: 'desc' };
let currentPickingStatusData = [];
let currentBatchDetailsData = [];
let currentProductMasterData = [];

// 탭 이동 시에도 유지될 필터 상태 객체 선언
let currentFilters = { product_code: '', barcode: '', product_name: '' };
let currentPickingFilters = {
    startDate: new Date().toISOString().split('T')[0],
    endDate: new Date().toISOString().split('T')[0],
    orderNumber: '',
    recipient: ''
};
let currentBatchFilters = {
    date: new Date().toISOString().split('T')[0],
    batchNumber: ''
};

// ▼▼▼ [추가] 엑셀 다운로드 파일명 및 스타일 처리를 위한 헬퍼 함수 ▼▼▼
function getFormattedDateStr() {
    const now = new Date();
    const yy = String(now.getFullYear()).slice(-2);
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const min = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');
    return `${yy}${mm}${dd}_${hh}${min}${ss}`;
}

function applyExcelStyle(worksheet) {
    if (!worksheet['!ref']) return;
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const colWidths = [];

    for (let C = range.s.c; C <= range.e.c; ++C) {
        let maxWidth = 10; // 기본 최소 너비
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellAddress = { c: C, r: R };
            const cellRef = XLSX.utils.encode_cell(cellAddress);
            const cell = worksheet[cellRef];
            
            if (!cell) continue;

            // 셀 너비 계산 (한글/다바이트 문자는 넓게 계산)
            let cellText = cell.v ? String(cell.v) : "";
            let textLength = 0;
            for (let i = 0; i < cellText.length; i++) {
                textLength += (cellText.charCodeAt(i) > 255) ? 2.2 : 1.2;
            }
            if (textLength > maxWidth) maxWidth = textLength;

            // 스타일 적용 (xlsx-js-style 지원)
            if (R === 0) { // 헤더 행
                cell.s = {
                    fill: { fgColor: { rgb: "5A73E4" } }, // 테마의 기본 Primary 파란색
                    font: { color: { rgb: "FFFFFF" }, bold: true },
                    alignment: { horizontal: "center", vertical: "center" },
                    border: {
                        top: { style: "thin", color: { rgb: "CBD5E1" } },
                        bottom: { style: "thin", color: { rgb: "CBD5E1" } },
                        left: { style: "thin", color: { rgb: "CBD5E1" } },
                        right: { style: "thin", color: { rgb: "CBD5E1" } }
                    }
                };
            } else { // 데이터 행
                cell.s = {
                    alignment: { horizontal: "center", vertical: "center" },
                    border: {
                        top: { style: "thin", color: { rgb: "E2E8F0" } },
                        bottom: { style: "thin", color: { rgb: "E2E8F0" } },
                        left: { style: "thin", color: { rgb: "E2E8F0" } },
                        right: { style: "thin", color: { rgb: "E2E8F0" } }
                    }
                };
            }
        }
        colWidths[C] = { wch: Math.min(Math.ceil(maxWidth) + 3, 60) }; // 약간의 여백 추가, 최대 60으로 너비 제한
    }
    worksheet['!cols'] = colWidths;
}
// ▲▲▲ [추가] 엑셀 다운로드 헬퍼 함수 끝 ▲▲▲


// 모던 달력 라이브러리(Flatpickr) 동적 로드 함수
const loadCalendarLibrary = async () => {
    if (document.getElementById('flatpickr-css')) return;
    
    const css = document.createElement('link');
    css.id = 'flatpickr-css';
    css.rel = 'stylesheet';
    css.href = 'https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css';
    document.head.appendChild(css);
    
    await new Promise(r => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/flatpickr';
        script.onload = r;
        document.head.appendChild(script);
    });
    await new Promise(r => {
        const script = document.createElement('script');
        script.src = 'https://npmcdn.com/flatpickr/dist/l10n/ko.js';
        script.onload = r;
        document.head.appendChild(script);
    });
};

// 모던 달력 적용 함수
const applyModernCalendar = () => {
    if(window.flatpickr) {
        flatpickr('input[type="date"]', {
            locale: "ko",
            dateFormat: "Y-m-d",
            altInput: true,
            altFormat: "Y-m-d (D)"
        });
    }
};


// --- 초기화 및 권한 확인 ---
(async () => {
    const { data: { session }, error } = await supabaseClient.auth.getSession();
    
    if (error || !session || (session.user.email !== 'eowert72@gmail.com' && (!session.user.user_metadata || !session.user.user_metadata.is_admin))) {
        await supabaseClient.auth.signOut();
        alert('로그인이 필요하거나 관리자 권한이 없습니다.');
        window.location.href = 'login.html';
        return;
    }
    
    userEmailDisplay.textContent = session.user.email;
    
    const superAdminOnlyButtons = ['nav-users']; 
    superAdminOnlyButtons.forEach(id => {
        const button = document.getElementById(id);
        if (button && session.user.email !== 'eowert72@gmail.com') {
            button.style.display = 'none';
        }
    });

    await loadCalendarLibrary(); // 모던 달력 로드
    await setupChannelSwitcher();
    
    document.querySelectorAll('nav button').forEach(btn => btn.classList.remove('active'));
    
    const initialActiveButton = document.getElementById('nav-picking-status');
    if(initialActiveButton) initialActiveButton.classList.add('active');
    
    showPickingStatus();
})();


// --- 네비게이션 및 채널 선택기 ---
async function setupChannelSwitcher() {
    const nav = document.querySelector('nav');
    if (document.getElementById('channel-switcher-container')) return;

    const { data: channels, error } = await supabaseClient.from('channels').select('*').order('name');
    if (error) { console.error('채널 목록 로딩 실패', error); return; }

    const storedChannelId = localStorage.getItem('adminSelectedChannelId');
    const isValidStoredId = channels.some(c => c.id == storedChannelId);

    if (isValidStoredId) {
        currentChannelId = storedChannelId;
    } else if (channels.length > 0) {
        currentChannelId = channels[0].id;
        localStorage.setItem('adminSelectedChannelId', currentChannelId);
    } else {
        currentChannelId = null;
        localStorage.removeItem('adminSelectedChannelId');
    }

    const switcherHTML = `<div id="channel-switcher-container" style="display: flex; align-items: center;"><label for="channel-switcher" style="color: #64748b; font-weight: 600; font-size: 0.85rem; margin-right: 8px;">채널:</label><select id="channel-switcher" style="padding: 0.4rem 1.5rem 0.4rem 0.8rem; border: 1px solid #cbd5e1; border-radius: 6px; background-color: #ffffff; color: #334155; font-size: 0.85rem; cursor: pointer; outline: none; min-width: 120px;">${channels.map(c => `<option value="${c.id}" ${c.id == currentChannelId ? 'selected' : ''}>${c.name}</option>`).join('')}</select></div>`;
    
    // 메뉴바(nav) 항목 왼쪽 정렬 (채널 선택 바로 옆으로 배치)
    const btnWrapper = document.createElement('div');
    btnWrapper.style.display = 'flex';
    btnWrapper.style.gap = '0.5rem';
    while(nav.firstChild) {
        btnWrapper.appendChild(nav.firstChild);
    }
    nav.style.display = 'flex';
    nav.style.justifyContent = 'flex-start'; // 왼쪽 정렬
    nav.style.alignItems = 'center';
    nav.style.gap = '1.5rem'; // 채널 선택 드롭다운과 탭 사이 간격 추가
    nav.innerHTML = switcherHTML;
    nav.appendChild(btnWrapper);

    document.getElementById('channel-switcher').addEventListener('change', e => {
        currentChannelId = e.target.value;
        localStorage.setItem('adminSelectedChannelId', currentChannelId);
        const activeNav = document.querySelector('nav button.active');
        if (activeNav) { activeNav.click(); } else { showPickingStatus(); }
    });
}

navButtons.forEach(button => {
    button.addEventListener('click', (e) => {
        navButtons.forEach(btn => btn.classList.remove('active'));
        e.target.classList.add('active');
        
        // 다른 메뉴로 이동 시 grid 스타일 초기화
        if (e.target.id !== 'nav-picking-status') {
             contentArea.style.display = 'flex';
             contentArea.style.gridTemplateRows = '';
             contentArea.style.overflowY = 'auto';
             contentArea.style.height = '';
        }

        switch (e.target.id) {
            case 'nav-picking-status': showPickingStatus(); break;
            case 'nav-batch-management': showBatchManagement(); break;
            case 'nav-products': showProductMaster(); break;
            case 'nav-users': showUserManagement(); break;
            case 'nav-channels': showChannelManagement(); break;
        }
    });
});


// --- 이벤트 리스너 통합 ---
contentArea.addEventListener('click', function(e) {
    const target = e.target.closest('button');
    if (!target) return;
    
    // 새로고침 시 초기화 로직 연동
    if (target.id === 'refresh-picking-status-btn') {
        const todayStr = new Date().toISOString().split('T')[0];
        currentPickingFilters = { startDate: todayStr, endDate: todayStr, orderNumber: '', recipient: '' };
        
        const startPicker = document.getElementById('status-start-date-picker');
        const endPicker = document.getElementById('status-end-date-picker');
        const orderFilter = document.getElementById('order-number-filter');
        const recipientFilter = document.getElementById('recipient-filter');
        
        if(startPicker) startPicker._flatpickr ? startPicker._flatpickr.setDate(todayStr) : startPicker.value = todayStr;
        if(endPicker) endPicker._flatpickr ? endPicker._flatpickr.setDate(todayStr) : endPicker.value = todayStr;
        if(orderFilter) orderFilter.value = '';
        if(recipientFilter) recipientFilter.value = '';
        
        loadPickingStatusByDateRange();
        loadTodaySummary();
    }
    
    if (target.id === 'refresh-batch-details-btn') {
        currentBatchFilters = { date: new Date().toISOString().split('T')[0], batchNumber: '' };
        loadBatchDetails();
        showBatchManagement(); // UI 리렌더링
    }
    
    if (target.id === 'refresh-product-master-btn') {
        currentFilters = { product_code: '', barcode: '', product_name: '' };
        contentArea.querySelectorAll('.filter-input').forEach(input => input.value = '');
        renderProductTable();
    }

    if (target.id === 'refresh-channels-btn') showChannelManagement();
    if (target.id === 'refresh-users-btn') showUserManagement();

    // ▼▼▼ [수정] 다운로드 파일명 시간 추가 및 스타일 적용 ▼▼▼
    if (target.id === 'download-picking-status-btn') {
        if (!currentPickingStatusData || currentPickingStatusData.length === 0) {
            alert('다운로드할 데이터가 없습니다.'); return;
        }
        const worksheet = XLSX.utils.json_to_sheet(currentPickingStatusData);
        applyExcelStyle(worksheet);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '출고현황');
        XLSX.writeFile(workbook, `출고현황_${getFormattedDateStr()}.xlsx`);
    }

    if (target.id === 'download-batch-details-btn') {
        if (!currentBatchDetailsData || currentBatchDetailsData.length === 0) {
            alert('다운로드할 데이터가 없습니다.'); return;
        }
        const dataToExport = currentBatchDetailsData.map(item => ({
            '출고일': item.date,
            '수취인': item.recipient,
            '출고지 주소': item.destination_address,
            '출고지시번호': item.order_number,
            '상품명': item.product_name,
            '상품코드(바코드)': item.barcode,
            '지시수량': item.expected_quantity,
            '완료수량': item.picked_quantity
        }));
        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        applyExcelStyle(worksheet);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '출고상세');
        XLSX.writeFile(workbook, `출고상세_${getFormattedDateStr()}.xlsx`);
    }

    if (target.id === 'download-product-master-btn') {
        if (!currentProductMasterData || currentProductMasterData.length === 0) {
            alert('다운로드할 데이터가 없습니다.'); return;
        }
        const dataToExport = currentProductMasterData.map(p => ({
            '상품코드': p.product_code,
            '바코드': p.barcode,
            '상품명': p.product_name
        }));
        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        applyExcelStyle(worksheet);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '상품마스터');
        XLSX.writeFile(workbook, `상품마스터_${getFormattedDateStr()}.xlsx`);
    }
    // ▲▲▲ [수정] 다운로드 버튼 반영 끝 ▲▲▲

    if (target.id === 'show-batch-summary-btn') {
        const date = contentArea.querySelector('#work-date-picker').value;
        if (!date) {
            alert('먼저 예정일을 선택해주세요.');
            return;
        }
        showBatchSummaryModal(date);
    }
    if (target.id === 'query-batch-btn') loadBatchDetails();
    if (target.id === 'new-batch-btn') handleCreateNewBatch();
    if (target.classList.contains('upload-standard-btn')) handleStandardOrderUpload(target);
    if (target.classList.contains('upload-corn-btn')) handleCornOrderUpload(target);
    if (target.classList.contains('delete-batch-btn')) handleDeleteBatch(target.dataset.id);
    if (target.classList.contains('download-standard-template')) handleStandardTemplateDownload();
    if (target.id === 'add-channel-btn') handleAddChannel();
    if (target.classList.contains('delete-channel-btn')) handleDeleteChannel(target.dataset.id);
    if (target.classList.contains('search-button')) {
        currentFilters.product_code = contentArea.querySelector('#filter-prod-code').value.trim();
        currentFilters.barcode = contentArea.querySelector('#filter-prod-barcode').value.trim();
        currentFilters.product_name = contentArea.querySelector('#filter-prod-name').value.trim();
        renderProductTable();
    }
    if (target.classList.contains('reset-button')) {
        currentFilters = { product_code: '', barcode: '', product_name: '' };
        contentArea.querySelectorAll('.filter-input').forEach(input => input.value = '');
        renderProductTable();
    }
    if (target.classList.contains('sortable')) {
        const column = target.dataset.column;
        if (currentSort.column === column) { currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc'; }
        else { currentSort.column = column; currentSort.direction = 'asc'; }
        renderProductTable();
    }
    if (target.classList.contains('download-template')) handleProductTemplateDownload();
    if (target.classList.contains('upload-data')) handleProductUpload();
    if (target.id === 'upload-corn-button') handleCornUpload();
    if (target.classList.contains('delete-selected')) handleDeleteSelectedProducts();
    if (target.classList.contains('approve-user-button')) handleApproveUser(e);
});

contentArea.addEventListener('change', function(e) {
    if (e.target.classList.contains('select-all-checkbox')) {
        contentArea.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = e.target.checked);
    }
    
    // 파일 업로드 UI 이름 변경 이벤트
    if (e.target.type === 'file') {
        const fileNameSpan = e.target.closest('label')?.querySelector('.file-name');
        if (fileNameSpan) {
            if (e.target.files.length > 0) {
                fileNameSpan.textContent = e.target.files[0].name;
                fileNameSpan.style.color = 'var(--text-primary-color)';
                fileNameSpan.style.fontWeight = '500';
            } else {
                fileNameSpan.textContent = '선택된 파일 없음';
                fileNameSpan.style.color = 'var(--text-secondary-color)';
                fileNameSpan.style.fontWeight = 'normal';
            }
        }
    }
});


if (batchSummaryModal) {
    batchSummaryModal.addEventListener('click', (e) => {
        if (e.target.classList.contains('modal-overlay') || e.target.classList.contains('modal-close-btn')) {
            batchSummaryModal.style.display = 'none';
        }
        const selectedRow = e.target.closest('tr');
        if (selectedRow && selectedRow.dataset.batchNumber) {
            const batchNumber = selectedRow.dataset.batchNumber;
            contentArea.querySelector('#batch-number-input').value = batchNumber;
            batchSummaryModal.style.display = 'none';
            loadBatchDetails();
        }
    });
}


async function showBatchSummaryModal(date) {
    batchSummaryModal.style.display = 'flex';
    const tbody = batchSummaryModal.querySelector('#batch-summary-table tbody');
    tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 2rem;">현황을 불러오는 중...</td></tr>`;

    try {
        const { data, error } = await supabaseClient.rpc('get_batch_summary_by_date', {
            p_channel_id: currentChannelId,
            p_batch_date: date
        });

        if (error) throw error;

        if (data.length === 0) {
            tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 2rem;">해당 날짜에 출고 차수가 없습니다.</td></tr>`;
            return;
        }

        let tableHtml = '';
        data.forEach(batch => {
            const progress = (batch.order_count > 0) ? (batch.completed_count / batch.order_count * 100) : 0;

            // ▼▼▼ [수정] 주문 건수와 완료 건수를 비교하여 상태 동적 계산 ▼▼▼
            let displayStatus = batch.status || '대기';
            if (batch.order_count > 0) {
                if (batch.completed_count >= batch.order_count) {
                    displayStatus = '완료';
                } else if (batch.completed_count > 0) {
                    displayStatus = '진행중';
                }
            }
            // ▲▲▲ [수정] 끝 ▲▲▲

            let progressHtml;
            if (progress > 0) {
                progressHtml = `
                    <div style="display: flex; align-items: center; gap: 8px; justify-content: center;">
                        <div style="width: ${Math.max(progress * 0.5, 10)}px; height: 18px; background-color: #007bff; border-radius: 4px;"></div>
                        <span style="font-size: 0.9rem; font-weight: 500; color: #212529;">${progress.toFixed(1)}%</span>
                    </div>
                `;
            } else {
                progressHtml = `<span style="font-size: 0.9rem; font-weight: 500; color: #212529;">0.0%</span>`;
            }

            tableHtml += `
                <tr data-batch-number="${batch.batch_number}">
                    <td style="text-align: center; font-weight: 700;">${batch.batch_number}차</td>
                    <td style="text-align: center;">${displayStatus}</td>
                    <td style="text-align: center;">${batch.order_count}</td>
                    <td style="text-align: center;">${batch.completed_count}</td>
                    <td style="text-align: center;">${progressHtml}</td>
                </tr>
            `;
        });
        tbody.innerHTML = tableHtml;

    } catch (err) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; color:red; padding: 2rem;">오류: ${err.message}</td></tr>`;
    }
}


// 채널 관리 UI 여백 및 디자인 개선
async function showChannelManagement() {
    contentArea.innerHTML = `
        <div class="content-section" style="max-width: 900px; margin: 0 auto; width: 100%; padding-top: 1rem;">
            <div class="page-header" style="margin-bottom: 2rem;">
                <h2>채널 관리</h2>
                <div class="actions-group">
                    <button id="refresh-channels-btn" class="btn-secondary">새로고침</button>
                </div>
            </div>
            <div class="card" style="margin-bottom: 2rem;">
                <div class="card-header" style="background-color: #f8fafd;">새 채널 추가</div>
                <div class="card-body" style="padding: 1.5rem;">
                    <div style="display: flex; align-items: center; border: 1px solid #e2e8f0; border-radius: 6px; background-color: #fff; overflow: hidden; width: 100%;">
                        <input type="text" id="new-channel-name" placeholder="추가할 새 채널 이름을 입력하세요 (예: 쿠팡 윙)" style="flex-grow: 1; padding: 0.8rem 1rem; border: none; outline: none; background: transparent; font-size: 0.95rem;">
                        <button id="add-channel-btn" class="btn-primary" style="margin: 4px; padding: 0.6rem 1.5rem; border-radius: 4px; white-space: nowrap;">추가하기</button>
                    </div>
                </div>
            </div>
            <div class="card">
                <div class="card-header" style="display: flex; justify-content: space-between; align-items: center; background-color: #f8fafd;">
                    <span>운영 중인 채널 목록</span>
                    <span style="font-size: 0.8rem; font-weight: normal; color: #ef4444;">※ 채널 삭제 시 관련된 모든 상품 및 출고 데이터가 삭제됩니다.</span>
                </div>
                <ul id="channel-list" class="management-list"></ul>
            </div>
        </div>`;

    const newChannelInput = contentArea.querySelector('#new-channel-name');
    newChannelInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            handleAddChannel();
        }
    });

    await loadChannelsForManagement();
}

async function loadChannelsForManagement() {
    const list = contentArea.querySelector('#channel-list');
    list.innerHTML = `<li>불러오는 중...</li>`;
    const { data, error } = await supabaseClient.from('channels').select('*').order('name');
    if (error) { list.innerHTML = `<li>오류: ${error.message}</li>`; return; }
    if (data.length === 0) { list.innerHTML = `<li style="padding:2rem; text-align:center; color:#64748b;">생성된 채널이 없습니다.</li>`; return; }
    list.innerHTML = data.map(c => `<li class="management-list-item" style="padding: 1.2rem 1.5rem;"><span>${c.name}</span><button class="btn-danger delete-channel-btn" data-id="${c.id}">삭제</button></li>`).join('');
}

async function handleAddChannel() {
    const nameInput = contentArea.querySelector('#new-channel-name');
    const name = nameInput.value.trim();
    if (!name) { alert('채널 이름을 입력하세요.'); return; }
    const { error } = await supabaseClient.from('channels').insert({ name });
    if (error) { alert('채널 추가 실패: ' + error.message); }
    else { alert('채널이 추가되었습니다. 페이지를 새로고침합니다.'); location.reload(); }
}

async function handleDeleteChannel(channelId) {
    if (confirm('이 채널을 삭제하면 관련된 모든 상품, 출고 데이터가 삭제됩니다. 계속하시겠습니까?')) {
        const { error } = await supabaseClient.from('channels').delete().eq('id', channelId);
        if (error) {
            alert('채널 삭제 실패: ' + error.message);
        } else {
            if (localStorage.getItem('adminSelectedChannelId') === channelId) {
                localStorage.removeItem('adminSelectedChannelId');
            }
            alert('채널이 삭제되었습니다. 페이지를 새로고침합니다.');
            location.reload();
        }
    }
}

// =================================================================
// 출고 차수 관리 (디자인 개선 및 아이콘 변경, 필터 유지)
// =================================================================
async function showBatchManagement() {
    contentArea.innerHTML = `
        <div class="content-section" style="max-width: 1400px; margin: 0 auto; width: 100%;">
            <div class="sticky-controls">
                <div class="page-header" style="margin-bottom: 1.5rem;">
                    <h2>출고 업로드 및 확인</h2>
                     <div class="actions-group">
                        <button id="refresh-batch-details-btn" class="btn-secondary">초기화 및 새로고침</button>
                        <button id="download-batch-details-btn" class="btn-primary">엑셀 다운로드</button>
                    </div>
                </div>
                
                <div class="card" style="margin-bottom: 1.5rem;">
                    <div class="card-body" style="display: flex; flex-direction: row; align-items: center; flex-wrap: wrap; gap: 1.5rem; justify-content: space-between;">
                        <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 1.5rem;">
                            <div style="display: flex; align-items: center; gap: 0.5rem;">
                                <label style="font-weight: 600; font-size: 0.85rem; color: #64748b; white-space: nowrap;">예정일:</label>
                                <input type="date" id="work-date-picker" value="${currentBatchFilters.date}" style="width: 150px;">
                            </div>
                            <div style="display: flex; align-items: center; gap: 0.5rem;">
                                <label style="font-weight: 600; font-size: 0.85rem; color: #64748b; white-space: nowrap;">업로드차수:</label>
                                <input type="number" id="batch-number-input" placeholder="차수 입력" min="1" value="${currentBatchFilters.batchNumber}" style="width: 120px;">
                            </div>
                            
                            <button id="show-batch-summary-btn" title="차수 현황 보기" style="width: 38px; height: 38px; display: flex; align-items: center; justify-content: center; background-color: #f0f4ff; color: #5A73E4; border: none; border-radius: 6px; cursor: pointer; transition: background-color 0.2s;">
                                <span class="material-symbols-outlined" style="font-size: 20px;">assignment</span>
                            </button>
                        </div>
                        <div style="display: flex; gap: 0.5rem;">
                            <button id="new-batch-btn" class="btn-secondary">신규 차수 생성</button>
                            <button id="query-batch-btn" class="btn-primary" style="white-space: nowrap;">조회하기</button>
                        </div>
                    </div>
                </div>
                <div id="batch-summary-section"></div>
                <div id="batch-upload-section"></div>
            </div>
            <div id="batch-details-section" class="table-wrapper" style="margin-top: 1rem;">
            </div>
        </div>`;

    const batchInput = contentArea.querySelector('#batch-number-input');
    batchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            loadBatchDetails();
        }
    });

    // 필터 데이터가 이미 있다면 자동 조회
    if(currentBatchFilters.date && currentBatchFilters.batchNumber) {
        loadBatchDetails();
    }
    
    applyModernCalendar(); // 모던 달력 적용
}

async function loadBatchDetails() {
    const batchInput = contentArea.querySelector('#batch-number-input');
    const dateInput = contentArea.querySelector('#work-date-picker');
    const date = dateInput ? dateInput.value : currentBatchFilters.date;
    const batchNumber = batchInput ? batchInput.value : currentBatchFilters.batchNumber;
    
    // 상태 저장
    currentBatchFilters.date = date;
    currentBatchFilters.batchNumber = batchNumber;

    const uploadSection = contentArea.querySelector('#batch-upload-section');
    const summarySection = contentArea.querySelector('#batch-summary-section');
    const detailsSection = contentArea.querySelector('#batch-details-section');

    if (!date || !batchNumber) {
        detailsSection.innerHTML = `<p style="text-align:center; padding:3rem; color:#64748b;">날짜와 차수를 입력하고 조회해주세요.</p>`;
        uploadSection.innerHTML = '';
        summarySection.innerHTML = '';
        return;
    }

    detailsSection.innerHTML = `<p style="text-align:center; padding:3rem;">데이터를 불러오는 중입니다...</p>`;
    summarySection.innerHTML = '';
    currentBatchDetailsData = [];

    const { data: batchData } = await supabaseClient.from('picking_batches').select('id, status').eq('channel_id', currentChannelId).eq('batch_date', date).eq('batch_number', batchNumber).single();

    const batchId = batchData ? batchData.id : null;
    const batchStatus = batchData ? batchData.status : '신규';

    uploadSection.innerHTML = `
        <div class="control-grid" style="margin-top: 0;">
            <div class="card">
                <div class="card-header" style="background-color: #f8fafd;"><strong>${batchNumber}차수</strong> - 표준 양식</div>
                <div class="card-body" style="padding: 1.5rem;">
                    <button class="download-standard-template btn-secondary">양식 다운로드</button>
                    <div class="custom-file-input" style="flex-grow:1;">
                        <label>
                            <input type="file" class="standard-excel-input" accept=".xlsx, .xls">
                            <span class="file-name">선택된 파일 없음</span>
                            <span class="file-btn">파일 선택</span>
                        </label>
                    </div>
                    <button class="upload-standard-btn btn-primary" data-date="${date}" data-batch="${batchNumber}">업로드</button>
                </div>
            </div>
            <div class="card">
                <div class="card-header" style="background-color: #f8fafd;"><strong>${batchNumber}차수</strong> - CORN 양식</div>
                <div class="card-body" style="padding: 1.5rem;">
                    <div class="custom-file-input" style="flex-grow:1;">
                        <label>
                            <input type="file" class="corn-excel-input" accept=".xlsx, .xls">
                            <span class="file-name">선택된 파일 없음</span>
                            <span class="file-btn">파일 선택</span>
                        </label>
                    </div>
                    <button class="upload-corn-btn btn-primary" data-date="${date}" data-batch="${batchNumber}">업로드</button>
                </div>
            </div>
        </div>`;

    if (batchData) {
        try {
            const { data: ordersWithItems, error: ordersError } = await supabaseClient
                .from('picking_orders')
                .select(`order_number, recipient, destination_address, picking_items (product_name, barcode, expected_quantity, picked_quantity)`)
                .eq('batch_id', batchId)
                .order('order_number', { ascending: true });

            if (ordersError) throw ordersError;

            const allItems = [];
            let totalQuantity = 0, totalPicked = 0, orderCount = 0;

            if (ordersWithItems && ordersWithItems.length > 0) {
                orderCount = ordersWithItems.length;
                ordersWithItems.forEach(order => {
                    if (order.picking_items && order.picking_items.length > 0) {
                        order.picking_items.forEach(item => {
                            const pickedQty = item.picked_quantity || 0;
                            allItems.push({ ...item, date: date, order_number: order.order_number, recipient: order.recipient, destination_address: order.destination_address, picked_quantity: pickedQty });
                            totalQuantity += item.expected_quantity;
                            totalPicked += pickedQty;
                        });
                    }
                });
            }

            currentBatchDetailsData = allItems;

            // ▼▼▼ [수정] 지시수량과 완료수량을 비교하여 상태 동적 표시 ▼▼▼
            let displayStatus = batchStatus;
            if (totalQuantity > 0) {
                if (totalPicked >= totalQuantity) {
                    displayStatus = '완료';
                } else if (totalPicked > 0) {
                    displayStatus = '진행중';
                }
            }
            // ▲▲▲ [수정] 끝 ▲▲▲

            const summaryHTML = `
                <div class="card" style="margin-top: 1.5rem; margin-bottom: 1.5rem;">
                    <div class="card-header" style="display:flex; justify-content:space-between; align-items:center; background-color:#fff;">
                        <span style="font-size: 1.1rem;">${batchNumber}차수 현황 (상태: <span style="color:var(--primary-color); font-weight:bold;">${displayStatus}</span>) | 총 주문 <strong>${orderCount}</strong>건, 총 지시 <strong>${totalQuantity}</strong>개, 총 완료 <strong>${totalPicked}</strong>개</span>
                        ${batchId ? `<button class="delete-batch-btn" data-id="${batchId}" title="이 차수를 삭제합니다" style="background: none; border: none; color: #ef4444; font-weight: 600; font-size: 0.95rem; font-style: italic; cursor: pointer; display: flex; align-items: center; gap: 4px; padding: 0.5rem;"><span class="material-symbols-outlined" style="font-size: 1.25rem;">delete_outline</span></button>` : ''}
                    </div>
                </div>`;
            summarySection.innerHTML = summaryHTML;

            let tableHTML;
            if (allItems.length > 0) {
                 tableHTML = `
                    <table>
                        <thead>
                            <tr>
                                <th style="min-width: 140px; white-space: nowrap;">출고일</th>
                                <th style="min-width: 140px;">수취인</th>
                                <th style="width: 30%; min-width: 250px;">출고지 주소</th>
                                <th style="min-width: 140px;">출고지시번호</th>
                                <th style="width: auto; min-width: 180px;">상품명</th>
                                <th style="min-width: 150px;">상품코드 (바코드)</th>
                                <th style="min-width: 80px;">지시수량</th>
                                <th style="min-width: 80px;">완료수량</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${allItems.map(item => `
                                <tr class="${item.picked_quantity >= item.expected_quantity ? 'completed-row' : ''}">
                                    <td>${item.date}</td>
                                    <td>${item.recipient || ''}</td>
                                    <td>${item.destination_address || ''}</td>
                                    <td>${item.order_number}</td>
                                    <td>${item.product_name || '이름 없음'}</td>
                                    <td>${item.barcode}</td>
                                    <td>${item.expected_quantity}</td>
                                    <td>${item.picked_quantity}</td>
                                </tr>`).join('')}
                        </tbody>
                    </table>`;
            } else {
                 tableHTML = `<div class="card-body"><p style="text-align:center; color:#64748b; padding:2rem;">업로드된 주문 상세 내역이 없습니다.</p></div>`;
            }
            detailsSection.innerHTML = tableHTML;

        } catch (error) {
            detailsSection.innerHTML = `<div class="card-body"><p style="color:red; text-align:center; padding:2rem;">데이터를 불러오는 중 오류가 발생했습니다: ${error.message}</p></div>`;
            summarySection.innerHTML = '';
        }
    } else {
        detailsSection.innerHTML = `<div class="card-body"><p style="text-align:center; color:#64748b; padding:2rem;">데이터가 없습니다. 양식 파일을 업로드하세요.</p></div>`;
        summarySection.innerHTML = '';
    }
}


async function handleCreateNewBatch() {
    const batchInput = contentArea.querySelector('#batch-number-input');
    const date = contentArea.querySelector('#work-date-picker').value;
    if (!date) { alert('날짜를 먼저 선택하세요.'); return; }
    const { data } = await supabaseClient.from('picking_batches').select('batch_number').eq('channel_id', currentChannelId).eq('batch_date', date);
    const maxBatchNum = data && data.length > 0 ? Math.max(...data.map(b => b.batch_number)) : 0;
    const newBatchNumber = maxBatchNum + 1;
    batchInput.value = newBatchNumber;
    alert(`신규 ${newBatchNumber}차수가 설정되었습니다.`);
    loadBatchDetails();
}

// ▼▼▼ [수정] 표준 양식(공양식) 다운로드 시간 추가 및 스타일 적용 ▼▼▼
async function handleStandardTemplateDownload() {
    const headers = ["order_number", "recipient", "destination_address", "barcode", "product_name", "expected_quantity"];
    const worksheet = XLSX.utils.json_to_sheet([], { header: headers });
    applyExcelStyle(worksheet);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Orders");
    XLSX.writeFile(workbook, `standard_order_template_${getFormattedDateStr()}.xlsx`);
}
// ▲▲▲ [수정] 끝 ▲▲▲

async function handleStandardOrderUpload(target) {
    const fileInput = contentArea.querySelector('.standard-excel-input');
    const file = fileInput.files[0];
    if (!file) { return alert("표준 양식 엑셀 파일을 선택하세요."); }
    const { date, batch } = target.dataset;
    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            let { data: batchData, error: selectError } = await supabaseClient
                .from('picking_batches').select('id').eq('channel_id', currentChannelId).eq('batch_date', date).eq('batch_number', batch).single();
            if (selectError && selectError.code !== 'PGRST116') throw selectError;
            if (!batchData) {
                const { data: newBatchData, error: insertError } = await supabaseClient
                    .from('picking_batches').insert({ channel_id: currentChannelId, batch_date: date, batch_number: batch, status: '대기' }).select('id').single();
                if (insertError) throw insertError;
                batchData = newBatchData;
            }

            const excelData = new Uint8Array(event.target.result);
            const workbook = XLSX.read(excelData, { type: 'array' });
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            if (jsonData.length === 0) throw new Error('엑셀 파일에 데이터가 없습니다.');

            const orders = {};
            jsonData.forEach(row => {
                if (!row.order_number) return;
                const orderKey = row.order_number.toString();
                if (!orders[orderKey]) {
                    orders[orderKey] = { recipient: row.recipient || '', address: row.destination_address, items: [], total_expected: 0 };
                }
                const quantity = Number(row.expected_quantity) || 0;
                orders[orderKey].items.push({ barcode: row.barcode.toString(), product_name: row.product_name, expected_quantity: quantity });
                orders[orderKey].total_expected += quantity;
            });

            await supabaseClient.from('picking_orders').delete().eq('batch_id', batchData.id);
            for (const orderKey in orders) {
                const orderData = orders[orderKey];
                const { data: insertedOrder, error: orderError } = await supabaseClient.from('picking_orders').insert({
                    batch_id: batchData.id, order_number: orderKey, recipient: orderData.recipient,
                    destination_address: orderData.address, total_expected_quantity: orderData.total_expected
                }).select().single();
                if (orderError) throw orderError;
                const itemsToInsert = orderData.items.map(item => ({ ...item, order_id: insertedOrder.id }));
                await supabaseClient.from('picking_items').insert(itemsToInsert);
            }
            alert(`표준 양식 업로드 성공!`);
            loadBatchDetails();
        } catch (error) { alert('업로드 실패: ' + error.message); }
    };
    reader.readAsArrayBuffer(file);
}

async function handleCornOrderUpload(target) {
    const fileInput = contentArea.querySelector('.corn-excel-input');
    const file = fileInput.files[0];
    if (!file) { return alert("CORN 양식 엑셀 파일을 선택하세요."); }
    const { date, batch } = target.dataset;
    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            const { data: products } = await supabaseClient.from('products').select('product_code, barcode, product_name').eq('channel_id', currentChannelId);
            if (!products || products.length === 0) throw new Error('상품 마스터에 데이터가 없습니다.');
            const productMap = new Map(products.map(p => [p.product_code, p]));

            let { data: batchData, error: selectError } = await supabaseClient
                .from('picking_batches').select('id').eq('channel_id', currentChannelId).eq('batch_date', date).eq('batch_number', batch).single();
            if (selectError && selectError.code !== 'PGRST116') throw selectError;
            if (!batchData) {
                const { data: newBatchData, error: insertError } = await supabaseClient
                    .from('picking_batches').insert({ channel_id: currentChannelId, batch_date: date, batch_number: batch, status: '대기' }).select('id').single();
                if (insertError) throw insertError;
                batchData = newBatchData;
            }

            const excelData = new Uint8Array(event.target.result);
            const workbook = XLSX.read(excelData, { type: 'array' });
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
            const dataRows = rows.slice(1);
            if (dataRows.length === 0) throw new Error('엑셀 파일에 데이터가 없습니다.');

            const orders = {};
            let notFoundCodes = new Set();
            dataRows.forEach(row => {
                const orderKey = row[1];
                if (!orderKey) return;

                const productCode = row[28];
                const productInfo = productMap.get(productCode);
                if (!productInfo) {
                    notFoundCodes.add(productCode);
                    return;
                }

                if (!orders[orderKey]) {
                    orders[orderKey] = {
                        recipient: row[17] || '',
                        address: row[21] || '',
                        items: [],
                        total_expected: 0
                    };
                }

                const quantity = Number(row[34]) || 0;
                const barcode = productInfo.barcode;

                const existingItem = orders[orderKey].items.find(item => item.barcode === barcode);

                if (existingItem) {
                    existingItem.expected_quantity += quantity;
                } else {
                    orders[orderKey].items.push({
                        barcode: barcode,
                        product_name: productInfo.product_name,
                        expected_quantity: quantity
                    });
                }
                orders[orderKey].total_expected += quantity;
            });

            await supabaseClient.from('picking_orders').delete().eq('batch_id', batchData.id);
            for (const orderKey in orders) {
                const orderData = orders[orderKey];
                const { data: insertedOrder, error: orderError } = await supabaseClient.from('picking_orders').insert({
                    batch_id: batchData.id, order_number: orderKey, recipient: orderData.recipient,
                    destination_address: orderData.address, total_expected_quantity: orderData.total_expected
                }).select().single();
                if(orderError) throw orderError;
                const itemsToInsert = orderData.items.map(item => ({ ...item, order_id: insertedOrder.id }));
                await supabaseClient.from('picking_items').insert(itemsToInsert);
            }
            let successMessage = `CORN 양식 업로드 성공!`;
            if (notFoundCodes.size > 0) { successMessage += `\n\n주의: 누락된 상품코드: ${Array.from(notFoundCodes).join(', ')}`; }
            alert(successMessage);
            loadBatchDetails();
        } catch (error) { alert('업로드 실패: ' + error.message); }
    };
    reader.readAsArrayBuffer(file);
}

async function handleDeleteBatch(batchId) {
    if (!batchId || !confirm('이 차수의 모든 출고 데이터가 삭제됩니다. 계속하시겠습니까?')) return;
    const { error } = await supabaseClient.from('picking_batches').delete().eq('id', batchId);
    if (error) { alert('삭제 실패: ' + error.message); }
    else {
        alert('삭제되었습니다.');
        contentArea.querySelector('#batch-details-section').innerHTML = '';
        contentArea.querySelector('#batch-upload-section').innerHTML = '';
        contentArea.querySelector('#batch-summary-section').innerHTML = '';
    }
}

// =================================================================
// 🚀 당일 출고 현황 요약 
// =================================================================
async function loadTodaySummary() {
    const todayStr = new Date().toISOString().split('T')[0];
    
    const pendingEl = document.getElementById('summary-pending-count');
    const progressEl = document.getElementById('summary-progress-count');
    const completedEl = document.getElementById('summary-completed-count');

    if (!pendingEl || !progressEl || !completedEl) return;

    pendingEl.textContent = '...';
    progressEl.textContent = '...';
    completedEl.textContent = '...';

    try {
        const { data: batches, error: batchError } = await supabaseClient
            .from('picking_batches')
            .select('id')
            .eq('channel_id', currentChannelId)
            .eq('batch_date', todayStr);

        if (batchError) throw batchError;

        if (!batches || batches.length === 0) {
            pendingEl.textContent = '0건';
            progressEl.textContent = '0건';
            completedEl.textContent = '0건';
            return;
        }

        const batchIds = batches.map(b => b.id);

        const { data: orders, error: orderError } = await supabaseClient
            .from('picking_orders')
            .select('status')
            .in('batch_id', batchIds);

        if (orderError) throw orderError;

        let pending = 0;
        let progress = 0;
        let completed = 0;

        orders.forEach(order => {
            const status = order.status || '미검수'; 
            if (status === '완료') {
                completed++;
            } else if (status === '검수중') {
                progress++;
            } else {
                pending++;
            }
        });

        pendingEl.textContent = `${pending}건`;
        progressEl.textContent = `${progress}건`;
        completedEl.textContent = `${completed}건`;

    } catch (error) {
        console.error('요약 데이터 로드 실패:', error);
        pendingEl.textContent = '-건';
        progressEl.textContent = '-건';
        completedEl.textContent = '-건';
    }
}

// =================================================================
// 출고 현황 조회 (필터 상태 유지)
// =================================================================
async function showPickingStatus() {
    
    // Grid 레이아웃을 main content-area에 적용
    contentArea.style.display = 'grid';
    contentArea.style.height = `calc(100vh - ${contentArea.offsetTop}px)`;
    contentArea.style.gridTemplateRows = 'auto 1fr';
    contentArea.style.overflow = 'hidden';

    contentArea.innerHTML = `
        <div class="top-controls" style="padding-bottom: 1.5rem; z-index: 10;">
            <div class="page-header">
                <h2>출고 현황 조회</h2>
                <div class="actions-group">
                    <button id="refresh-picking-status-btn" class="btn-secondary">초기화 및 새로고침</button>
                    <button id="download-picking-status-btn" class="btn-primary">엑셀 다운로드</button>
                </div>
            </div>

            <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-bottom: 1.25rem;">
                <div style="background-color: white; padding: 1.25rem; border-radius: 0.5rem; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05); display: flex; align-items: center; gap: 1rem;">
                    <div style="width: 3rem; height: 3rem; background-color: #f0f4ff; border-radius: 0.5rem; display: flex; align-items: center; justify-content: center; color: #5A73E4;">
                        <span class="material-symbols-outlined" style="font-size: 28px;">assignment</span>
                    </div>
                    <div>
                        <p style="font-size: 13px; font-weight: 700; color: #94a3b8; margin-bottom: 0.125rem; letter-spacing: 0.025em;">대기 중인 출고</p>
                        <h2 id="summary-pending-count" style="font-size: 1.5rem; font-weight: 700; color: #1e293b; line-height: 1; margin: 0;">0건</h2>
                    </div>
                </div>
                <div style="background-color: white; padding: 1.25rem; border-radius: 0.5rem; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05); display: flex; align-items: center; gap: 1rem;">
                    <div style="width: 3rem; height: 3rem; background-color: #fff8e6; border-radius: 0.5rem; display: flex; align-items: center; justify-content: center; color: #f59e0b;">
                        <span class="material-symbols-outlined" style="font-size: 28px;">local_shipping</span>
                    </div>
                    <div>
                        <p style="font-size: 13px; font-weight: 700; color: #94a3b8; margin-bottom: 0.125rem; letter-spacing: 0.025em;">진행 중인 출고</p>
                        <h2 id="summary-progress-count" style="font-size: 1.5rem; font-weight: 700; color: #1e293b; line-height: 1; margin: 0;">0건</h2>
                    </div>
                </div>
                <div style="background-color: white; padding: 1.25rem; border-radius: 0.5rem; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05); display: flex; align-items: center; gap: 1rem;">
                    <div style="width: 3rem; height: 3rem; background-color: #ecfdf5; border-radius: 0.5rem; display: flex; align-items: center; justify-content: center; color: #10b981;">
                        <span class="material-symbols-outlined" style="font-size: 28px;">check_circle</span>
                    </div>
                    <div>
                        <p style="font-size: 13px; font-weight: 700; color: #94a3b8; margin-bottom: 0.125rem; letter-spacing: 0.025em;">오늘 완료</p>
                        <h2 id="summary-completed-count" style="font-size: 1.5rem; font-weight: 700; color: #1e293b; line-height: 1; margin: 0;">0건</h2>
                    </div>
                </div>
            </div>

            <div class="card" style="margin-bottom: 0;">
                <div class="card-body" style="display: flex; flex-direction: row; align-items: center; flex-wrap: wrap; gap: 1.5rem; justify-content: space-between;">
                    
                    <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 1.5rem;">
                        <div style="display: flex; gap: 0.5rem; align-items: center;">
                            <label style="font-weight: 600; font-size: 0.85rem; color: #64748b; white-space: nowrap;">날짜 범위:</label>
                            <input type="date" id="status-start-date-picker" value="${currentPickingFilters.startDate}">
                            <span style="color: #94a3b8; font-weight: bold;">~</span>
                            <input type="date" id="status-end-date-picker" value="${currentPickingFilters.endDate}">
                            <button id="today-btn" class="btn-secondary" style="margin-left: 0.25rem; font-size: 0.8rem; padding: 0.4rem 0.8rem; white-space: nowrap;">금일 세팅</button>
                        </div>
                        
                        <div style="display: flex; gap: 0.5rem; align-items: center;">
                            <label style="font-weight: 600; font-size: 0.85rem; color: #64748b; white-space: nowrap;">필터 검색:</label>
                            <input type="text" id="order-number-filter" placeholder="출고지시번호 검색" class="filter-input" style="width: 180px;" value="${currentPickingFilters.orderNumber}">
                            <input type="text" id="recipient-filter" placeholder="수취인 검색" class="filter-input" style="width: 180px;" value="${currentPickingFilters.recipient}">
                        </div>
                    </div>
                    
                    <button id="query-picking-status-btn" class="btn-primary" style="margin-left: auto; white-space: nowrap;">조회하기</button>
                </div>
            </div>
        </div>
        
        <div class="table-scroll-container" style="overflow-y: auto; margin-top: 1rem;">
             <h3 style="font-size: 1.1rem; font-weight: bold; color: #1e293b; margin-bottom: 0.8rem;">출고 진행 현황</h3>
             <div id="picking-status-table-card" style="background: white; border-radius: 8px; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px rgba(0,0,0,0.05); overflow: hidden;">
                <div class="table-wrapper" style="border: none;">
                    <table id="picking-status-table" style="width: 100%; text-align: center;">
                        <thead>
                            <tr>
                                <th style="text-align: center;">출고일자</th>
                                <th style="text-align: center;">차수</th>
                                <th style="text-align: center;">출고지시번호</th>
                                <th style="text-align: center;">출고지 주소</th>
                                <th style="text-align: center;">수취인</th>
                                <th style="text-align: center;">지시수량</th>
                                <th style="text-align: center;">완료수량</th>
                                <th style="text-align: center;">상태</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>
                </div>
            </div>
        </div>
        `;
    
    // 이벤트 리스너 설정
    document.getElementById('query-picking-status-btn').addEventListener('click', loadPickingStatusByDateRange);
    
    document.getElementById('today-btn').addEventListener('click', () => {
        const todayStr = new Date().toISOString().split('T')[0];
        
        const startPicker = document.getElementById('status-start-date-picker');
        const endPicker = document.getElementById('status-end-date-picker');
        
        if (startPicker._flatpickr) startPicker._flatpickr.setDate(todayStr);
        else startPicker.value = todayStr;
        
        if (endPicker._flatpickr) endPicker._flatpickr.setDate(todayStr);
        else endPicker.value = todayStr;
        
        loadPickingStatusByDateRange(); // 클릭시 즉시 조회되도록 연결
    });

    contentArea.querySelectorAll('.filter-input').forEach(input => {
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                loadPickingStatusByDateRange();
            }
        });
    });
    
    // 렌더링 직후 데이터 로드 실행
    await loadTodaySummary();
    await loadPickingStatusByDateRange();
    applyModernCalendar(); // 모던 달력 적용
}

async function loadPickingStatusByDateRange() {
    const tbody = contentArea.querySelector('#picking-status-table tbody');
    if (!tbody) return;

    // 현재 검색조건 값을 필터 유지 객체에 저장
    currentPickingFilters.startDate = contentArea.querySelector('#status-start-date-picker').value;
    currentPickingFilters.endDate = contentArea.querySelector('#status-end-date-picker').value;
    currentPickingFilters.orderNumber = contentArea.querySelector('#order-number-filter').value;
    currentPickingFilters.recipient = contentArea.querySelector('#recipient-filter').value;

    const startDate = currentPickingFilters.startDate;
    const endDate = currentPickingFilters.endDate;

    if (!startDate || !endDate) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;">날짜 범위를 올바르게 선택해주세요.</td></tr>`;
        currentPickingStatusData = [];
        return;
    }

    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding: 2rem;">현황을 불러오는 중...</td></tr>`;

    try {
        const { data: batches, error: batchError } = await supabaseClient
            .from('picking_batches')
            .select('id')
            .eq('channel_id', currentChannelId)
            .gte('batch_date', startDate)
            .lte('batch_date', endDate);

        if (batchError) throw batchError;
        
        if (!batches || batches.length === 0) {
            tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding: 2rem;">해당 기간에 데이터가 없습니다.</td></tr>`;
            currentPickingStatusData = [];
            return;
        }

        const batchIds = batches.map(b => b.id);

        const { data, error } = await supabaseClient
            .from('picking_orders')
            .select(`*, picking_batches (batch_date, batch_number)`)
            .in('batch_id', batchIds)
            .order('id', { ascending: false });

        if (error) throw error;
        
        const allDataInPeriod = data.map(order => ({
            "출고일자": order.picking_batches.batch_date,
            "차수": order.picking_batches.batch_number,
            "출고지시번호": order.order_number,
            "출고지 주소": order.destination_address,
            "수취인": order.recipient,
            "지시수량": order.total_expected_quantity,
            "완료수량": order.total_picked_quantity,
            "상태": order.status
        }));

        const orderFilter = currentPickingFilters.orderNumber.trim().toLowerCase();
        const recipientFilter = currentPickingFilters.recipient.trim().toLowerCase();

        let filteredData = allDataInPeriod;
        if (orderFilter) {
            filteredData = filteredData.filter(order => 
                order['출고지시번호'].toLowerCase().includes(orderFilter)
            );
        }
        if (recipientFilter) {
            filteredData = filteredData.filter(order => 
                order['수취인'] && order['수취인'].toLowerCase().includes(recipientFilter)
            );
        }

        currentPickingStatusData = filteredData;
        
        if (filteredData.length === 0) {
            tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding: 2rem;">조회된 데이터가 없습니다.</td></tr>`;
        } else {
            tbody.innerHTML = filteredData.map(order =>
                `<tr class="${order.상태 === '완료' ? 'completed-row' : ''}">
                    <td style="text-align: center;">${order.출고일자}</td>
                    <td style="text-align: center;">${order.차수}</td>
                    <td style="text-align: center;">${order.출고지시번호}</td>
                    <td style="text-align: center;">${order['출고지 주소'] || ''}</td>
                    <td style="text-align: center;">${order.수취인 || ''}</td>
                    <td style="text-align: center;">${order.지시수량}</td>
                    <td style="text-align: center;">${order.완료수량}</td>
                    <td style="text-align: center;">
                        <span class="user-status ${order.상태 === '완료' ? 'status-approved' : 'status-pending'}">
                            ${order.상태}
                        </span>
                    </td>
                </tr>`
            ).join('');
        }
        
        const tableWrapper = contentArea.querySelector('#picking-status-table-card');
        const tableHeaders = contentArea.querySelectorAll('#picking-status-table thead th');
        
        if (tableWrapper && tableHeaders.length > 0) {
            tableHeaders.forEach(th => {
                th.style.position = 'sticky';
                th.style.top = '0'; 
                th.style.zIndex = '1';
                th.style.backgroundColor = '#f8f9fa';
            });
        }

    } catch (err) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:red; padding: 2rem;">데이터 조회 중 오류가 발생했습니다: ${err.message}</td></tr>`;
        currentPickingStatusData = [];
    }
}


// =================================================================
// [수정1] 상품 마스터 관리 (가운데 정렬 반영 및 1400px 최대 너비 지정)
// =================================================================
async function showProductMaster() {
    contentArea.innerHTML = `
    <div id="products-section" class="content-section active" style="max-width: 1400px; margin: 0 auto; width: 100%;">
        <div class="sticky-controls">
            <div class="page-header">
                <h2>상품 마스터 관리</h2>
                <div class="actions-group">
                    <button id="refresh-product-master-btn" class="btn-secondary">초기화 및 새로고침</button>
                    <button id="download-product-master-btn" class="btn-primary">엑셀 다운로드</button>
                </div>
            </div>
            
            <div class="card" style="margin-top: 1.5rem; margin-bottom: 1.5rem;">
                <div class="card-body" style="display: flex; flex-direction: row; align-items: center; justify-content: space-between; gap: 1.5rem; flex-wrap: wrap;">
                    <div style="display: flex; gap: 0.5rem; align-items: center; flex-wrap: wrap;">
                        <label style="font-weight: 600; font-size: 0.85rem; color: #64748b; white-space: nowrap;">필터 검색:</label>
                        <input type="text" id="filter-prod-code" class="filter-input" placeholder="상품코드 검색" value="${currentFilters.product_code || ''}" style="width: 150px;">
                        <input type="text" id="filter-prod-barcode" class="filter-input" placeholder="바코드 검색" value="${currentFilters.barcode || ''}" style="width: 150px;">
                        <input type="text" id="filter-prod-name" class="filter-input" placeholder="상품명 검색" value="${currentFilters.product_name || ''}" style="width: 150px;">
                    </div>
                    <div style="display: flex; gap: 0.5rem;">
                        <button class="reset-button btn-secondary" style="white-space: nowrap;">초기화</button>
                        <button class="search-button btn-primary" style="white-space: nowrap;">조회하기</button>
                    </div>
                </div>
            </div>
            
            <div class="card" style="margin-bottom: 1.5rem;">
                <div class="card-body" style="display: flex; flex-direction: row; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 1.5rem;">
                     <div style="display: flex; align-items: center; gap: 0.5rem; flex-wrap: wrap;">
                        <label style="font-weight: 600; font-size: 0.85rem; color: #64748b; white-space: nowrap;">데이터 업로드:</label>
                        <button class="download-template btn-secondary" style="white-space: nowrap;">양식 다운로드</button>
                        <div class="custom-file-input" style="width: 190px;">
                            <label>
                                <input type="file" id="upload-file" class="upload-file" accept=".xlsx, .xls">
                                <span class="file-name">선택된 파일 없음</span>
                                <span class="file-btn">파일 선택</span>
                            </label>
                        </div>
                        <button class="upload-data btn-primary" style="white-space: nowrap;">표준 업로드</button>
                        <div style="width: 1px; height: 20px; background-color: #cbd5e1; margin: 0 0.25rem;"></div>
                        <div class="custom-file-input" style="width: 190px;">
                            <label>
                                <input type="file" id="upload-corn-file" class="upload-file" accept=".xlsx, .xls">
                                <span class="file-name">선택된 파일 없음</span>
                                <span class="file-btn">파일 선택</span>
                            </label>
                        </div>
                        <button id="upload-corn-button" class="btn-primary" style="white-space: nowrap;">CORN 업로드</button>
                    </div>
                    <button class="delete-selected btn-danger" style="white-space: nowrap;">선택 항목 삭제</button>
                </div>
            </div>
            
        </div>
        <div class="table-wrapper" style="flex-grow: 1;">
            <div class="table-container"></div>
        </div>
    </div>`;

    const searchButton = contentArea.querySelector('.search-button');
    contentArea.querySelectorAll('.filter-input').forEach(input => {
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                searchButton.click();
            }
        });
    });

    await renderProductTable();
}

async function renderProductTable() {
    const tableContainer = contentArea.querySelector('.table-container');
    if (!tableContainer) return;

    // 필터 값 업데이트 (상태 유지용)
    currentFilters.product_code = contentArea.querySelector('#filter-prod-code') ? contentArea.querySelector('#filter-prod-code').value.trim() : currentFilters.product_code;
    currentFilters.barcode = contentArea.querySelector('#filter-prod-barcode') ? contentArea.querySelector('#filter-prod-barcode').value.trim() : currentFilters.barcode;
    currentFilters.product_name = contentArea.querySelector('#filter-prod-name') ? contentArea.querySelector('#filter-prod-name').value.trim() : currentFilters.product_name;

    tableContainer.innerHTML = '불러오는 중...';
    let query = supabaseClient.from('products').select('*').eq('channel_id', currentChannelId);
    if (currentFilters.product_code) query = query.ilike('product_code', `%${currentFilters.product_code}%`);
    if (currentFilters.barcode) query = query.ilike('barcode', `%${currentFilters.barcode}%`);
    if (currentFilters.product_name) query = query.ilike('product_name', `%${currentFilters.product_name}%`);
    if (currentSort.column) { query = query.order(currentSort.column, { ascending: currentSort.direction === 'asc' }); }
    const { data, error } = await fetchAllWithPagination(query);

    currentProductMasterData = data || [];

    if (error) { tableContainer.innerHTML = `<p>데이터 로딩 오류: ${error.message}</p>`; return; }
    if (data.length === 0) {
        tableContainer.innerHTML = `<p style="text-align:center; padding: 2rem;">표시할 데이터가 없습니다.</p>`;
    } else {
        let tableHTML = `<table style="width: 100%; text-align: center;"><thead><tr><th style="text-align: center;"><input type="checkbox" class="select-all-checkbox"></th><th style="text-align: center;">No.</th><th class="sortable" data-column="product_code" style="text-align: center;">상품코드</th><th class="sortable" data-column="barcode" style="text-align: center;">바코드</th><th class="sortable" data-column="product_name" style="text-align: center;">상품명</th></tr></thead><tbody>`;
        data.forEach((p, index) => { tableHTML += `<tr><td style="text-align: center;"><input type="checkbox" class="row-checkbox" data-id="${p.id}"></td><td style="text-align: center;">${index + 1}</td><td style="text-align: center;">${p.product_code || ''}</td><td style="text-align: center;">${p.barcode}</td><td style="text-align: center;">${p.product_name}</td></tr>`; });
        tableHTML += '</tbody></table>';
        tableContainer.innerHTML = tableHTML;
    }
    updateSortIndicator();
}

function updateSortIndicator() {
    const section = contentArea.querySelector('#products-section');
    if (!section) return;
    section.querySelectorAll('.sortable').forEach(th => {
        th.classList.remove('asc', 'desc');
        if (th.dataset.column === currentSort.column) {
            th.classList.add(currentSort.direction);
        }
    });
}

async function handleCornUpload() {
    const fileInput = contentArea.querySelector('#upload-corn-file');
    if (!fileInput.files[0]) { return alert('CORN 양식 엑셀 파일을 선택하세요.'); }
    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            const dataRows = rows.slice(2);
            if (dataRows.length === 0) throw new Error('엑셀 파일에 데이터가 없습니다 (제목 행 제외).');
            const productsToUpsert = dataRows.map(row => ({
                product_code: row[2],
                product_name: row[3],
                barcode: row[10],
                channel_id: currentChannelId
            })).filter(p => p.barcode && p.product_name);
            if (productsToUpsert.length === 0) { throw new Error('업로드할 유효한 상품 데이터가 없습니다. C, D, K 열에 데이터가 있는지 확인하세요.'); }
            const { error } = await supabaseClient.from('products').upsert(productsToUpsert, { onConflict: 'barcode, channel_id' });
            if (error) throw error;
            alert(`총 ${productsToUpsert.length}개의 상품을 성공적으로 추가/수정했습니다.`);
            renderProductTable();
        } catch (error) {
            alert('CORN 양식 업로드 실패: ' + error.message);
        } finally {
            fileInput.value = '';
            // 파일 초기화 시 UI 텍스트도 원상복구
            const fileNameSpan = fileInput.closest('label').querySelector('.file-name');
            if(fileNameSpan) {
                fileNameSpan.textContent = '선택된 파일 없음';
                fileNameSpan.style.color = 'var(--text-secondary-color)';
                fileNameSpan.style.fontWeight = 'normal';
            }
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

async function handleProductUpload() {
    const fileInput = contentArea.querySelector('#upload-file');
    if (!fileInput.files[0]) { return alert('업로드할 엑셀 파일을 선택하세요.'); }
    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            if (jsonData.length === 0) { throw new Error('엑셀 파일에 데이터가 없습니다.'); }
            const dataWithChannel = jsonData.map(item => ({...item, channel_id: currentChannelId }));
            const { error } = await supabaseClient.from('products').upsert(dataWithChannel, { onConflict: 'barcode, channel_id' });
            if (error) throw error;
            alert('상품 마스터 업로드(추가/수정) 성공!');
            renderProductTable();
        } catch (error) { alert('업로드 실패: ' + error.message); } finally { 
            fileInput.value = ''; 
            // 파일 초기화 시 UI 텍스트도 원상복구
            const fileNameSpan = fileInput.closest('label').querySelector('.file-name');
            if(fileNameSpan) {
                fileNameSpan.textContent = '선택된 파일 없음';
                fileNameSpan.style.color = 'var(--text-secondary-color)';
                fileNameSpan.style.fontWeight = 'normal';
            }
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

// ▼▼▼ [수정] 상품마스터 양식(공양식) 다운로드 시간 추가 및 스타일 적용 ▼▼▼
function handleProductTemplateDownload() {
    const headers = ["product_code", "barcode", "product_name"];
    const worksheet = XLSX.utils.json_to_sheet([], { header: headers });
    applyExcelStyle(worksheet);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Products");
    XLSX.writeFile(workbook, `products_template_${getFormattedDateStr()}.xlsx`);
}
// ▲▲▲ [수정] 끝 ▲▲▲

async function handleDeleteSelectedProducts() {
    const checkedBoxes = contentArea.querySelectorAll('.row-checkbox:checked');
    if (checkedBoxes.length === 0) { return alert('삭제할 항목을 선택하세요.'); }
    if (confirm(`선택된 ${checkedBoxes.length}개의 상품을 삭제하시겠습니까?`)) {
        const idsToDelete = Array.from(checkedBoxes).map(cb => cb.dataset.id);
        const { error } = await supabaseClient.from('products').delete().in('id', idsToDelete);
        if (error) { alert('삭제 실패: ' + error.message); } else { alert('선택한 상품을 삭제했습니다.'); renderProductTable(); }
    }
}

async function fetchAllWithPagination(query, pageSize = 1000) {
    let allData = [];
    let page = 0;
    while (true) {
        const { data, error } = await query.range(page * pageSize, (page + 1) * pageSize - 1);
        if (error) return { data: null, error };
        if (data.length > 0) allData = allData.concat(data);
        if (data.length < pageSize) break;
        page++;
    }
    return { data: allData, error: null };
}

async function showUserManagement() {
    contentArea.innerHTML = `
        <div class="content-section" style="max-width: 900px; margin: 0 auto; width: 100%; padding-top: 1rem;">
            <div class="page-header" style="margin-bottom: 2rem;">
                <h2>사용자 관리</h2>
                <div class="actions-group">
                    <button id="refresh-users-btn" class="btn-secondary">새로고침</button>
                </div>
            </div>
            <div class="card">
                <div class="card-header" style="background-color: #f8fafd; display: flex; justify-content: space-between;">
                    <span>가입된 사용자 목록</span>
                    <span style="font-size: 0.8rem; font-weight: normal; color: #64748b;">승인 대기 중인 사용자를 승인하여 권한을 부여하세요.</span>
                </div>
                <ul id="user-list" class="management-list"></ul>
            </div>
        </div>`;
    await loadUsers();
}

async function loadUsers() {
    const list = contentArea.querySelector('#user-list');
    list.innerHTML = `<li>불러오는 중...</li>`;
    const { data, error } = await supabaseClient.rpc('list_all_users');
    if (error) { list.innerHTML = `<li>사용자 목록 로딩 오류: ${error.message}</li>`; return; }
    if (!data || data.length === 0) { list.innerHTML = `<li style="padding:2rem; text-align:center; color:#64748b;">가입한 사용자가 없습니다.</li>`; return; }

    list.innerHTML = data.map(user => {
        const isAdmin = user.user_metadata && user.user_metadata.is_admin === true;
        const isSuperAdmin = user.email === 'eowert72@gmail.com';
        
        const isApproved = isAdmin || isSuperAdmin;

        const statusClass = isApproved ? 'status-approved' : 'status-pending';
        const statusText = isApproved ? '승인 완료' : '승인 대기';
        
        let actionButton = '';
        if (!isAdmin && !isSuperAdmin) {
            actionButton = `<button class="btn-approve approve-user-button" data-id="${user.id}">승인하기</button>`;
        }

        return `<li class="management-list-item" style="padding: 1.2rem 1.5rem;">
                    <span style="font-weight: 500;">${user.email}</span>
                    <div style="display: flex; align-items: center; gap: 0.5rem;">
                        <span class="user-status ${statusClass}">${statusText}</span>
                        ${actionButton}
                    </div>
                </li>`;
    }).join('');
}


async function handleApproveUser(e) {
    const userId = e.target.dataset.id;
    if (confirm('이 사용자를 승인하고 관리자 권한을 부여하시겠습니까?')) {
        const { data, error } = await supabaseClient.rpc('approve_and_grant_admin', {
            user_id_to_approve: userId
        });

        if (error) {
            alert('승인 실패: ' + error.message);
        } else {
            alert(data || '사용자가 승인되었습니다.');
            showUserManagement();
        }
    }
}

logoutButton.addEventListener('click', async () => {
    if (confirm('로그아웃하시겠습니까?')) {
        await supabaseClient.auth.signOut();
        window.location.href = 'login.html';
    }
});
