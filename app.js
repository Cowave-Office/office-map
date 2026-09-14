// 파일명을 다시 원본으로 롤백 (_2.xlsx 제거, _4.xlsx도 서버에 올리실땐 원본이름으로 변경 권장)
        const offices = {
            'dk17': { name: '현대테라타워 가산DK 17층 (CC사업부)', file: 'dk17.xlsx', data: [] },
            'dk18': { name: '현대테라타워 가산DK 18층 (CC사업부)', file: 'dk18.xlsx', data: [] },
            'daerung': { name: '대륭포스트타워6차 17층 (HQ, CC CTO, CISO)', file: 'daerung.xlsx', data: [] },
            'gplus': { name: '지플러스타워 14층 (플레이오토)', file: 'gplus.xlsx', data: [] },
            'woorim': { name: '우림라이온스밸리 A동 14층 (메이크샵, 몰테일)', file: 'woorim.xlsx', data: [] }
        };
        let currentTab = document.body.dataset.office || 'dk17';
        
        // 검색 관리를 위한 글로벌 상태
        let currentSearchQuery = '';
        let globalMatches = [];
        let currentMatchIdx = 0;

        let directoryScope = null;
        const normalizeText = OfficeMapSearch.normalize;
        const escapeHtml = value => String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
        const searchIndex = new Map();
        const indexedData = new Map();
        function ensureSearchIndex(tab) {
            if (indexedData.get(tab) !== offices[tab].data) {
                searchIndex.set(tab, OfficeMapSearch.buildIndex(offices[tab].data));
                indexedData.set(tab, offices[tab].data);
            }
        }

        document.addEventListener('DOMContentLoaded', () => {
            let zoomLevel = 1.0;
            let tableObserver;

            const viewport = document.getElementById('viewport');
            const floorplan = document.getElementById('floorplanContainer');
            const dynamicFloorplan = document.getElementById('dynamicFloorplan');
            const searchInput = document.getElementById('searchInput');
            const searchBox = document.getElementById('searchBox');
            const searchPopup = document.getElementById('searchPopup');
            const searchResults = document.getElementById('searchResults');
            const searchSummary = document.getElementById('searchSummary');
            let searchTimer = 0;
            let composing = false;
            let activeOption = -1;
            const clearSearchBtn = document.getElementById('clearSearchBtn');
            const excelInput = document.getElementById('excelInput');
            const downloadExcelBtn = document.getElementById('downloadExcelBtn');
            const toggleDirBtn = document.getElementById('toggleDirBtn');
            const dirPanel = document.getElementById('dirPanel');
            const adminLoginBtn = document.getElementById('adminLoginBtn');
            const adminControls = document.getElementById('adminControls');
            const changePwBtn = document.getElementById('changePwBtn');
            const printScaleInput = document.getElementById('printScaleInput');
            const camera = new OfficeMapCamera(viewport, floorplan, zoom => {
                zoomLevel = zoom;
                document.getElementById('zoomVal').textContent = Math.round(zoom * 100) + '%';
            });
            document.getElementById('dirContent').addEventListener('click', event => {
                const button = event.target.closest('[data-directory-value]');
                if (!button) return;
                directoryScope = currentTab;
                currentSearchQuery = '';
                searchInput.value = button.dataset.directoryValue;
                toggleDirBtn.click();
                performSearch({ navigate: true });
            });
            const fullscreenBtn = document.getElementById('fullscreenBtn');
            fullscreenBtn.addEventListener('click', async () => {
                try {
                    if (document.fullscreenElement) await document.exitFullscreen();
                    else await document.documentElement.requestFullscreen();
                } catch {
                    alert('이 브라우저에서는 전체화면 전환을 지원하지 않습니다.');
                }
            });
            document.addEventListener('fullscreenchange', () => {
                const active = Boolean(document.fullscreenElement);
                fullscreenBtn.setAttribute('aria-pressed', String(active));
                fullscreenBtn.title = active ? '전체화면 종료' : '전체화면';
                fullscreenBtn.setAttribute('aria-label', fullscreenBtn.title);
                fullscreenBtn.querySelector('i').className = active ? 'fa-solid fa-compress' : 'fa-solid fa-expand';
            });

            // --- 탭 전환 로직 ---
            const tabButtons = document.querySelectorAll('.tab-btn');
            tabButtons.forEach(btn => {
                btn.addEventListener('click', (e) => { 
                    clearTimeout(searchTimer);
                    closeSearch();
                    if (searchInput.value.trim() !== '') {
                        searchInput.value = '';
                        clearSearchBtn.classList.add('hidden');
                        currentSearchQuery = '';
                        globalMatches = [];
                    }
                    directoryScope = null;
                    switchTab(e.currentTarget.dataset.tab); 
                });
            });

            function switchTab(tabKey) {
                if (!offices[tabKey]) return;
                currentTab = tabKey;
                tabButtons.forEach(btn => {
                    btn.setAttribute('aria-pressed', String(btn.dataset.tab === tabKey));
                    if (btn.dataset.tab === tabKey) {
                        btn.classList.add('bg-[#4F81BD]', 'text-white', 'font-bold', 'shadow-md');
                        btn.classList.remove('text-stone-500', 'hover:text-stone-700', 'hover:bg-stone-200', 'font-medium');
                    } else {
                        btn.classList.remove('bg-[#4F81BD]', 'text-white', 'font-bold', 'shadow-md');
                        btn.classList.add('text-stone-500', 'hover:text-stone-700', 'hover:bg-stone-200', 'font-medium');
                    }
                });
                const activeTab = document.querySelector(`.tab-btn[data-tab="${tabKey}"]`);
                const nav = activeTab.parentElement;
                const activeRect = activeTab.getBoundingClientRect(), navRect = nav.getBoundingClientRect();
                nav.scrollLeft += activeRect.left - navRect.left - (nav.clientWidth - activeRect.width) / 2;
                document.getElementById('tabTitle').innerHTML = `<span class="w-3.5 h-3.5 bg-[#4F81BD] rounded-full inline-block"></span> ${offices[tabKey].name}`;
                renderDatasetToTable(offices[tabKey].data);
                camera.fit();
            }

            // --- 관리자 설정 ---
            if (!localStorage.getItem('adminPassword_OfficeMap')) {
                localStorage.setItem('adminPassword_OfficeMap', 'cowave8888');
            }

            adminLoginBtn.addEventListener('click', () => {
                if (!adminControls.classList.contains('hidden')) {
                    adminControls.classList.add('hidden');
                    return;
                }
                const currentPw = localStorage.getItem('adminPassword_OfficeMap');
                const inputPw = prompt('관리자 비밀번호를 입력하세요.');
                if (inputPw === currentPw) adminControls.classList.remove('hidden');
                else if (inputPw !== null) alert('비밀번호가 일치하지 않습니다.');
            });

            changePwBtn.addEventListener('click', () => {
                const currentPw = localStorage.getItem('adminPassword_OfficeMap');
                const inputPw = prompt('현재 비밀번호를 입력하세요.');
                if (inputPw === currentPw) {
                    const newPw = prompt('새로운 비밀번호를 입력하세요.');
                    if (newPw) {
                        localStorage.setItem('adminPassword_OfficeMap', newPw);
                        alert('비밀번호가 성공적으로 변경되었습니다.');
                    }
                } else if (inputPw !== null) {
                    alert('현재 비밀번호가 일치하지 않습니다.');
                }
            });

            document.getElementById('zoomInBtn').addEventListener('click', () => camera.setZoom(camera.zoom + 0.1));
            document.getElementById('zoomOutBtn').addEventListener('click', () => camera.setZoom(camera.zoom - 0.1));
            document.getElementById('zoomResetBtn').addEventListener('click', () => camera.fit());

            // --- 글씨 크기 조절 ---
            let currentFontSize = 10;
            document.getElementById('fontIncreaseBtn').addEventListener('click', () => {
                if (currentFontSize < 24) { currentFontSize += 1; document.documentElement.style.setProperty('--seat-font-size', currentFontSize + 'px'); }
            });
            document.getElementById('fontDecreaseBtn').addEventListener('click', () => {
                if (currentFontSize > 6) { currentFontSize -= 1; document.documentElement.style.setProperty('--seat-font-size', currentFontSize + 'px'); }
            });
            document.getElementById('fontResetBtn').addEventListener('click', () => {
                currentFontSize = 10; document.documentElement.style.setProperty('--seat-font-size', '10px');
            });

            toggleDirBtn.addEventListener('click', function() {
                closeSearch();
                dirPanel.classList.toggle('hidden');
                this.setAttribute('aria-expanded', String(!dirPanel.classList.contains('hidden')));
                const icon = this.querySelector('i');
                if (dirPanel.classList.contains('hidden')) {
                    icon.className = 'fa-solid fa-list';
                    this.className = 'h-[38px] px-4 bg-white hover:bg-stone-100 text-[#4F81BD] border border-[#4F81BD] rounded-lg text-sm font-bold transition flex items-center gap-2 shadow-sm whitespace-nowrap';
                } else {
                    icon.className = 'fa-solid fa-chevron-up';
                    this.className = 'h-[38px] px-4 bg-[#4F81BD] text-white border border-[#4F81BD] rounded-lg text-sm font-bold transition flex items-center gap-2 shadow-sm whitespace-nowrap';
                }
            });

            // --- 데이터 렌더링 및 디렉토리 추출 로직 ---
            function updateDirectory(items) {
                const teams = new Set(); const rooms = new Set();
                items.forEach(it => {
                    if (!it.v) return; const val = it.v.trim();
                    if (val === '' || val.includes('출입문') || val === '(예비 공석)' || val.includes('외부인력') || val.includes('망분리 PC') || val.includes('택배PC') || val.includes('우편') || val.includes('공용창고') || val.includes('STORAGE')) return;
                    const vUp = val.toUpperCase();
                    
                    // 회의실 분류 (임원실/CEO 제외, 1인회의실, OA존 포함)
                    // OA 또는 OA존 키워드 추가
                    const isRoom = val.includes('화장실') || val.includes('회의') || val.includes('포커스') || vUp.includes('CONFERENCE') || vUp.includes('A80') || vUp.includes('A1201') || vUp.includes('B80') || vUp.includes('B1201') || val.includes('라운지') || vUp.includes('LOUNGE') || vUp.includes('서버실') || vUp.includes('SERVER') || vUp.includes('TEST') || val.includes('ROOM') || vUp.includes('OA') || val.includes('탕비') || val.includes('메일룸');
                    
                    // 팀 및 부서 분류 (임원실/CEO 포함, 1407호 포함, 지플러스 셀/그룹/담당/본부/실/부 등 포함)
                    // OA존은 회의실로 가야하므로 isTitle에서 OA 제외
                    const isTitle = (!vUp.includes('OA') && !val.includes('회의') && !val.includes('화장실') && !val.includes('서버') && !val.includes('탕비') && (val.includes('임원') || vUp.includes('CEO') || val.includes('부서') || val.includes('팀') || val.includes('본부') || val.includes('실') || val.includes('파트') || val.includes('사업') || val.includes('셀') || val.includes('그룹') || val.includes('담당') || val.includes('센터') || vUp.includes('WORK STATION') || val.includes('구역') || val.includes('존') || vUp.includes('ZONE'))) && !val.includes('전산파트');
                    
                    const cleanVal = val.replace(/\n/g, ' ').trim();
                    if (isTitle) teams.add(cleanVal); else if (isRoom && !isTitle) rooms.add(cleanVal); 
                });

                const buildSection = (title, icon, set) => {
                    if (set.size === 0) return '';
                    const arr = Array.from(set).sort();
                    let html = `<div class="mb-4"><h3 class="text-sm font-extrabold text-stone-700 mb-3 flex items-center gap-1.5 border-b border-stone-300 pb-2"><i class="${icon}"></i> ${title} <span class="text-xs text-stone-400 font-normal ml-1">${arr.length}</span></h3><div class="flex flex-wrap gap-2">`;
                    arr.forEach(item => { html += `<button data-directory-value="${escapeHtml(item)}" class="px-2.5 py-1.5 bg-white border border-stone-300 hover:border-[#4F81BD] hover:text-[#4F81BD] hover:bg-blue-50 rounded-md text-xs font-bold text-stone-600 transition shadow-sm">${escapeHtml(item)}</button>`; });
                    html += `</div></div>`; return html;
                };

                document.getElementById('dirContent').innerHTML = buildSection('부서 및 팀', 'fa-solid fa-users-viewfinder', teams) + buildSection('회의실 및 시설', 'fa-solid fa-door-closed', rooms);
            }

            function renderDatasetToTable(items) {
                dynamicFloorplan.innerHTML = '';
                if (!items || items.length === 0) {
                    dynamicFloorplan.innerHTML = '<div class="text-stone-500 font-bold p-10">데이터를 불러오는 중이거나 데이터가 없습니다.</div>';
                    return;
                }

                updateDirectory(items);

                let maxR = 1, maxC = 1; 
                items.forEach(it => {
                    if (it.r + it.rs - 1 > maxR) maxR = it.r + it.rs - 1;
                    if (it.c + it.cs - 1 > maxC) maxC = it.c + it.cs - 1;
                });
                let minR = 1, minC = 1; 

                const cellMap = {}; const covered = new Set();
                items.forEach(it => {
                    cellMap[`${it.r},${it.c}`] = it;
                    if (it.rs > 1 || it.cs > 1) {
                        for (let r = it.r; r < it.r + it.rs; r++) {
                            for (let c = it.c; c < it.c + it.cs; c++) {
                                if (!(r === it.r && c === it.c)) covered.add(`${r},${c}`);
                            }
                        }
                    }
                });

                const tableContainer = document.createElement('div');
                tableContainer.style.position = 'relative';

                const borderCells = [];
                const visibleItems = [];
                const table = document.createElement('table');
                table.className = 'excel-table-layout';
                table.style.width = ((maxC - minC + 1) * 32) + 'px';

                const colgroup = document.createElement('colgroup');
                for (let c = minC; c <= maxC; c++) {
                    const col = document.createElement('col');
                    col.style.width = '32px'; col.style.minWidth = '32px'; col.style.maxWidth = '32px';
                    colgroup.appendChild(col);
                }
                table.appendChild(colgroup);

                for (let r = minR; r <= maxR; r++) {
                    const tr = document.createElement('tr');
                    tr.style.height = '26px'; 

                    for (let c = minC; c <= maxC; c++) {
                        if (covered.has(`${r},${c}`)) continue;

                        const item = cellMap[`${r},${c}`];
                        const td = document.createElement('td');

                        if (item) {
                            // DOM 검색을 위한 데이터셋 삽입
                            visibleItems.push(item);
                            td.dataset.r = item.r;
                            td.dataset.c = item.c;

                            if (item.rs > 1) td.rowSpan = item.rs;
                            if (item.cs > 1) td.colSpan = item.cs;

                            const valTrimmed = (item.v || '').trim();
                            const vUp = valTrimmed.toUpperCase();

                            const wrapperDiv = document.createElement('div');
                            wrapperDiv.style.width = '100%'; wrapperDiv.style.height = '100%';
                            wrapperDiv.style.display = 'flex'; wrapperDiv.style.alignItems = 'center';
                            wrapperDiv.style.justifyContent = 'center'; wrapperDiv.style.flexDirection = 'column';

                            // 가로 한줄 예외 처리 (팀, 부서, 센터, 셀, 그룹, 1407호 등 추가)
                            const isHorizontalForce = valTrimmed.includes('임원') || vUp.includes('CEO') || valTrimmed.includes('부서') || valTrimmed.includes('팀') || valTrimmed.includes('본부') || valTrimmed.includes('실') || valTrimmed.includes('센터') || valTrimmed.includes('셀') || valTrimmed.includes('그룹') || valTrimmed.includes('담당') || vUp.includes('ZONE') || valTrimmed.includes('파트') || valTrimmed.includes('라운지') || valTrimmed.includes('공용창고') || valTrimmed.includes('회의');

                            if (isHorizontalForce) {
                                wrapperDiv.style.whiteSpace = 'nowrap';
                                wrapperDiv.innerHTML = `<span>${escapeHtml(valTrimmed.replace(/\n/g, ' '))}</span>`;
                            } else {
                                const isVertical = (item.align && (item.align.textRotation === '255' || item.align.textRotation === '90')) || (item.rs > item.cs && valTrimmed.length >= 2 && !valTrimmed.includes('출입문') && !vUp.includes('PC')) || valTrimmed.includes('김성실');
                                
                                if (isVertical && valTrimmed.length > 0) {
                                    const parts = valTrimmed.split('\n');
                                    let htmlContent = '';
                                    if (parts.length > 0) htmlContent += `<div style="writing-mode: vertical-rl; text-orientation: upright; letter-spacing: 0;">${escapeHtml(parts[0])}</div>`;
                                    if (parts.length > 1) htmlContent += `<div style="margin-top: 4px; font-weight: 900; font-size: 0.95em;">${escapeHtml(parts.slice(1).join(' '))}</div>`;
                                    wrapperDiv.innerHTML = htmlContent;
                                } else {
                                    wrapperDiv.innerHTML = escapeHtml(valTrimmed).replace(/\n/g, '<br>');
                                }
                            }
                            
                            td.appendChild(wrapperDiv);
                            
                            if (valTrimmed !== '' && !isHorizontalForce && !valTrimmed.includes('화장실') && !valTrimmed.includes('서버룸')) {
                                td.classList.add('seat-td');
                            }
                            
                            td.style.backgroundColor = item.bg || 'transparent';
                            td.style.border = 'none';

                            if (item.border && Object.keys(item.border).length) borderCells.push({ td, item });

                            td.style.wordBreak = 'break-all';
                            td.style.color = '#33302a';

                            if (item.bg && item.bg !== 'transparent') {
                                const hex = item.bg.replace('#', '');
                                if (hex.length === 6) {
                                    const rv = parseInt(hex.substr(0, 2), 16);
                                    const gv = parseInt(hex.substr(2, 2), 16);
                                    const bv = parseInt(hex.substr(4, 2), 16);
                                    const yiq = ((rv * 299) + (gv * 587) + (bv * 114)) / 1000;
                                    if (yiq < 128) td.style.color = '#ffffff';
                                }
                            }
                        }
                        tr.appendChild(td);
                    }
                    table.appendChild(tr);
                }

                tableContainer.appendChild(table);
                dynamicFloorplan.appendChild(tableContainer);
                ensureSearchIndex(currentTab);
                drawBorders(tableContainer, table, borderCells);
                tableObserver?.disconnect();
                let size = `${table.offsetWidth},${table.offsetHeight}`;
                tableObserver = new ResizeObserver(() => {
                    const nextSize = `${table.offsetWidth},${table.offsetHeight}`;
                    if (nextSize === size) return;
                    size = nextSize;
                    drawBorders(tableContainer, table, borderCells);
                    camera.resize();
                });
                tableObserver.observe(table);
            }


            function drawBorders(container, table, cells) {
                container.querySelector('.map-borders')?.remove();
                const ns = 'http://www.w3.org/2000/svg';
                const svg = document.createElementNS(ns, 'svg');
                svg.classList.add('map-borders');
                svg.setAttribute('aria-hidden', 'true');
                const width = table.offsetWidth, height = table.offsetHeight;
                svg.setAttribute('width', width);
                svg.setAttribute('height', height);
                svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
                // Read layout once per render, before appending any line nodes.
                const tableRect = table.getBoundingClientRect();
                const scale = tableRect.width / width || 1;
                const segments = new Map();
                cells.forEach(({ td, item }) => {
                    const rect = td.getBoundingClientRect();
                    const x = (rect.left - tableRect.left) / scale, y = (rect.top - tableRect.top) / scale;
                    const w = rect.width / scale, h = rect.height / scale;
                    const edges = { top: [x,y,x+w,y], bottom: [x,y+h,x+w,y+h], left: [x,y,x,y+h], right: [x+w,y,x+w,y+h] };
                    Object.entries(item.border).forEach(([side, border]) => {
                        const style = String(border.style || '').toLowerCase();
                        if (!edges[side] || !style || style === 'none') return;
                        const points = edges[side];
                        const weight = style === 'thick' || style === 'double' ? 3 : style.includes('medium') ? 2 : 1;
                        const key = points.map(value => value.toFixed(2)).join(',');
                        const previous = segments.get(key);
                        if (!previous || previous.weight <= weight) segments.set(key, { points, weight, style, color: '#' + String(border.color || '000000').replace(/^#/, '').slice(-6) });
                    });
                });
                segments.forEach(({ points, weight, style, color }) => {
                    const addLine = (offset = 0, lineWidth = weight) => {
                        const line = document.createElementNS(ns, 'line');
                        const horizontal = points[1] === points[3];
                        ['x1','y1','x2','y2'].forEach((attr, i) => line.setAttribute(attr, points[i] + ((horizontal ? i % 2 === 1 : i % 2 === 0) ? offset : 0)));
                        line.setAttribute('stroke', color);
                        line.style.setProperty('--line-width', lineWidth + 'px');
                        if (style.includes('dash')) line.setAttribute('stroke-dasharray', '6 3');
                        else if (style.includes('dot')) line.setAttribute('stroke-dasharray', '1 2');
                        svg.appendChild(line);
                    };
                    if (style === 'double') { addLine(-1, 1); addLine(1, 1); }
                    else addLine();
                });
                container.appendChild(svg);
            }

            // --- 엑셀 파싱 로직 ---
            function applyTint(hex, tint) {
                if (!hex) return hex;
                hex = hex.replace('#', '');
                if (hex.length === 8) hex = hex.substring(2); 
                let r = parseInt(hex.substring(0, 2), 16);
                let g = parseInt(hex.substring(2, 4), 16);
                let b = parseInt(hex.substring(4, 6), 16);

                if (tint > 0) {
                    r = Math.round(r * (1 - tint) + (255 * tint));
                    g = Math.round(g * (1 - tint) + (255 * tint));
                    b = Math.round(b * (1 - tint) + (255 * tint));
                } else if (tint < 0) {
                    r = Math.round(r * (1 + tint));
                    g = Math.round(g * (1 + tint));
                    b = Math.round(b * (1 + tint));
                }
                const toHex = (n) => { const h = Math.max(0, Math.min(255, n)).toString(16); return h.length === 1 ? '0' + h : h; };
                return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
            }

            async function parseXlsxDirectXML(fileArrayBuffer) {
                const zip = await JSZip.loadAsync(fileArrayBuffer);
                const stylesXmlText = await zip.file('xl/styles.xml').async('string');
                const parser = new DOMParser();
                const stylesDoc = parser.parseFromString(stylesXmlText, 'text/xml');

                const themeColors = {
                    0: "#FFFFFF", 1: "#000000", 2: "#E7E6E6", 3: "#44546A",
                    4: "#5B9BD5", 5: "#ED7D31", 6: "#A5A5A5", 7: "#FFC000",
                    8: "#4472C4", 9: "#70AD47", 10: "#FFFF00", 11: "#FF0000"
                };

                const fills = [];
                stylesDoc.querySelectorAll('fills > fill').forEach(f => {
                    const fg = f.querySelector('fgColor');
                    if (fg) {
                        const rgb = fg.getAttribute('rgb'); const theme = fg.getAttribute('theme'); const tint = parseFloat(fg.getAttribute('tint') || '0');
                        let baseColor = "";
                        if (rgb) baseColor = "#" + (rgb.length === 8 ? rgb.substring(2) : rgb);
                        else if (theme !== null) baseColor = themeColors[theme] || "#E2E8F0";
                        if (baseColor && tint !== 0) baseColor = applyTint(baseColor, tint);
                        fills.push(baseColor);
                    } else fills.push("");
                });

                const borders = [];
                stylesDoc.querySelectorAll('borders > border').forEach(b => {
                    const borderObj = {};
                    ['top', 'bottom', 'left', 'right'].forEach(side => {
                        const sideNode = b.querySelector(side);
                        if (sideNode && sideNode.getAttribute('style')) {
                            const colorNode = sideNode.querySelector('color');
                            let color = "FF000000";
                            if (colorNode) {
                                const rgb = colorNode.getAttribute('rgb'); const theme = colorNode.getAttribute('theme'); const tint = parseFloat(colorNode.getAttribute('tint') || '0');
                                if (rgb) color = rgb;
                                else if (theme !== null) {
                                    let baseC = themeColors[theme] || "#000000";
                                    if (tint !== 0) baseC = applyTint(baseC, tint);
                                    color = "FF" + baseC.replace('#', '');
                                }
                            }
                            borderObj[side] = { style: sideNode.getAttribute('style'), color: color };
                        }
                    });
                    borders.push(borderObj);
                });

                const xfFills = []; const xfBorders = []; const xfAligns = [];
                stylesDoc.querySelectorAll('cellXfs > xf').forEach(xf => {
                    const fillId = parseInt(xf.getAttribute('fillId') || '0', 10);
                    const borderId = parseInt(xf.getAttribute('borderId') || '0', 10);
                    xfFills.push(fills[fillId] || ""); xfBorders.push(borders[borderId] || {});
                    
                    const aNode = xf.querySelector('alignment');
                    if (aNode) xfAligns.push({ wrapText: aNode.getAttribute('wrapText'), textRotation: aNode.getAttribute('textRotation') });
                    else xfAligns.push({});
                });

                const sheetXmlText = await zip.file('xl/worksheets/sheet1.xml').async('string');
                const sheetDoc = parser.parseFromString(sheetXmlText, 'text/xml');

                let sharedStrings = [];
                const ssFile = zip.file('xl/sharedStrings.xml');
                if (ssFile) {
                    const ssXmlText = await ssFile.async('string');
                    const ssDoc = parser.parseFromString(ssXmlText, 'text/xml');
                    ssDoc.querySelectorAll('si').forEach(si => sharedStrings.push(si.textContent.trim()));
                }

                const mergeMap = {};
                sheetDoc.querySelectorAll('mergeCells > mergeCell').forEach(m => {
                    const ref = m.getAttribute('ref'); const parts = ref.split(':');
                    const start = XLSX.utils.decode_cell(parts[0]); const end = XLSX.utils.decode_cell(parts[1]);
                    mergeMap[`${start.r + 1},${start.c + 1}`] = { rs: end.r - start.r + 1, cs: end.c - start.c + 1, endR: end.r + 1, endC: end.c + 1 };
                });

                const rawCells = {};
                sheetDoc.querySelectorAll('sheetData > row > c').forEach(c => {
                    const cellAddr = c.getAttribute('r'); const decoded = XLSX.utils.decode_cell(cellAddr);
                    const row = decoded.r + 1; const col = decoded.c + 1;
                    const styleIdx = parseInt(c.getAttribute('s') || '0', 10);
                    const bg = xfFills[styleIdx] || ""; const border = xfBorders[styleIdx] || {}; const align = xfAligns[styleIdx] || {};
                    const type = c.getAttribute('t'); const vNode = c.querySelector('v');
                    let val = "";
                    if (vNode) {
                        if (type === 's') val = sharedStrings[parseInt(vNode.textContent, 10)] || "";
                        else val = vNode.textContent;
                    } else {
                        const isNode = c.querySelector('is > t');
                        if (isNode) val = isNode.textContent;
                    }
                    rawCells[`${row},${col}`] = { r: row, c: col, v: val, bg: bg, border: border, align: align };
                });

                const parsedItems = [];
                Object.keys(rawCells).forEach(key => {
                    const cell = rawCells[key];
                    const span = mergeMap[`${cell.r},${cell.c}`] || { rs: 1, cs: 1 };
                    let mergedBorder = JSON.parse(JSON.stringify(cell.border));

                    if (span.rs > 1 || span.cs > 1) {
                        const endR = cell.r + span.rs - 1; const endC = cell.c + span.cs - 1;
                        const bottomCell = rawCells[`${endR},${cell.c}`];
                        if (bottomCell && bottomCell.border && bottomCell.border.bottom) mergedBorder.bottom = bottomCell.border.bottom;
                        const rightCell = rawCells[`${cell.r},${endC}`];
                        if (rightCell && rightCell.border && rightCell.border.right) mergedBorder.right = rightCell.border.right;
                        const cornerCell = rawCells[`${endR},${endC}`];
                        if (cornerCell && cornerCell.border) {
                            if (cornerCell.border.bottom) mergedBorder.bottom = cornerCell.border.bottom;
                            if (cornerCell.border.right) mergedBorder.right = cornerCell.border.right;
                        }
                    }

                    if (cell.v || cell.bg || span.rs > 1 || span.cs > 1 || Object.keys(mergedBorder).length > 0) {
                        parsedItems.push({ r: cell.r, c: cell.c, v: cell.v, bg: cell.bg, rs: span.rs, cs: span.cs, border: mergedBorder, align: cell.align, isExcel: true });
                    }
                });

                return parsedItems;
            }

            excelInput.addEventListener('change', async (e) => {
                const file = e.target.files[0];
                if (!file) return;

                try {
                    const arrayBuffer = await file.arrayBuffer();
                    const parsedItems = await parseXlsxDirectXML(arrayBuffer);
                    if (parsedItems && parsedItems.length > 0) {
                        offices[currentTab].data = parsedItems;
                        localStorage.setItem('savedFloorplanData_' + currentTab, JSON.stringify(parsedItems));
                        renderDatasetToTable(parsedItems);
                        camera.fit();
                        closeSearch();
                        globalMatches = [];
                        alert(`현재 탭(${offices[currentTab].name})에 업로드가 완료되었습니다.`);
                    } else {
                        alert("엑셀 도면 데이터를 읽어오지 못했습니다.");
                    }
                } catch (err) {
                    console.error(err);
                    alert('엑셀 분석 중 오류가 발생했습니다.');
                }
            });

            downloadExcelBtn.addEventListener('click', () => {
                const activeData = offices[currentTab].data;
                if (!activeData || activeData.length === 0) return alert("다운로드할 데이터가 없습니다.");

                const wb = XLSX.utils.book_new();
                const ws = {}; const merges = [];

                let maxR = 75, maxC = 100;
                activeData.forEach(item => {
                    if (item.r + item.rs > maxR) maxR = item.r + item.rs;
                    if (item.c + item.cs > maxC) maxC = item.c + item.cs;
                });

                activeData.forEach(item => {
                    const r = item.r - 1; const c = item.c - 1;
                    if (item.rs > 1 || item.cs > 1) {
                        merges.push({ s: { r: r, c: c }, e: { r: r + item.rs - 1, c: c + item.cs - 1 } });
                    }
                });

                activeData.forEach(item => {
                    const startR = item.r - 1; const startC = item.c - 1;
                    const endR = startR + item.rs - 1; const endC = startC + item.cs - 1;

                    for (let currR = startR; currR <= endR; currR++) {
                        for (let currC = startC; currC <= endC; currC++) {
                            const cellRef = XLSX.utils.encode_cell({r: currR, c: currC});
                            let cellStyle = {
                                font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "33302A" } },
                                alignment: { vertical: "center", horizontal: "center", wrapText: true }
                            };

                            let textValue = "";
                            if (currR === startR && currC === startC) {
                                textValue = item.v || "";
                                if (item.align && (item.align.textRotation === '255' || item.align.textRotation === '90')) {
                                    cellStyle.alignment.textRotation = 255;
                                }
                            }

                            if (item.bg && item.bg.trim() !== "") {
                                cellStyle.fill = { fgColor: { rgb: item.bg.replace("#", "") } };
                                const hex = item.bg.replace('#', '');
                                if (hex.length === 6) {
                                    const rv = parseInt(hex.substr(0, 2), 16); const gv = parseInt(hex.substr(2, 2), 16); const bv = parseInt(hex.substr(4, 2), 16);
                                    if (((rv * 299) + (gv * 587) + (bv * 114)) / 1000 < 128) cellStyle.font.color = { rgb: "FFFFFF" };
                                }
                            }

                            if (item.border && Object.keys(item.border).length > 0) {
                                cellStyle.border = {};
                                if (currR === startR && item.border.top) cellStyle.border.top = { style: item.border.top.style, color: { rgb: item.border.top.color.replace(/^FF/, '').replace(/^#/, '') || "000000" } };
                                if (currR === endR && item.border.bottom) cellStyle.border.bottom = { style: item.border.bottom.style, color: { rgb: item.border.bottom.color.replace(/^FF/, '').replace(/^#/, '') || "000000" } };
                                if (currC === startC && item.border.left) cellStyle.border.left = { style: item.border.left.style, color: { rgb: item.border.left.color.replace(/^FF/, '').replace(/^#/, '') || "000000" } };
                                if (currC === endC && item.border.right) cellStyle.border.right = { style: item.border.right.style, color: { rgb: item.border.right.color.replace(/^FF/, '').replace(/^#/, '') || "000000" } };
                            }
                            if (!ws[cellRef]) ws[cellRef] = { v: textValue, t: "s", s: cellStyle };
                            else { ws[cellRef].v = textValue; ws[cellRef].s = cellStyle; }
                        }
                    }
                });

                ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR - 1, c: maxC - 1 } });
                ws['!merges'] = merges; ws['!cols'] = Array(maxC).fill({ wpx: 32 }); ws['!rows'] = Array(maxR).fill({ hpx: 26 }); 
                XLSX.utils.book_append_sheet(wb, ws, currentTab);
                XLSX.writeFile(wb, `${offices[currentTab].name}.xlsx`);
            });

            function closeSearch() {
                searchPopup.hidden = true;
                activeOption = -1;
                searchInput.setAttribute('aria-expanded', 'false');
                searchInput.removeAttribute('aria-activedescendant');
            }

            function positionSearch() {
                if (searchPopup.hidden) return;
                const field = searchBox.getBoundingClientRect();
                const toolbar = document.getElementById('secondaryToolbar').getBoundingClientRect();
                const width = Math.min(Math.max(field.width, 420), toolbar.width - 24);
                const left = Math.max(12, Math.min(field.left - toolbar.left, toolbar.width - width - 12));
                const bottom = window.visualViewport ? window.visualViewport.offsetTop + window.visualViewport.height : window.innerHeight;
                searchPopup.style.left = left + 'px';
                searchPopup.style.top = (field.bottom - toolbar.top + 6) + 'px';
                searchPopup.style.width = width + 'px';
                searchPopup.style.setProperty('--search-height', Math.max(60, Math.min(360, bottom - field.bottom - 18)) + 'px');
            }

            function seatAddress(item) {
                let col = item.c, letters = '';
                while (col > 0) { col--; letters = String.fromCharCode(65 + col % 26) + letters; col = Math.floor(col / 26); }
                return letters + item.r;
            }

            function renderSearch() {
                activeOption = -1;
                searchInput.removeAttribute('aria-activedescendant');
                searchResults.innerHTML = globalMatches.map((match, index) => {
                    const team = match.isPerson ? match.teamName || '팀 확인 필요' : match.isOrganization ? '부서' : '시설 / 구역';
                    const office = offices[match.tab].name;
                    return `<button type="button" role="option" id="search-result-${index}" aria-selected="false" tabindex="-1" class="search-option" data-search-index="${index}">
                        <span class="search-person"><strong>${escapeHtml(match.name)}</strong><span class="search-team" title="${match.isPerson ? '엑셀 도면의 좌석 구역 기준' : ''}">${escapeHtml(team)}</span></span>
                        <span class="search-location">${escapeHtml(office)} · ${seatAddress(match.item)}</span></button>`;
                }).join('');
                searchSummary.textContent = globalMatches.length ? `${globalMatches.length}개 결과 · 도면 기준` : '검색 결과가 없습니다';
                searchResults.scrollTop = 0;
                searchPopup.hidden = false;
                searchInput.setAttribute('aria-expanded', 'true');
                positionSearch();
            }

            function performSearch({ navigate = false } = {}) {
                const query = normalizeText(searchInput.value);
                clearSearchBtn.classList.toggle('hidden', query.length === 0);
                currentSearchQuery = query;
                globalMatches = OfficeMapSearch.findMatches(searchIndex, query, currentTab, directoryScope);
                currentMatchIdx = 0;
                if (!query) {
                    closeSearch();
                    document.querySelectorAll('.highlight-seat, .active-match').forEach(td => td.classList.remove('highlight-seat', 'active-match'));
                } else if (navigate) {
                    closeSearch();
                    if (globalMatches.length) goToMatch(globalMatches[0]);
                } else {
                    renderSearch();
                }
            }

            // 특정 매칭 타겟으로 이동 및 하이라이트 처리 함수
            function goToMatch(matchObj) {
                if (currentTab !== matchObj.tab) {
                    switchTab(matchObj.tab); 
                }
                
                document.querySelectorAll('.highlight-seat, .active-match').forEach(td => td.classList.remove('highlight-seat', 'active-match'));

                // 타겟 좌석 찾기
                const targetTd = document.querySelector(`.excel-table-layout td[data-r="${matchObj.item.r}"][data-c="${matchObj.item.c}"]`);
                
                if (targetTd) {
                    // 이전 active-match 포커스 제거 후 타겟에 부여 (주황색 강한 하이라이트)
                    document.querySelectorAll('.excel-table-layout td.active-match').forEach(td => td.classList.remove('active-match'));
                    targetTd.classList.add('active-match', 'highlight-seat');

                    camera.focus(targetTd);
                }
            }

            function selectMatch(index) {
                const match = globalMatches[index];
                if (!match) return;
                clearTimeout(searchTimer);
                currentMatchIdx = index;
                closeSearch();
                goToMatch(match);
                searchInput.blur();
            }

            searchResults.addEventListener('mousedown', event => event.preventDefault());
            searchResults.addEventListener('click', event => {
                const option = event.target.closest('[data-search-index]');
                if (option) selectMatch(Number(option.dataset.searchIndex));
            });
            searchInput.addEventListener('keydown', (e) => {
                if (e.isComposing || composing || e.keyCode === 229) return;
                if (e.key === 'Escape') { clearTimeout(searchTimer); closeSearch(); return; }
                if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                    e.preventDefault();
                    clearTimeout(searchTimer);
                    if (searchPopup.hidden || normalizeText(searchInput.value) !== currentSearchQuery) performSearch();
                    if (!globalMatches.length) return;
                    activeOption = activeOption < 0 ? (e.key === 'ArrowDown' ? 0 : globalMatches.length - 1)
                        : (activeOption + (e.key === 'ArrowDown' ? 1 : -1) + globalMatches.length) % globalMatches.length;
                    [...searchResults.children].forEach((option, index) => option.setAttribute('aria-selected', String(index === activeOption)));
                    const option = searchResults.children[activeOption];
                    searchInput.setAttribute('aria-activedescendant', option.id);
                    if (option.offsetTop < searchResults.scrollTop) searchResults.scrollTop = option.offsetTop;
                    else if (option.offsetTop + option.offsetHeight > searchResults.scrollTop + searchResults.clientHeight) {
                        searchResults.scrollTop = option.offsetTop + option.offsetHeight - searchResults.clientHeight;
                    }
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    clearTimeout(searchTimer);
                    if (searchPopup.hidden || normalizeText(searchInput.value) !== currentSearchQuery) performSearch();
                    selectMatch(activeOption < 0 ? 0 : activeOption);
                }
            });

            const queueSearch = () => {
                clearTimeout(searchTimer);
                closeSearch();
                directoryScope = null;
                if (!dirPanel.classList.contains('hidden')) toggleDirBtn.click();
                clearSearchBtn.classList.toggle('hidden', !searchInput.value);
                if (!composing) searchTimer = setTimeout(() => {
                    if (document.activeElement === searchInput) performSearch();
                }, 120);
            };
            searchInput.addEventListener('compositionstart', () => { composing = true; clearTimeout(searchTimer); closeSearch(); });
            searchInput.addEventListener('compositionend', () => { composing = false; queueSearch(); });
            searchInput.addEventListener('input', queueSearch);
            searchInput.addEventListener('focus', () => { directoryScope = null; if (searchInput.value) queueSearch(); });
            document.addEventListener('pointerdown', event => {
                if (!searchBox.contains(event.target) && !searchPopup.contains(event.target)) { clearTimeout(searchTimer); closeSearch(); }
            });
            document.addEventListener('focusin', event => {
                if (!searchBox.contains(event.target) && !searchPopup.contains(event.target)) { clearTimeout(searchTimer); closeSearch(); }
            });
            window.addEventListener('resize', positionSearch);
            window.visualViewport?.addEventListener('resize', positionSearch);
            window.visualViewport?.addEventListener('scroll', positionSearch);
            clearSearchBtn.addEventListener('click', () => {
                clearTimeout(searchTimer);
                directoryScope = null;
                searchInput.value = '';
                performSearch();
            });

            // --- 엑셀 파일 초기 로드 ---
            async function loadAllSharedExcels() {
                await Promise.all(Object.keys(offices).map(async key => {
                    try {
                        const response = await fetch(offices[key].file);
                        if (!response.ok) throw new Error('HTTP ' + response.status);
                        offices[key].data = await parseXlsxDirectXML(await response.arrayBuffer());
                    } catch (err) {
                        console.warn(`[${key}] 엑셀 파일을 불러오지 못했습니다.`, err);
                        try {
                            const saved = localStorage.getItem('savedFloorplanData_' + key);
                            if (saved) offices[key].data = JSON.parse(saved);
                        } catch {}
                    }
                    ensureSearchIndex(key);
                }));
                
                switchTab(currentTab);
                if (document.activeElement === searchInput && searchInput.value && !composing) performSearch();
            }
            
            loadAllSharedExcels().catch(error => { dynamicFloorplan.textContent = '지도를 불러오지 못했습니다. 새로고침해 주세요.'; console.error(error); });
        });
