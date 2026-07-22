const $ = id => document.getElementById(id);
const oneDriveConfig = { clientId: '22d17617-e89a-4cb4-a40e-f15ec8e71eb3', authority: 'https://login.microsoftonline.com/consumers', scope: 'openid profile Files.ReadWrite.AppFolder' };
let cloudSyncTimer;
const questionFields = ['question1', 'question2', 'question3', 'question4', 'question5'];
const profileFields = ['company', 'role', 'jobDescription', ...questionFields, 'keywords'];
const defaultProjects = [
  { id: 'pickup-2024', title: 'Chip Pick Up Tool 개선', period: '2024년', challenge: 'Eject Pin과 상하 방식 Pick Up Tool 사용 중 Chip 파손 불량률 2% 발생.', action: '상하 방식 대신 좌우 드래그 방식의 Pick Up Tool을 제안하고 설비 제조사 하드웨어·소프트웨어 엔지니어 2명과 설비 개조 및 공정 테스트를 진행.', result: 'Chip Pick Up 및 Chip 파손 불량률 0% 달성, 6개월간 검증 후 양산 공정 적용.', meta: '반도체 DIE Bonding 공정·설비 개선 · SPC · FMEA' },
  { id: 'ausn-2025', title: 'Substrate AuSn 설계 변경', period: '2025년 · 6개월', challenge: 'AuSn 영역이 Chip Isolation 구간을 넘어 본딩되며 통전 및 역전류 불량률 1% 발생.', action: '제작업체와 협업해 AuSn 사이즈 축소 설계를 수립하고 주문 제작 및 공정 적용 전 과정을 주도.', result: '통전·역전류 불량률 0% 달성.', meta: '외주 제작업체 협업 · 설계 변경 · 양산 적용' }
];
function readSaved(key) { try { return JSON.parse(localStorage.getItem(key) || 'null'); } catch { localStorage.removeItem(key); return null; } }
const legacy = readSaved('resumeStudio');
const saved = readSaved('resumeStudioV2');
const state = saved && typeof saved === 'object' ? saved : { projects: defaultProjects, profile: legacy || {}, allocations: {}, draft: '' };
if (!Array.isArray(state.projects)) state.projects = defaultProjects;
if (!state.profile || typeof state.profile !== 'object') state.profile = {};
if (!state.profile.question1 && state.profile.question) state.profile.question1 = state.profile.question;
if (!state.allocations || typeof state.allocations !== 'object') state.allocations = {};
if (!state.allocations.question1 && Array.isArray(state.selectedIds)) state.allocations.question1 = state.selectedIds;
if (typeof state.draft !== 'string') state.draft = '';
if (!Array.isArray(state.archives)) state.archives = [];
if (typeof state.activeArchiveId !== 'string') state.activeArchiveId = '';

function save() { localStorage.setItem('resumeStudioV2', JSON.stringify(state)); if (oneDriveToken()) queueCloudSync(); }
function showToast(message) { $('toast').textContent = message; $('toast').classList.add('visible'); setTimeout(() => $('toast').classList.remove('visible'), 2300); }
function redirectUri() { return `${window.location.origin}${window.location.pathname}`; }
function oneDriveToken() { const token = readSaved('resumeStudioOneDriveToken'); return token?.accessToken && token.expiresAt > Date.now() ? token.accessToken : ''; }
function base64Url(bytes) { return btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
async function beginOneDriveLogin() {
  try {
    if (window.location.protocol === 'file:') { showToast('OneDrive 로그인은 웹사이트를 배포한 뒤 사용할 수 있습니다.'); return; }
    if (!window.crypto?.subtle) throw new Error('crypto');
    updateOneDriveStatus('Microsoft 로그인 페이지로 이동 중…');
    const verifier = base64Url(crypto.getRandomValues(new Uint8Array(32))); const challenge = base64Url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier))); const loginState = base64Url(crypto.getRandomValues(new Uint8Array(16)));
    sessionStorage.setItem('resumeStudioPkceVerifier', verifier); sessionStorage.setItem('resumeStudioLoginState', loginState);
    const params = new URLSearchParams({ client_id: oneDriveConfig.clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query', scope: oneDriveConfig.scope, code_challenge: challenge, code_challenge_method: 'S256', state: loginState });
    window.location.href = `${oneDriveConfig.authority}/oauth2/v2.0/authorize?${params}`;
  } catch { updateOneDriveStatus(); showToast('로그인을 시작하지 못했습니다. 페이지를 새로고침한 뒤 다시 시도해 주세요.'); }
}
async function finishOneDriveLogin() {
  const params = new URLSearchParams(window.location.search); const code = params.get('code'); if (!code) return;
  const verifier = sessionStorage.getItem('resumeStudioPkceVerifier'); const expectedState = sessionStorage.getItem('resumeStudioLoginState'); if (!verifier || params.get('state') !== expectedState) { showToast('Microsoft 로그인 상태를 확인할 수 없습니다. 다시 로그인해 주세요.'); return; }
  try {
    const body = new URLSearchParams({ client_id: oneDriveConfig.clientId, grant_type: 'authorization_code', code, redirect_uri: redirectUri(), code_verifier: verifier });
    const response = await fetch(`${oneDriveConfig.authority}/oauth2/v2.0/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }); if (!response.ok) throw new Error('token');
    const token = await response.json(); localStorage.setItem('resumeStudioOneDriveToken', JSON.stringify({ accessToken: token.access_token, expiresAt: Date.now() + (token.expires_in - 60) * 1000 })); sessionStorage.removeItem('resumeStudioPkceVerifier'); sessionStorage.removeItem('resumeStudioLoginState'); history.replaceState({}, '', redirectUri()); await loadFromOneDrive(); updateOneDriveStatus(); showToast('OneDrive에 연결했습니다.');
  } catch { showToast('OneDrive 로그인에 실패했습니다. Redirect URI 설정을 확인해 주세요.'); }
}
async function graphRequest(path, options = {}) { const token = oneDriveToken(); if (!token) throw new Error('login'); return fetch(`https://graph.microsoft.com/v1.0${path}`, { ...options, headers: { ...(options.headers || {}), Authorization: `Bearer ${token}` } }); }
async function loadFromOneDrive() {
  const response = await graphRequest('/me/drive/special/approot:/resume-studio-data.json:/content');
  if (response.status === 404) { await syncToOneDrive(); return; } if (!response.ok) throw new Error('download'); const cloudState = await response.json();
  if (cloudState && Array.isArray(cloudState.projects) && Array.isArray(cloudState.archives)) { Object.keys(state).forEach(key => delete state[key]); Object.assign(state, cloudState); localStorage.setItem('resumeStudioV2', JSON.stringify(state)); location.reload(); }
}
async function syncToOneDrive() {
  try { const response = await graphRequest('/me/drive/special/approot:/resume-studio-data.json:/content', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(state) }); if (!response.ok) throw new Error('upload'); updateOneDriveStatus('OneDrive에 자동 저장됨'); }
  catch { updateOneDriveStatus('OneDrive 저장 대기'); }
}
function queueCloudSync() { clearTimeout(cloudSyncTimer); cloudSyncTimer = setTimeout(syncToOneDrive, 900); }
function updateOneDriveStatus(message) { const connected = Boolean(oneDriveToken()); $('oneDriveStatus').textContent = message || (connected ? 'OneDrive 연결됨' : 'OneDrive 연결 전'); $('oneDriveLogin').classList.toggle('hidden', connected); $('oneDriveSync').classList.toggle('hidden', !connected); }
function escapeHtml(value = '') { return String(value).replace(/[&<>'"]/g, char => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', "'":'&#039;', '"':'&quot;' })[char]); }
function profileData() { return Object.fromEntries(profileFields.map(id => [id, $(id).value.trim()])); }
function usedProjectIds(exceptQuestion = '') { return Object.entries(state.allocations).filter(([question]) => question !== exceptQuestion).flatMap(([, ids]) => Array.isArray(ids) ? ids : []); }
function projectsFor(questionId) { return state.projects.filter(project => (state.allocations[questionId] || []).includes(project.id)); }
function allSelectedProjects() { return state.projects.filter(project => Object.values(state.allocations).flat().includes(project.id)); }
function filledQuestions(data = profileData()) { return questionFields.filter(id => data[id]); }
function projectContent(project) { return `${project.title} ${project.challenge} ${project.action} ${project.result} ${project.meta}`.toLowerCase(); }
function termsFor(data, question = '') { return `${data.role} ${data.jobDescription} ${data.keywords} ${question}`.toLowerCase().split(/[\s,/·()]+/).filter(term => term.length > 1); }
function scoreProject(project, data, question = '') { const content = projectContent(project); return termsFor(data, question).reduce((score, term) => score + (content.includes(term) ? 1 : 0), 0); }
function recommendedProjects(data, question) { return [...state.projects].sort((a, b) => scoreProject(b, data, question) - scoreProject(a, data, question)); }
function coreCompetencies(data) {
  const source = `${data.role} ${data.jobDescription} ${data.keywords} ${allSelectedProjects().map(projectContent).join(' ')}`.toLowerCase();
  const candidates = [
    ['공정 개선', ['공정', '개선', '수율', 'bonding']], ['설비 개선·관리', ['설비', '장비', 'tool', '자동화']], ['품질 관리·불량 분석', ['품질', '불량', 'spc', 'fmea', '신뢰성']], ['데이터 기반 문제 해결', ['데이터', '분석', '지표', '수치', '통계']], ['양산 안정화', ['양산', '생산', '가동률', '생산성', 'ct']], ['협력사 기술 협업', ['협력사', '제작업체', '외주', '협업', 'vendor']], ['설계·사양 최적화', ['설계', '사양', '치수', 'substrate', 'ausn']], ['원인 분석 및 재발 방지', ['원인', '재발', '리스크', '문제']]
  ];
  return candidates.map(([name, terms], index) => ({ name, index, score: terms.reduce((total, term) => total + (source.includes(term) ? 1 : 0), 0) })).sort((a, b) => b.score - a.score || a.index - b.index).slice(0, 5).map(item => item.name);
}
function renderCompetencies() { $('competencyList').innerHTML = coreCompetencies(profileData()).map(name => `<span class="competency-chip">${name}</span>`).join(''); }
function switchTab(tab) {
  ['info', 'selection', 'draft', 'interview', 'archive', 'projects'].forEach(name => { $(`${name}Panel`).classList.toggle('hidden', name !== tab); });
  document.querySelectorAll('.tab-button').forEach(button => button.classList.toggle('active', button.dataset.tab === tab));
  if (tab === 'selection') renderAllocations(); if (tab === 'draft') renderDraft(); if (tab === 'interview') renderInterview(); if (tab === 'archive') renderArchives();
}
function renderAllocations() {
  const data = profileData(); const questions = filledQuestions(data); $('recommendationSummary').innerHTML = `<strong>${escapeHtml(data.role || '지원 직무')}</strong>에 필요한 핵심역량은 ${coreCompetencies(data).join(' · ')}입니다. 질문마다 추천 이력을 하나씩 선택해 보세요.`;
  $('questionAllocation').innerHTML = questions.map((questionId, index) => {
    const question = data[questionId]; const allocated = state.allocations[questionId] || []; const used = usedProjectIds(questionId); const recommendations = recommendedProjects(data, question); const topId = recommendations[0]?.id;
    const options = recommendations.map(project => { const disabled = used.includes(project.id); const checked = allocated.includes(project.id); return `<label class="allocation-option ${project.id === topId ? 'recommended' : ''} ${disabled ? 'unavailable' : ''}"><input type="checkbox" data-question="${questionId}" value="${project.id}" ${checked ? 'checked' : ''} ${disabled ? 'disabled' : ''}><span><strong>${escapeHtml(project.title)}${project.id === topId ? ' <em>추천</em>' : ''}</strong><small>${escapeHtml(project.result)}</small></span></label>`; }).join('');
    return `<article class="allocation-card"><p class="allocation-number">질문 ${index + 1}</p><h3>${escapeHtml(question)}</h3><p class="allocation-hint">추천: ${escapeHtml(recommendations[0]?.title || '저장된 프로젝트 없음')} — JD 및 핵심역량과의 연관성을 기준으로 추천했습니다.</p><div class="allocation-options">${options}</div></article>`;
  }).join('') || '<div class="empty-library">지원 정보 탭에서 질문 1을 입력해 주세요.</div>';
  document.querySelectorAll('#questionAllocation input').forEach(input => input.addEventListener('change', () => { const id = input.value; const question = input.dataset.question; const ids = state.allocations[question] || []; state.allocations[question] = input.checked ? [...ids, id] : ids.filter(item => item !== id); save(); renderAllocations(); }));
}
function phrase(text) { return String(text || '').trim().replace(/[.!?]+$/, '').replace(/\s+/g, ' '); }
function naturalStory(project, index) { const opening = index === 0 ? '대표적으로' : '또한'; return `${opening} ${project.title} 프로젝트를 수행하며 ${phrase(project.challenge)}라는 문제를 확인했습니다. 불량 원인을 제거하고 생산 공정에 적용 가능한 개선안을 마련하는 것을 과제로 삼았습니다. 이에 ${phrase(project.action)} 방식으로 개선을 추진했습니다. 그 결과 ${phrase(project.result)}라는 개선 결과를 확인했습니다.`; }
function makeAnswer(question, projects, data) {
  const company = data.company || '지원 기업'; const role = data.role || '지원 직무'; const atsTerms = data.keywords || data.jobDescription || '공정 개선, 설비 개선, 품질 향상, 불량 분석, SPC, FMEA'; const evidence = projects.map(naturalStory).join(' ');
  return `${question}\n\n${company} ${role}에 지원하는 이유는 현장의 문제를 수치로 확인하고, 공정·설비 개선을 통해 검증 가능한 성과로 전환해 온 경험이 있기 때문입니다.\n\n${evidence}\n\n이 경험을 통해 ${atsTerms} 역량을 실무에서 축적했습니다. 개선안의 선택 근거, 협업 방식, 검증 기간과 수치 성과를 바탕으로 입사 후에도 과장된 목표보다 현장 데이터와 재현 가능한 실행으로 품질과 생산 안정성을 높이겠습니다.`;
}
function makeDraft(data) { return filledQuestions(data).map(questionId => makeAnswer(data[questionId], projectsFor(questionId), data)).join('\n\n━━━━━━━━━━━━━━━━━━━━\n\n'); }
function fileDate(date = new Date()) { return date.toISOString().slice(0, 10); }
function archiveFilename(data) { const company = (data.company || '지원기업').replace(/[\\/:*?"<>|]/g, '').trim() || '지원기업'; return `${company}_${fileDate()}_자기소개서`; }
function saveArchive() {
  const data = profileData(); const filename = archiveFilename(data); const current = state.archives.find(item => item.id === state.activeArchiveId);
  if (current) { Object.assign(current, { filename, company: data.company || '지원기업', role: data.role || '', draft: state.draft, updatedAt: new Date().toISOString() }); }
  else { const item = { id: `archive-${Date.now()}`, filename, company: data.company || '지원기업', role: data.role || '', createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(), draft: state.draft, profile: data, allocations: JSON.parse(JSON.stringify(state.allocations)) }; state.archives.unshift(item); state.activeArchiveId = item.id; }
  save(); renderArchiveMeta();
}
function renderArchiveMeta() { const item = state.archives.find(entry => entry.id === state.activeArchiveId); $('archiveName').textContent = item ? `${item.filename} · 보관함 저장됨` : '저장 전'; }
function interviewQuestions() {
  const data = profileData(); const projects = allSelectedProjects(); const first = projects[0]; const second = projects[1] || first;
  if (!first) return [];
  return [
    { q: `${first.title}에서 가장 먼저 확인한 원인은 무엇입니까?`, a: `${phrase(first.challenge)}라는 현상을 확인한 뒤, 공정 조건과 설비 동작을 중심으로 원인을 검토했습니다. 단순히 불량을 선별하는 대신 ${phrase(first.action)} 방식으로 원인을 제거하는 개선안을 추진했습니다. 그 결과 ${phrase(first.result)}라는 성과를 확인했습니다.` },
    { q: '개선안의 효과와 양산 적용 가능성은 어떻게 검증했습니까?', a: `개선 후에는 불량률과 공정 적용 결과를 기준으로 효과를 확인했습니다. 특히 ${phrase(first.result)}라는 수치를 통해 개선 효과를 판단했으며, 검증 과정에서 확인된 조건을 공정에 적용했습니다.` },
    { q: '협력사 또는 유관 부서와 의견이 달랐을 때 어떻게 조율했습니까?', a: `${first.meta ? `${first.meta} 경험을 바탕으로 ` : ''}문제 현상과 목표 수치를 먼저 공유하고, 역할별로 실행 가능한 대안을 정리했습니다. 개선안은 설비·제작 관점의 의견을 반영해 검증했으며, 결과 데이터를 기준으로 공정 적용 여부를 결정했습니다.` },
    { q: `${second.title} 경험이 ${data.role || '지원 직무'}에 어떤 강점이 된다고 생각합니까?`, a: `${phrase(second.challenge)} 문제를 해결하는 과정에서 원인 분석, 개선안 수립, 협업, 양산 적용까지 경험했습니다. 이 경험을 통해 ${data.keywords || '공정 개선과 품질 향상'} 업무에서 현장 문제를 수치로 설명하고 실행으로 연결하는 역량을 갖추었습니다.` },
    { q: `입사 후 ${data.company || '지원 기업'}에서 가장 먼저 개선하고 싶은 부분은 무엇입니까?`, a: `입사 초기에는 담당 공정의 불량 유형, 설비 상태, 생산 데이터를 우선 파악하겠습니다. 이후 영향도가 큰 문제부터 개선 과제를 설정하고, 검증 가능한 수치 목표와 협업 계획을 바탕으로 품질과 생산 안정성 향상에 기여하겠습니다.` }
  ];
}
function renderInterview() { const questions = interviewQuestions(); $('interviewList').innerHTML = questions.length ? questions.map((item, index) => `<article class="interview-card"><p>예상 질문 ${index + 1}</p><h3>${escapeHtml(item.q)}</h3><div><strong>모범 답안</strong><p>${escapeHtml(item.a)}</p></div></article>`).join('') : '<div class="empty-library">초안을 만들기 전에 질문별 이력을 선택해 주세요.</div>'; }
function renderArchives() {
  const keyword = $('archiveSearch').value.trim().toLowerCase(); const entries = state.archives.filter(item => `${item.filename} ${item.company} ${item.role}`.toLowerCase().includes(keyword)); $('archiveCount').textContent = state.archives.length; $('archiveMeta').textContent = `${state.archives.length}개 저장됨`; $('archiveList').innerHTML = entries.length ? entries.map(item => `<article class="archive-card"><div><p>${escapeHtml(item.filename)}</p><h3>${escapeHtml(item.company)} · ${escapeHtml(item.role || '직무 미입력')}</h3><small>${new Date(item.createdAt).toLocaleDateString('ko-KR')} 생성 · ${item.draft.length.toLocaleString()}자</small></div><button class="secondary" data-load-archive="${item.id}">불러오기</button></article>`).join('') : '<div class="empty-library">검색 결과가 없습니다.</div>'; document.querySelectorAll('[data-load-archive]').forEach(button => button.addEventListener('click', () => loadArchive(button.dataset.loadArchive)));
}
function loadArchive(id) { const item = state.archives.find(entry => entry.id === id); if (!item) return; state.activeArchiveId = id; state.draft = item.draft; state.profile = item.profile || state.profile; state.allocations = item.allocations || state.allocations; profileFields.forEach(field => { $(field).value = state.profile[field] || ''; }); save(); renderCompetencies(); switchTab('draft'); showToast('저장한 초안을 불러왔습니다.'); }
function exportBackup() {
  const payload = { version: 1, exportedAt: new Date().toISOString(), data: state };
  const url = URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' }));
  const link = document.createElement('a'); link.href = url; link.download = `지원서스튜디오_백업_${fileDate()}.json`; link.click(); URL.revokeObjectURL(url);
  showToast('전체 자료 백업 파일을 만들었습니다.');
}
function importBackup(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const backup = JSON.parse(reader.result); const data = backup?.data;
      if (!data || !Array.isArray(data.projects) || !Array.isArray(data.archives)) throw new Error('invalid');
      if (!window.confirm('현재 브라우저의 프로젝트와 보관함을 백업 파일 내용으로 바꾸겠습니까?')) return;
      Object.keys(state).forEach(key => delete state[key]); Object.assign(state, data); save(); location.reload();
    } catch { showToast('지원서 스튜디오 백업 파일인지 확인해 주세요.'); }
  };
  reader.readAsText(file);
}
function feedbackFor(text, data, proofread = false) {
  const informal = (text.match(/(?:했다|이다|된다|하였다|겠다)\./g) || []).length; const spaces = / {2,}/.test(text); const commonTypos = /(되요|않되|왠만|몇일)/.test(text); const cards = [
    { title: '직무 연관성', score: data.role && data.company ? '좋음' : '보완 필요', text: data.role && data.company ? '회사와 직무가 초안에 자연스럽게 연결되어 있습니다.' : '회사명과 지원 직무를 입력해 주세요.' },
    { title: '수치 성과', score: /\d/.test(text) ? '좋음' : '보완 필요', text: /\d/.test(text) ? '등록한 수치 성과가 포함되어 있습니다.' : '프로젝트에 기간이나 개선율 같은 수치를 보태 주세요.' },
    { title: proofread ? '맞춤법·표현' : '이력 배정', score: proofread ? (informal || spaces || commonTypos ? '확인 필요' : '기초 점검 완료') : `${allSelectedProjects().length}개 선택`, text: proofread ? (informal ? '일부 문장에 -다체가 남아 있습니다. -습니다체로 다듬어 주세요.' : spaces || commonTypos ? '공백 또는 자주 틀리는 표현을 확인해 주세요.' : '기초 맞춤법 패턴과 문체를 점검했습니다. 최종 제출 전 전문 교정을 권장합니다.') : '질문별로 중복되지 않는 프로젝트만 초안에 반영했습니다.' }
  ]; $('feedbackCards').innerHTML = cards.map(card => `<article class="feedback"><span class="score">${card.score}</span><h3>${card.title}</h3><p>${card.text}</p></article>`).join('');
}
function renderDraft() { $('draft').value = state.draft; $('charCount').textContent = `${state.draft.length.toLocaleString()}자`; renderArchiveMeta(); feedbackFor(state.draft, profileData()); }
function renderProjectCards() { $('projectCards').innerHTML = state.projects.length ? state.projects.map(project => `<article class="project-card"><div class="project-card-head"><div><p>${escapeHtml(project.period || '기간 미입력')}</p><h3>${escapeHtml(project.title)}</h3></div><div class="card-actions"><button data-edit="${project.id}" class="text-button">수정</button><button data-delete="${project.id}" class="delete-button">삭제</button></div></div><dl><div><dt>문제</dt><dd>${escapeHtml(project.challenge)}</dd></div><div><dt>실행</dt><dd>${escapeHtml(project.action)}</dd></div><div><dt>성과</dt><dd>${escapeHtml(project.result)}</dd></div></dl>${project.meta ? `<p class="project-meta">${escapeHtml(project.meta)}</p>` : ''}</article>`).join('') : '<div class="empty-library">아직 저장한 프로젝트가 없습니다.</div>'; $('projectCount').textContent = state.projects.length; document.querySelectorAll('[data-edit]').forEach(button => button.addEventListener('click', () => beginEdit(button.dataset.edit))); document.querySelectorAll('[data-delete]').forEach(button => button.addEventListener('click', () => removeProject(button.dataset.delete))); }
function beginEdit(id) { const project = state.projects.find(item => item.id === id); if (!project) return; $('editingProjectId').value = project.id; ['Title','Period','Challenge','Action','Result','Meta'].forEach(key => { $(`project${key}`).value = project[key.toLowerCase()]; }); $('projectFormTitle').textContent = '프로젝트 수정'; $('cancelEdit').classList.remove('hidden'); }
function resetProjectForm() { $('projectForm').reset(); $('editingProjectId').value = ''; $('projectFormTitle').textContent = '프로젝트 추가'; $('cancelEdit').classList.add('hidden'); }
function removeProject(id) { state.projects = state.projects.filter(project => project.id !== id); Object.keys(state.allocations).forEach(question => { state.allocations[question] = (state.allocations[question] || []).filter(item => item !== id); }); save(); renderProjectCards(); showToast('프로젝트를 삭제했습니다.'); }

profileFields.forEach(id => { $(id).value = state.profile[id] || ''; $(id).addEventListener('input', () => { state.profile = profileData(); save(); renderCompetencies(); }); });
document.querySelectorAll('.tab-button').forEach(button => button.addEventListener('click', () => switchTab(button.dataset.tab)));
$('profileForm').addEventListener('submit', event => { event.preventDefault(); state.profile = profileData(); save(); renderCompetencies(); switchTab('selection'); });
$('backToInfo').addEventListener('click', () => switchTab('info'));
$('goDraft').addEventListener('click', () => { const data = profileData(); const missing = filledQuestions(data).filter(question => !projectsFor(question).length); if (missing.length) { showToast(`질문 ${missing.map(id => id.replace('question', '')).join(', ')}에 프로젝트를 선택해 주세요.`); return; } state.profile = data; state.draft = makeDraft(data); state.activeArchiveId = ''; saveArchive(); switchTab('draft'); showToast('질문별로 선택한 이력으로 초안을 만들고 보관함에 저장했습니다.'); });
$('draft').addEventListener('input', () => { state.draft = $('draft').value; save(); $('charCount').textContent = `${state.draft.length.toLocaleString()}자`; });
$('reviewButton').addEventListener('click', () => { feedbackFor(state.draft, profileData()); showToast('초안 내용을 다시 점검했습니다.'); });
$('proofreadButton').addEventListener('click', () => { feedbackFor(state.draft, profileData(), true); showToast('맞춤법·표현 기초 검수를 완료했습니다.'); });
$('copyButton').addEventListener('click', async () => { await navigator.clipboard.writeText(state.draft); showToast('초안을 클립보드에 복사했습니다.'); });
$('saveButton').addEventListener('click', () => { saveArchive(); showToast('현재 초안을 보관함에 저장했습니다.'); });
$('goInterview').addEventListener('click', () => { saveArchive(); switchTab('interview'); });
$('backToDraft').addEventListener('click', () => switchTab('draft'));
$('archiveSearch').addEventListener('input', renderArchives);
$('exportBackup').addEventListener('click', exportBackup);
$('importBackup').addEventListener('change', event => { if (event.target.files?.[0]) importBackup(event.target.files[0]); event.target.value = ''; });
$('oneDriveLogin').addEventListener('click', beginOneDriveLogin);
$('oneDriveSync').addEventListener('click', async () => { await syncToOneDrive(); showToast('OneDrive에 저장을 요청했습니다.'); });
$('projectForm').addEventListener('submit', event => { event.preventDefault(); const id = $('editingProjectId').value || `project-${Date.now()}`; const project = { id, title: $('projectTitle').value.trim(), period: $('projectPeriod').value.trim(), challenge: $('projectChallenge').value.trim(), action: $('projectAction').value.trim(), result: $('projectResult').value.trim(), meta: $('projectMeta').value.trim() }; const index = state.projects.findIndex(item => item.id === id); if (index === -1) state.projects.unshift(project); else state.projects[index] = project; save(); renderProjectCards(); resetProjectForm(); showToast(index === -1 ? '프로젝트를 추가했습니다.' : '프로젝트를 수정했습니다.'); });
$('cancelEdit').addEventListener('click', resetProjectForm);
renderCompetencies(); renderProjectCards(); renderArchives(); if (state.draft) renderDraft();
updateOneDriveStatus(); finishOneDriveLogin();
