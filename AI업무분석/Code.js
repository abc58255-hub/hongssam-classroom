// AI 업무분석 앱 — 공문/메시지/사진을 AI로 분석해 할 일을 뽑고 TickTick에 등록
// 키(OpenRouter·TickTick)·인증은 ClassCore 라이브러리에 위임 — 하드코딩 없음
// 설정 키: OpenRouter키 / AI모델명 / TickTick클라이언트ID / TickTick시크릿 / TickTick액세스토큰 / TickTick갱신토큰

// ── 라우팅 ──
function doGet() {
  try { ClassCore.registerAppUrl('AI업무분석', _serviceUrl_()); } catch(_) {}
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('AI 업무분석')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function _serviceUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(_) { return ''; }
}

// ── 교사 인증 (ClassCore 위임) ──
function teacherLogin(pw)            { return ClassCore.teacherLogin(pw); }
function teacherLogout(token)        { return ClassCore.teacherLogout(token); }
function validateTeacherSession(token) { return ClassCore.verifyTeacher(token); }

// ── 설정 (ClassCore에서 읽기) ──
function _cfg(k, d) { try { return ClassCore.getConfig(k, d || ''); } catch(_) { return d || ''; } }
function _setCfg(k, v) { try { ClassCore.setConfig(k, v); } catch(_) {} }

function getApiSettings() {
  var m = _cfg('AI모델명', '') || _cfg('AI_MODEL', '');
  if (m === 'google/gemini-2.5-flash-preview') m = 'google/gemini-2.5-flash';
  if (m === 'google/gemini-2.5-pro-preview')   m = 'google/gemini-2.5-pro';
  return {
    openrouterKey: _cfg('OpenRouter키', '') || _cfg('OPENROUTER_KEY', ''),
    model:         m || 'anthropic/claude-3.5-sonnet',
    ttClientId:    _cfg('TickTick클라이언트ID', ''),
    ttSecret:      _cfg('TickTick시크릿', ''),
    ttToken:       _cfg('TickTick액세스토큰', ''),
    ttRefresh:     _cfg('TickTick갱신토큰', '')
  };
}

// ── AI 분석 ──
function analyzeTaskText(text, fileAttachment) {
  try {
    var cfg = getApiSettings();
    if (!cfg.openrouterKey) return { success: false, message: 'OpenRouter 키가 없습니다. ClassCore 설정(OpenRouter키)을 확인하세요.' };

    var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 MM월 dd일 (E)');
    var systemPrompt = '너는 한국 교사의 업무를 도와주는 AI야. 공문, 메시지, 일정이 적힌 사진, 첨부파일을 분석해서 반드시 유효한 JSON만 반환해. 사진이 첨부되면 사진 속 텍스트를 꼼꼼히 읽어서 날짜·시간·할 일을 모두 추출해. 다른 텍스트, 마크다운 코드블록, 설명은 절대 포함하지 마.';
    var userPrompt = '오늘 날짜: ' + today + '\n\n'
      + '아래 공문/메시지를 분석하고 정확히 다음 JSON 형식으로 반환해.\n\n'
      + '⚠️ 문체 규칙 (반드시 준수):\n'
      + '- 모든 텍스트 필드는 개조식으로 작성\n'
      + '- "합니다/입니다/됩니다/있습니다" 등 문장형 종결어미 절대 금지\n'
      + '- 명사형 또는 동사 원형으로 짧게 끝낼 것 (예: "제출", "확인 필요", "양식 작성")\n\n'
      + '기타 규칙:\n'
      + '- deadline: 날짜 언급 있으면 yyyy-MM-dd, 없으면 null\n'
      + '- time: 시간 언급 있으면 HH:mm (24시간), 없으면 null\n'
      + '- priority: 5=매우긴급, 3=보통, 1=낮음\n'
      + '- project: 학교업무/연수/학급운영/행사/제출/회의/기타 중 하나\n'
      + '- tags: 핵심 키워드 최대 3개 (명사, 한글)\n'
      + '- checklist: 세부 단계 있으면 배열, 없으면 빈 배열\n'
      + '- repeat: 반복 언급 있으면 RRULE 형식, 없으면 null\n\n'
      + '{\n'
      + '  "summary": "핵심 내용 2~3줄 개조식 요약 (문장 아닌 항목 나열)",\n'
      + '  "tasks": [\n'
      + '    {\n'
      + '      "title": "할 일 제목 (10자 이내, 명사형)",\n'
      + '      "deadline": "yyyy-MM-dd 또는 null",\n'
      + '      "time": "HH:mm 또는 null",\n'
      + '      "detail": "핵심 내용만 개조식으로 (2~3항목, 각 항목 20자 이내)",\n'
      + '      "priority": 1~5,\n'
      + '      "project": "분류명",\n'
      + '      "tags": ["태그1", "태그2"],\n'
      + '      "checklist": ["단계1", "단계2"],\n'
      + '      "repeat": "RRULE 또는 null"\n'
      + '    }\n'
      + '  ],\n'
      + '  "cautions": ["주의사항 개조식"]\n'
      + '}\n\n'
      + '[분석할 내용]\n' + text;

    var userContent;
    if (fileAttachment && fileAttachment.data) {
      var ext = String(fileAttachment.name || '').split('.').pop().toLowerCase();
      var isImage = ['png','jpg','jpeg','gif','webp','heic','heif'].indexOf(ext) >= 0;
      var isPdf = ext === 'pdf';
      if (isImage) {
        var mimeType = (fileAttachment.mimeType && fileAttachment.mimeType.indexOf('image/') === 0) ? fileAttachment.mimeType : 'image/jpeg';
        userContent = [
          { type: 'image_url', image_url: { url: 'data:' + mimeType + ';base64,' + fileAttachment.data } },
          { type: 'text', text: userPrompt + (text ? '\n\n[추가 텍스트]\n' + text : '') }
        ];
      } else if (isPdf) {
        userContent = userPrompt + '\n\n[PDF 첨부: ' + fileAttachment.name + ' - 텍스트 내용을 직접 붙여넣어 주세요]';
      } else {
        try {
          var decoded = Utilities.newBlob(Utilities.base64Decode(fileAttachment.data)).getDataAsString('UTF-8');
          userContent = userPrompt + '\n\n[첨부 파일: ' + fileAttachment.name + ']\n' + decoded.substring(0, 8000);
        } catch(e) { userContent = userPrompt + '\n\n[파일 첨부됨: ' + fileAttachment.name + ' - 텍스트 추출 실패]'; }
      }
    } else { userContent = userPrompt; }

    var payload = {
      model: cfg.model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user',   content: userContent }
      ],
      max_tokens: 8000,
      temperature: 0.3
    };

    var res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + cfg.openrouterKey,
        'Content-Type': 'application/json',
        'HTTP-Referer': 'https://script.google.com',
        'X-Title': 'AI Task Analyzer'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    if (code !== 200) {
      var errBody; try { errBody = JSON.parse(res.getContentText()); } catch(_) { errBody = {}; }
      var errMsg = (errBody.error && errBody.error.message) ? errBody.error.message : res.getContentText();
      return { success: false, message: 'API 오류 (' + code + '): ' + errMsg };
    }
    var body = JSON.parse(res.getContentText());
    var content = body.choices[0].message.content.trim();
    content = content.replace(/^```[a-z]*\n?/i, '').replace(/\n?```$/, '').trim();
    var parsed = _parseAiJsonLoose_(content);
    if (!parsed) return { success: false, message: 'JSON 파싱 실패: ' + content.substring(0, 150) };
    return { success: true, result: parsed, model: cfg.model };
  } catch(e) { return { success: false, message: '분석 오류: ' + e.toString() }; }
}

function _parseAiJsonLoose_(text) {
  if (!text) return null;
  var start = text.indexOf('{');
  if (start < 0) return null;
  var body = text.substring(start);
  try { return JSON.parse(body); } catch(_) {}
  var lastBrace = body.lastIndexOf('}');
  if (lastBrace > 0) { try { return JSON.parse(body.substring(0, lastBrace + 1)); } catch(_) {} }
  var stack = [], inStr = false, esc = false;
  for (var i = 0; i < body.length; i++) {
    var c = body[i];
    if (esc) { esc = false; continue; }
    if (c === '\\') { esc = true; continue; }
    if (c === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (c === '{' || c === '[') stack.push(c);
    else if (c === '}' || c === ']') stack.pop();
  }
  var patched = body;
  if (inStr) patched += '"';
  patched = patched.replace(/[\s]*,[\s]*$/, '').replace(/[\s]*:[\s]*$/, ':null');
  for (var j = stack.length - 1; j >= 0; j--) { patched += (stack[j] === '{') ? '}' : ']'; }
  try { return JSON.parse(patched); } catch(_) {}
  try { return JSON.parse(patched.replace(/,(\s*[}\]])/g, '$1')); } catch(_) { return null; }
}

// ── TickTick OAuth ──
var TT_REDIRECT_URI = 'https://example.com';

function getTickTickAuthUrl() {
  try {
    var cfg = getApiSettings();
    if (!cfg.ttClientId) return { success: false, message: 'TickTick 클라이언트ID가 없습니다.' };
    var authUrl = 'https://ticktick.com/oauth/authorize'
      + '?client_id=' + encodeURIComponent(cfg.ttClientId)
      + '&scope=' + encodeURIComponent('tasks:read tasks:write')
      + '&redirect_uri=' + encodeURIComponent(TT_REDIRECT_URI)
      + '&response_type=code';
    return { success: true, authUrl: authUrl };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function exchangeTickTickCode(code) {
  try {
    var cfg = getApiSettings();
    var credential = Utilities.base64Encode(cfg.ttClientId + ':' + cfg.ttSecret);
    var res = UrlFetchApp.fetch('https://ticktick.com/oauth/token', {
      method: 'post',
      headers: { 'Authorization': 'Basic ' + credential, 'Content-Type': 'application/x-www-form-urlencoded' },
      payload: 'code=' + encodeURIComponent(code) + '&grant_type=authorization_code&redirect_uri=' + encodeURIComponent(TT_REDIRECT_URI),
      muteHttpExceptions: true
    });
    var token = JSON.parse(res.getContentText());
    if (!token.access_token) return { success: false, message: '코드가 잘못됐거나 만료됐어요: ' + res.getContentText() };
    _setCfg('TickTick액세스토큰', token.access_token);
    if (token.refresh_token) _setCfg('TickTick갱신토큰', token.refresh_token);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function refreshTickTickToken() {
  try {
    var cfg = getApiSettings();
    if (!cfg.ttRefresh) return { success: false, message: 'Refresh Token 없음. 재인증 필요.' };
    var credential = Utilities.base64Encode(cfg.ttClientId + ':' + cfg.ttSecret);
    var res = UrlFetchApp.fetch('https://ticktick.com/oauth/token', {
      method: 'post',
      headers: { 'Authorization': 'Basic ' + credential, 'Content-Type': 'application/x-www-form-urlencoded' },
      payload: 'grant_type=refresh_token&refresh_token=' + encodeURIComponent(cfg.ttRefresh),
      muteHttpExceptions: true
    });
    var token = JSON.parse(res.getContentText());
    if (!token.access_token) return { success: false, message: '갱신 실패. 재인증 필요.' };
    _setCfg('TickTick액세스토큰', token.access_token);
    if (token.refresh_token) _setCfg('TickTick갱신토큰', token.refresh_token);
    return { success: true, token: token.access_token };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function addToTickTick(taskData) {
  try {
    var cfg = getApiSettings();
    var token = cfg.ttToken;
    if (!token) return { success: false, message: 'TickTick 미연동. 설정에서 먼저 연동해주세요.' };
    if (!cfg.ttClientId) return { success: false, message: 'TickTick 클라이언트ID가 없습니다.' };

    function postTask(accessToken) {
      var pri = parseInt(taskData.priority);
      if ([0,1,3,5].indexOf(pri) < 0) pri = 0;
      var body = { title: String(taskData.title || '').trim() || '제목 없음', content: taskData.detail || '', priority: pri };
      if (taskData.deadline) {
        if (taskData.time) { body.dueDate = taskData.deadline + 'T' + taskData.time + ':00+0900'; body.isAllDay = false; }
        else { body.dueDate = taskData.deadline + 'T00:00:00+0900'; body.isAllDay = true; }
        body.timeZone = 'Asia/Seoul';
      }
      if (taskData.projectId) body.projectId = taskData.projectId;
      if (Array.isArray(taskData.tags) && taskData.tags.length > 0) body.tags = taskData.tags.map(function(t){ return String(t); }).filter(Boolean);
      if (Array.isArray(taskData.checklist) && taskData.checklist.length > 0) {
        body.items = taskData.checklist.map(function(item, i) { return { title: String(item), status: 0, sortOrder: i }; });
      }
      if (taskData.repeat && /^RRULE:/i.test(String(taskData.repeat))) body.repeatFlag = taskData.repeat;
      return UrlFetchApp.fetch('https://ticktick.com/open/v1/task', {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
    }

    var res = postTask(token);
    var code = res.getResponseCode();
    if (code === 401) {
      var refreshed = refreshTickTickToken();
      if (!refreshed.success) return { success: false, message: '토큰 만료. ' + refreshed.message };
      res = postTask(refreshed.token);
      code = res.getResponseCode();
    }
    if (code === 200 || code === 201) return { success: true };
    return { success: false, message: 'TickTick 오류 (' + code + '): ' + res.getContentText() };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getTickTickProjects() {
  try {
    var cfg = getApiSettings();
    var token = cfg.ttToken;
    if (!token) return { success: false, message: 'TickTick 미연동' };
    function fetchProjects(accessToken) {
      return UrlFetchApp.fetch('https://ticktick.com/open/v1/project', {
        method: 'get', headers: { 'Authorization': 'Bearer ' + accessToken }, muteHttpExceptions: true
      });
    }
    var res = fetchProjects(token);
    if (res.getResponseCode() === 401) {
      var refreshed = refreshTickTickToken();
      if (!refreshed.success) return { success: false, message: refreshed.message };
      res = fetchProjects(refreshed.token);
    }
    var projects = JSON.parse(res.getContentText());
    return { success: true, projects: projects.map(function(p) { return { id: p.id, name: p.name, color: p.color || '#6366f1' }; }) };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getTickTickStatus() {
  try {
    var cfg = getApiSettings();
    return { success: true, connected: !!(cfg.ttToken), model: cfg.model, hasApiKey: !!(cfg.openrouterKey) };
  } catch(e) { return { success: false, connected: false }; }
}


// ── PWA(Pages) 프론트 → GAS 백엔드 (google.script.run 어댑터) + 토큰 게이트 ──
var RPC_WHITELIST = ["addToTickTick", "analyzeTaskText", "exchangeTickTickCode", "getTickTickAuthUrl", "getTickTickProjects", "getTickTickStatus", "teacherLogin", "teacherLogout", "validateTeacherSession"];
var RPC_NOAUTH = ['teacherLogin','teacherLogout'];
function _atOk_(token) { try { var v = ClassCore.verifyTeacher(token); return (v === true) || !!(v && (v.success || v.valid || v.ok)); } catch(_) { return false; } }
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (RPC_WHITELIST.indexOf(req.fn) < 0) throw new Error('허용되지 않은 함수: ' + req.fn);
    if (RPC_NOAUTH.indexOf(req.fn) < 0 && !_atOk_(req.token)) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: '로그인이 필요해요', needAuth: true })).setMimeType(ContentService.MimeType.JSON);
    }
    var fn = globalThis[req.fn];
    if (typeof fn !== 'function') throw new Error('함수를 찾을 수 없음: ' + req.fn);
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
