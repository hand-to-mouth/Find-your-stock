/* app.js — 산업연관표 연관산업 탐색기 (2단계: 4개 분류 단위 + 드릴다운)
 *
 * 흐름
 *   1) manifest.json · companies.csv · sector-tree.xlsx 를 읽어 준비
 *   2) 분류 단위(대/중/소/기본)를 고르면 그 단위의 Excel 을 내려받아 파싱
 *      → 단위마다 파일이 다르고, 부문 코드도 다릅니다.
 *   3) STEP 2 원자료 탐색 / STEP 3 연관산업 / 드릴다운 / STEP 4 대표기업
 *
 * 드릴다운
 *   대분류를 고르면 그 안의 중분류들을, 중분류를 고르면 소분류들을
 *   **그룹 평균**으로 집계해 보여줍니다(io.js 의 aggregateGroup).
 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  /* 시트 설명 — "각 시트를 구분하여 파악"할 수 있도록 화면에 함께 표시합니다.
     group 은 시트를 4개 묶음으로 나눠 목록에 소제목으로 보여주기 위한 값입니다. */
  var SHEET_GROUPS = [
    { key: 'trade', title: '거래표', unit: '백만원', hint: '실제로 오간 금액' },
    { key: 'input', title: '투입계수', unit: '배', hint: '생산 1단위당 투입 비율' },
    { key: 'effect', title: '유발계수', unit: '배', hint: '최종수요 1단위당 파급효과' },
    { key: 'final', title: '최종수요 항목별 유발액', unit: '백만원', hint: '소비·투자·수출별 실적' }
  ];

  var SHEET_META = {
    'A표_총거래표(생산자)': {
      group: 'trade',
      desc: '국산과 수입을 합친 총거래표입니다. 행은 투입(사서 쓴 것), 열은 산출(팔린 것)을 뜻하며, 행·열이 같은 부문이면 같은 산업끼리의 거래입니다.'
    },
    'A표_수입거래표(생산자)': {
      group: 'trade',
      desc: '총거래 중 수입품만 따로 떼어낸 표입니다. 국내 생산 파급에서 빠지는 부분, 즉 해외로 새는 정도를 보여줍니다.'
    },
    'A표_국산거래표(생산자)': {
      group: 'trade',
      desc: '총거래 중 국산품만 남긴 표입니다. 이 앱이 쓰는 국산 생산유발계수가 바로 이 표에서 계산됩니다.'
    },
    '총투입계수(A)': {
      group: 'input',
      desc: '거래 금액을 산출액으로 나눈 비율입니다. "1단위를 만들려면 무엇을 얼마나 넣어야 하는가"를 보여줍니다.'
    },
    '수입투입계수(Am)': {
      group: 'input',
      desc: '투입 비율 중 수입품이 차지하는 몫입니다. 값이 크면 그 산업이 수입 원료에 더 의존한다는 뜻입니다.'
    },
    '국산투입계수(Ad)': {
      group: 'input',
      desc: '투입 비율 중 국산품이 차지하는 몫입니다. 국산 생산유발계수를 계산하는 핵심 행렬입니다.'
    },
    '생산유발계수': {
      group: 'effect',
      accent: true,
      desc: '이 앱의 분석 기준표입니다. 어떤 산업에 최종수요 1단위가 생길 때 각 산업의 생산이 몇 단위 늘어나는지를 담고 있습니다. STEP 3의 후방·전방연쇄 목록이 모두 이 표에서 나옵니다.'
    },
    '수입유발계수': {
      group: 'effect',
      desc: '최종수요 1단위가 국내 생산 대신 수입을 얼마나 불러오는지 보여줍니다. 국내 파급이 해외로 새는 정도입니다.'
    },
    '부가가치유발계수': {
      group: 'effect',
      desc: '최종수요 1단위가 임금·영업잉여 같은 부가가치를 얼마나 만들어내는지 보여줍니다. 생산이 늘어도 부가가치가 적으면 남는 것이 적습니다.'
    },
    '최종수요 항목별 생산유발액': {
      group: 'final',
      desc: '민간소비·정부소비·투자·수출 등 최종수요 항목별로 생산을 얼마나 유발했는지 정리한 표입니다.'
    },
    '최종수요 항목별 수입유발액': {
      group: 'final',
      desc: '최종수요 항목별로 수입을 얼마나 유발했는지 정리한 표입니다.'
    },
    '최종수요 항목별 부가가치유발액': {
      group: 'final',
      desc: '최종수요 항목별로 부가가치를 얼마나 유발했는지 정리한 표입니다.'
    }
  };

  var LEVEL_LABEL = { large: '대분류', medium: '중분류', small: '소분류', basic: '기본부문' };
  var CSV_LEVEL = { large: '대분류', medium: '중분류', small: '소분류', basic: '기본부문' };

  var state = {
    manifest: null,
    tree: null,
    units: {},          // level → dataset 정의
    dataCache: {},      // level → IO.parse 결과 (행렬만 보관)
    level: null,        // 현재 분류 단위
    dataset: null,
    workbook: null,     // 현재 단위의 워크북 (STEP 2용)
    data: null,
    companies: [],
    currentSheet: null,
    selectedIndex: -1,
    selectedRowCode: null,
    drill: null
  };

  /* ── 유틸 ─────────────────────────────────────────── */
  function fmtCoef(v) {
    if (!isFinite(v)) return '—';
    return Number(v).toFixed(4);
  }

  function fmtNum(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v !== 'number') return String(v);
    var abs = Math.abs(v);
    if (abs >= 1000) return v.toLocaleString('ko-KR', { maximumFractionDigits: 0 });
    if (abs >= 1) return v.toLocaleString('ko-KR', { maximumFractionDigits: 3 });
    return v.toLocaleString('ko-KR', { maximumFractionDigits: 6 });
  }

  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }

  function setStatus(node, text, kind) {
    if (!node) return;
    node.textContent = text;
    node.classList.remove('is-busy', 'is-error', 'is-ok');
    if (kind) node.classList.add(kind);
  }

  function unlock(sectionId) {
    var section = $(sectionId);
    if (section) section.classList.remove('is-locked');
  }

  function pct(v) {
    if (!isFinite(v)) return '—';
    return (v * 100).toFixed(1) + '%';
  }

  /* ── CSV ─────────────────────────────────────────── */
  function parseCSV(text) {
    var rows = [];
    var row = [];
    var field = '';
    var quoted = false;

    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (quoted) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else quoted = false;
        } else field += ch;
      } else if (ch === '"') {
        quoted = true;
      } else if (ch === ',') {
        row.push(field); field = '';
      } else if (ch === '\n') {
        row.push(field); field = '';
        rows.push(row); row = [];
      } else if (ch !== '\r') {
        field += ch;
      }
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }

    if (!rows.length) return [];
    var head = rows.shift().map(function (h) { return h.trim(); });
    return rows.filter(function (r) {
      return r.some(function (c) { return c && c.trim() !== ''; });
    }).map(function (r) {
      var obj = {};
      head.forEach(function (h, idx) { obj[h] = (r[idx] || '').trim(); });
      return obj;
    });
  }

  /* ── 1. 초기화 ────────────────────────────────────── */
  async function init() {
    renderCompanyPlaceholder();
    $('load-btn').disabled = true;
    setUnitButtonsEnabled(false);

    try {
      var res = await fetch('./data/manifest.json', { cache: 'no-cache' });
      if (!res.ok) throw new Error('manifest.json을 읽지 못했습니다 (HTTP ' + res.status + ')');
      state.manifest = await res.json();
      (state.manifest.datasets || []).forEach(function (d) {
        if (d.enabled) state.units[d.id] = d;
      });
      renderDatasetPlan();
    } catch (e) {
      setStatus($('load-status'), '설정 파일을 읽지 못했습니다: ' + e.message, 'is-error');
      $('load-btn').disabled = false;
      return;
    }

    try {
      var cres = await fetch(state.manifest.companies.file, { cache: 'no-cache' });
      if (cres.ok) state.companies = parseCSV(await cres.text());
    } catch (e) {
      state.companies = [];
    }

    // 분류 계층표 — 실패해도 1단계 기능은 그대로 동작해야 합니다.
    try {
      var treeCfg = state.manifest.sectorTree;
      if (treeCfg) {
        var tres = await fetch(treeCfg.file, { cache: 'no-cache' });
        if (tres.ok) {
          var tbuf = await tres.arrayBuffer();
          var twb = XLSX.read(tbuf, { type: 'array' });
          var tws = twb.Sheets[treeCfg.sheet] || twb.Sheets[twb.SheetNames[0]];
          var tgrid = XLSX.utils.sheet_to_json(tws, { header: 1, raw: true, defval: null, blankrows: false });
          state.tree = IO_TREE.parse(tgrid);
          renderTreeSummary();
        }
      }
    } catch (e) {
      state.tree = null;
      setStatus($('load-status'), '분류 계층표를 읽지 못했습니다: ' + e.message + ' (드릴다운을 사용할 수 없습니다)', 'is-error');
    }

    setUnitButtonsEnabled(true);
    syncUnitButtons();
    renderMetaList();          // 계층표를 읽은 뒤 다시 그려 '계층 관계' 값을 채웁니다.
    initFinder();              // 검색창도 분류표 준비 상태에 맞춰 잠금/해제합니다.
    $('load-btn').disabled = false;
    $('load-btn').dataset.ready = 'true';
    $('load-btn').textContent = '자료 불러오기';
  }

  function setUnitButtonsEnabled(on) {
    Array.prototype.forEach.call($('unit-switch').querySelectorAll('button'), function (b) {
      var level = b.dataset.unit;
      b.disabled = !on || !state.units[level];
    });
  }

  function syncUnitButtons() {
    Array.prototype.forEach.call($('unit-switch').querySelectorAll('button'), function (b) {
      var active = b.dataset.unit === state.level;
      b.classList.toggle('is-active', active);
      b.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
  }

  function renderDatasetPlan() {
    var first = state.units.large || (state.manifest.datasets || [])[0];
    if (!first) return;
    state.level = state.level || first.id;
    state.dataset = state.units[state.level] || first;
    renderMetaList();
  }

  function renderMetaList() {
    var meta = $('dataset-meta');
    var m = state.manifest.meta;
    var target = state.dataset;

    var rows = [
      ['기준연도', m.year + '년 (연장표)'],
      ['가격평가', m.price],
      ['분류', m.classification],
      ['사용 계수', '국산 생산유발계수'],
      ['분류 단위', target.label + ' (' + target.sectorCount + '부문)'],
      ['계층 관계', state.tree
        ? '대분류 ' + state.tree.counts.large + ' · 중분류 ' + state.tree.counts.medium +
          ' · 소분류 ' + state.tree.counts.small + ' · 기본부문 ' + state.tree.counts.basic
        : '확인 중']
    ];
    rows.push(['자료 파일', target.originalFile]);

    meta.replaceChildren.apply(meta, rows.map(function (pair) {
      var div = el('div');
      div.appendChild(el('dt', null, pair[0]));
      div.appendChild(el('dd', null, pair[1]));
      return div;
    }));
    meta.hidden = false;
  }

  function renderTreeSummary() {
    var t = state.tree;
    if (!t) return;
    $('unit-hint').textContent =
      '단위를 고르면 그 단위의 파일을 새로 내려받습니다. ' +
      '분류표에서 확인한 부문 수는 대분류 ' + t.counts.large + ' · 중분류 ' + t.counts.medium +
      ' · 소분류 ' + t.counts.small + ' · 기본부문 ' + t.counts.basic + '개입니다.' +
      ' 기본부문(380부문)은 파일이 커서 시간이 걸립니다.';
    initFinder();
  }

  /* ── 1-b. 이름으로 분류 찾기 ──────────────────────────
   *
   * 사용자는 "자동차"가 대분류·중분류·소분류·기본부문 중 어디에 속하는지 모릅니다.
   * 분류표(sector-tree.xlsx)는 초기화 때 이미 4단계 661개 부문을 메모리에 올려두므로
   * **검색 자체는 추가 다운로드가 없습니다.** 결과를 눌렀을 때만 그 단위 파일을 받습니다.
   */
  var FINDER_LIMIT = 24;
  var FINDER_DEPTH = { large: 0, medium: 1, small: 2, basic: 3 };

  function finderTotalCount() {
    var t = state.tree;
    if (!t) return 0;
    return t.counts.large + t.counts.medium + t.counts.small + t.counts.basic;
  }

  function initFinder() {
    var input = $('finder-search');
    if (!input) return;
    if (!state.tree) {
      input.disabled = true;
      $('finder-note').textContent = '분류표를 읽지 못해 이름 검색을 사용할 수 없습니다.';
      return;
    }
    input.disabled = false;
    $('finder-note').textContent =
      '대분류·중분류·소분류·기본부문 ' + finderTotalCount() +
      '개 부문에서 이름이나 코드로 찾습니다.';
  }

  /** 이름·코드로 4단계 전체를 검색합니다. 정확히 일치 → 앞부분 일치 → 포함 순. */
  function findSectors(query) {
    var t = state.tree;
    if (!t) return [];
    var needle = query.toLowerCase();
    var out = [];

    IO_TREE.LEVELS.forEach(function (level) {
      t.nodes[level].forEach(function (node) {
        var name = node.name.toLowerCase();
        var code = node.code.toLowerCase();
        var score;
        if (name === needle || code === needle) score = 0;
        else if (name.indexOf(needle) === 0) score = 1;
        else if (name.indexOf(needle) >= 0 || code.indexOf(needle) >= 0) score = 2;
        else return;
        out.push({ node: node, score: score, depth: FINDER_DEPTH[level] });
      });
    });

    // 같은 점수면 상위 분류(대분류)를 먼저 보여 "어디에 속하는지"가 바로 보이게 합니다.
    out.sort(function (a, b) {
      if (a.score !== b.score) return a.score - b.score;
      if (a.depth !== b.depth) return a.depth - b.depth;
      return a.node.name.length - b.node.name.length;
    });
    return out;
  }

  function renderFinderResults() {
    var input = $('finder-search');
    var list = $('finder-results');
    var note = $('finder-note');
    if (!input || !list) return;

    var q = input.value.trim();
    if (!q) {
      list.hidden = true;
      list.replaceChildren();
      initFinder();
      return;
    }

    var hits = findSectors(q);
    if (!hits.length) {
      list.hidden = true;
      list.replaceChildren();
      note.textContent = '"' + q + '"에 해당하는 부문이 없습니다. 다른 이름으로 찾아보세요.';
      return;
    }

    var shown = hits.slice(0, FINDER_LIMIT);
    note.textContent = shown.length < hits.length
      ? hits.length + '개가 검색되어 앞 ' + shown.length + '개만 표시합니다. 검색어를 더 넣어 좁혀 보세요.'
      : hits.length + '개 부문이 검색되었습니다. 누르면 그 분류 자료를 바로 불러옵니다.';

    list.replaceChildren.apply(list, shown.map(function (hit) {
      var node = hit.node;
      var li = document.createElement('li');
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'finder-hit';
      btn.title = node.label + ' ' + node.code + ' ' + node.name + ' — 누르면 이 분류 자료를 불러옵니다';

      btn.appendChild(el('span', 'finder-level is-' + node.level, node.label));
      btn.appendChild(el('span', 'finder-code', node.code));
      btn.appendChild(el('span', 'finder-name', node.name));

      // 상위 분류 경로 — "대분류 제조업 › 중분류 …" 처럼 어디에 속하는지 보여줍니다.
      var path = state.tree.lineage(node.level, node.code)
        .filter(function (n) { return n.code !== node.code; })
        .map(function (n) { return n.name; });
      if (path.length) btn.appendChild(el('span', 'finder-path', path.join(' › ')));

      btn.addEventListener('click', function () { chooseFromFinder(node); });
      li.appendChild(btn);
      return li;
    }));
    list.hidden = false;
  }

  /** 검색 결과를 누르면 그 분류 단위로 전환하고 해당 부문을 선택합니다. */
  function chooseFromFinder(node) {
    if (!state.units[node.level]) {
      setStatus($('load-status'), node.label + ' 자료를 사용할 수 없습니다.', 'is-error');
      return;
    }
    $('finder-results').hidden = true;
    jumpToUnit(node.code, node.level, true);
  }

  /* ── 2. 자료 불러오기 ──────────────────────────────── */
  async function loadUnit(level, opts) {
    opts = opts || {};
    var target = state.units[level];
    if (!target) return;

    var btn = $('load-btn');
    var status = $('load-status');
    var prevSelection = opts.keepSelection === false ? null : currentSectorCode();

    state.level = level;
    state.dataset = target;
    syncUnitButtons();
    renderMetaList();

    btn.disabled = true;
    setUnitButtonsEnabled(false);
    setStatus(status, target.label + ' 파일을 내려받는 중…', 'is-busy');

    try {
      var data = state.dataCache[level];
      var wb;

      if (data) {
        // 행렬은 이미 계산해 두었습니다. STEP 2를 위해 워크북만 다시 받습니다.
        setStatus(status, target.label + ' 워크북을 다시 여는 중…', 'is-busy');
        var res0 = await fetch(target.file);
        if (!res0.ok) throw new Error('파일을 내려받지 못했습니다 (HTTP ' + res0.status + ')');
        wb = XLSX.read(await res0.arrayBuffer(), { type: 'array' });
      } else {
        var res = await fetch(target.file);
        if (!res.ok) throw new Error('파일을 내려받지 못했습니다 (HTTP ' + res.status + ')');
        var buf = await res.arrayBuffer();

        setStatus(status, 'Excel 파일을 해석하는 중… (' + Math.round(buf.byteLength / 1024) + ' KB)', 'is-busy');
        wb = XLSX.read(buf, { type: 'array' });

        var sheet = wb.Sheets[target.sheet];
        if (!sheet) throw new Error('「' + target.sheet + '」 시트를 찾을 수 없습니다.');
        var grid = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
        data = IO.parse(grid, target.label);
        state.dataCache[level] = data;
      }

      state.workbook = wb;
      state.data = data;
      state.selectedRowCode = null;

      buildSheetList();
      buildSectorSelect();
      renderValidation();

      setStatus(
        status,
        target.label + ' ' + data.n + '부문 자료를 불러왔습니다. 아래에서 산업을 선택하세요.',
        'is-ok'
      );

      btn.textContent = '다시 불러오기';
      btn.disabled = false;
      setUnitButtonsEnabled(true);

      unlock('explore');
      unlock('analysis');
      unlock('companies-anchor');

      $('sheet-search').disabled = false;
      $('sector-search').disabled = false;
      $('sector-select').disabled = false;
      $('sheet-search').value = '';
      $('sector-search').value = '';

      var firstSheet = (wb.SheetNames || []).filter(function (n) { return n === '생산유발계수'; })[0]
        || wb.SheetNames[0];
      showSheet(firstSheet);

      clearDrill();

      var idx = 0;
      if (pendingSelection && pendingSelection.level === level) {
        var pendingIdx = data.sectors.findIndex(function (s) { return s.code === pendingSelection.code; });
        if (pendingIdx >= 0) idx = pendingIdx;
        pendingSelection = null;
      } else if (prevSelection) {
        var mapped = mapSelectionCode(prevSelection.fromLevel, prevSelection.code, level);
        if (mapped) {
          var found = data.sectors.findIndex(function (s) { return s.code === mapped; });
          if (found >= 0) idx = found;
        }
      }
      $('sector-select').value = String(idx);
      selectSector(idx);

      /* STEP 1 검색에서 넘어온 경우에만 분석 영역으로 이동합니다.
         (분류 단위 버튼으로 직접 바꿀 때는 화면이 튀지 않도록 그대로 둡니다.) */
      if (state.keepScrollTo) {
        state.keepScrollTo = false;
        $('analysis').scrollIntoView({ block: 'start' });
      }
    } catch (e) {
      setStatus(status, '불러오기 실패: ' + e.message, 'is-error');
      btn.disabled = false;
      btn.textContent = '다시 시도';
      setUnitButtonsEnabled(true);
    }
  }

  function currentSectorCode() {
    var d = state.data;
    if (!d || state.selectedIndex < 0 || !d.sectors[state.selectedIndex]) return null;
    return { fromLevel: state.level, code: d.sectors[state.selectedIndex].code };
  }

  /**
   * 분류 단위를 바꿔도 같은 계통을 유지합니다.
   *   위로 갈 때(중분류→대분류): 조상 코드
   *   아래로 갈 때(대분류→중분류): 첫 번째 자손 코드
   */
  function mapSelectionCode(fromLevel, code, toLevel) {
    var tree = state.tree;
    if (!tree || fromLevel === toLevel) return code;

    var path = tree.lineage(fromLevel, code);
    if (!path.length) return null;

    for (var i = 0; i < path.length; i++) {
      if (path[i].level === toLevel) return path[i].code;
    }

    // toLevel 이 더 아래 — 첫 자손을 따라 내려갑니다.
    var cur = path[path.length - 1];
    var guard = 0;
    while (cur && cur.level !== toLevel && guard++ < 8) {
      var kids = tree.childrenOf(cur.level, cur.code);
      if (!kids.length) return null;
      cur = kids[0];
    }
    return cur && cur.level === toLevel ? cur.code : null;
  }

  /* ── 3. 검증 결과 ─────────────────────────────────── */
  function renderValidation() {
    var d = state.data;
    var box = $('validation-box');
    var body = $('validation-body');
    if (!d) { box.hidden = true; return; }

    var v = d.validation;
    var residualOk = v.maxResidual < 1e-6;
    var sumOk = Math.abs(v.sumInfluence - v.expectedSum) < 0.01 &&
                Math.abs(v.sumSensitivity - v.expectedSum) < 0.01;

    var items = [
      [residualOk, '계수 재계산 검증', '최대 잔차 ' + v.maxResidual.toExponential(2) + ' (허용 1e-6)'],
      [sumOk, '영향력계수 합 = 부문 수', fmtNum(v.sumInfluence) + ' / 기대값 ' + v.expectedSum],
      [sumOk, '감응도계수 합 = 부문 수', fmtNum(v.sumSensitivity) + ' / 기대값 ' + v.expectedSum],
      [d.columnOrderMatches, '행·열 부문 순서 일치', d.columnOrderMatches ? '일치' : '불일치 — 코드 기준으로 대응'],
      [true, '부문 수', d.n + '개 (파일 표기 ' + (state.dataset ? state.dataset.sectorCount : d.n) + '개)']
    ];

    if (state.tree) {
      var treeCount = state.tree.counts[state.level];
      var match = treeCount === d.n;
      items.push([match, '분류표 계층과 일치',
        match ? '분류표 ' + treeCount + '개 = 파일 ' + d.n + '개'
              : '분류표 ' + treeCount + '개 ≠ 파일 ' + d.n + '개']);
    }

    var list = el('ul', 'check-list');
    items.forEach(function (item) {
      var li = el('li');
      li.appendChild(el('span', item[0] ? 'ok' : 'warn', item[0] ? '✓' : '!'));
      var span = el('span');
      span.appendChild(el('strong', null, item[1] + ' — '));
      span.appendChild(document.createTextNode(item[2]));
      li.appendChild(span);
      list.appendChild(li);
    });

    body.replaceChildren(list);
    box.hidden = false;
  }

  /* ── 4. 시트 탐색 ─────────────────────────────────── */
  function buildSheetList() {
    var wb = state.workbook;
    var names = wb.SheetNames || [];
    var list = $('sheet-list');
    var nodes = [];

    SHEET_GROUPS.forEach(function (group) {
      var members = names.filter(function (n) {
        var m = SHEET_META[n];
        return m && m.group === group.key;
      });
      if (!members.length) return;

      var head = el('li', 'sheet-group');
      head.appendChild(el('span', 'sheet-group-title', group.title));
      head.appendChild(el('span', 'sheet-group-unit', group.unit));
      head.appendChild(el('span', 'sheet-group-hint', group.hint));
      nodes.push(head);

      members.forEach(function (name) {
        var meta = SHEET_META[name] || {};
        var li = el('li');
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.dataset.sheet = name;
        btn.appendChild(document.createTextNode(name));
        if (meta.accent) {
          btn.appendChild(document.createTextNode(' '));
          btn.appendChild(el('span', 'sheet-badge', '분석 기준'));
        }
        btn.addEventListener('click', function () { showSheet(name); });
        li.appendChild(btn);
        nodes.push(li);
      });
    });

    list.replaceChildren.apply(list, nodes);
  }

  function metaOf(name) {
    return SHEET_META[name] || { desc: '이 시트에 대한 설명은 준비되지 않았습니다.', group: '' };
  }

  function unitOf(name) {
    var g = SHEET_GROUPS.filter(function (x) { return x.key === metaOf(name).group; })[0];
    return g ? g.unit : '—';
  }

  function showSheet(name) {
    if (!state.workbook) return;
    var ws = state.workbook.Sheets[name];
    if (!ws) return;
    state.currentSheet = name;

    Array.prototype.forEach.call($('sheet-list').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-current', b.dataset.sheet === name ? 'true' : 'false');
    });

    $('sheet-current').textContent = name;

    var box = $('sheet-note');
    box.replaceChildren();
    box.appendChild(el('span', 'unit-chip', '단위 ' + unitOf(name)));
    box.appendChild(document.createTextNode(' ' + metaOf(name).desc));

    $('sheet-search').value = '';
    renderSheetTable();
  }

  /* 미리보기는 원자료를 "그대로" 보여주는 것이 목적이므로 행을 잘라내지 않습니다.
     40행에서 자르면 중분류(90행)·소분류(172행)·기본부문(387행)의 산업 대부분이
     화면에서 사라져 "산업을 구분하여 파악"할 수 없게 됩니다.
     안전 상한만 두고, 넘으면 캡션에 명시합니다. */
  var SHEET_ROW_LIMIT = 800;

  function renderSheetTable() {
    var name = state.currentSheet;
    if (!name) return;
    var ws = state.workbook.Sheets[name];
    var grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });

    var q = $('sheet-search').value.trim();
    var maxCols = Math.min(Math.max.apply(null, grid.map(function (r) { return r.length; })), 12);

    var rows = grid.map(function (row, i) { return { row: row, i: i }; });

    if (q) {
      rows = rows.filter(function (r) {
        return r.row.some(function (c) { return c !== null && String(c).indexOf(q) >= 0; });
      });
    }

    var truncated = rows.length > SHEET_ROW_LIMIT;
    var shown = truncated ? rows.slice(0, SHEET_ROW_LIMIT) : rows;

    var thead = $('sheet-thead');
    var headRow = el('tr');
    headRow.appendChild(el('th', null, '#'));
    for (var c = 0; c < maxCols; c++) {
      var label = String.fromCharCode(65 + c);
      headRow.appendChild(el('th', null, label));
    }
    thead.replaceChildren(headRow);

    var tbody = $('sheet-tbody');
    tbody.replaceChildren.apply(tbody, shown.map(function (item) {
      var tr = el('tr');
      var first = item.row[0] === null || item.row[0] === undefined ? '' : String(item.row[0]).trim();
      if (['9801', '9802', '9590', '9090'].indexOf(first) >= 0) tr.className = 'is-special';
      // STEP 3에서 고른 산업의 행은 다시 그려도 강조를 유지합니다.
      if (state.selectedRowCode && first === state.selectedRowCode) {
        tr.className = (tr.className ? tr.className + ' ' : '') + 'is-selected-row';
      }
      tr.appendChild(el('td', 'num', String(item.i + 1)));
      for (var k = 0; k < maxCols; k++) {
        var val = item.row[k];
        var td = el('td', typeof val === 'number' ? 'num' : null);
        td.textContent = typeof val === 'number' ? fmtNum(val) : (val === null || val === undefined ? '' : String(val));
        tr.appendChild(td);
      }
      return tr;
    }));

    var caption = name +
      ' · ' + (state.data ? state.data.n + '부문' : '') +
      ' · 표시 ' + shown.length + '행 / 전체 ' + grid.length + '행';
    if (q) caption += ' (검색어 "' + q + '" · ' + rows.length + '행 일치)';
    if (truncated) caption += ' · 앞 ' + SHEET_ROW_LIMIT + '행만 표시';
    caption += ' · 앞 ' + maxCols + '열 표시';
    $('sheet-caption').textContent = caption;
  }

  /* ── 5. 산업 선택 ─────────────────────────────────── */
  function buildSectorSelect() {
    var d = state.data;
    var sel = $('sector-select');
    sel.replaceChildren.apply(sel, d.sectors.map(function (s, i) {
      var opt = document.createElement('option');
      opt.value = String(i);
      opt.textContent = s.code + ' · ' + s.name;
      return opt;
    }));
    $('picker-note').textContent =
      d.label + ' ' + d.n + '개 부문 · 코드 또는 이름으로 검색할 수 있습니다.';
  }

  function filterSectors() {
    var d = state.data;
    var q = $('sector-search').value.trim().toLowerCase();
    var sel = $('sector-select');
    var list = d.sectors.filter(function (s) {
      if (!q) return true;
      return (s.code + ' ' + s.name).toLowerCase().indexOf(q) >= 0;
    });

    /* 검색어를 지운 경우 — 전체 목록을 '안내 항목 없이' 그대로 되돌리고
       현재 분석 중인 산업이 선택된 상태로 남깁니다. */
    if (!q) {
      sel.replaceChildren.apply(sel, d.sectors.map(function (s, i) {
        var opt = document.createElement('option');
        opt.value = String(i);
        opt.textContent = s.code + ' · ' + s.name;
        return opt;
      }));
      if (state.selectedIndex >= 0) sel.value = String(state.selectedIndex);
      $('picker-note').textContent =
        d.label + ' ' + d.n + '개 부문 · 코드 또는 이름으로 검색할 수 있습니다.';
      return;
    }

    if (!list.length) {
      sel.replaceChildren();
      var none = document.createElement('option');
      none.value = '';
      none.textContent = '검색 결과 없음';
      sel.appendChild(none);
      $('picker-note').textContent = '"' + q + '"에 해당하는 산업이 없습니다.';
      return;
    }

    /* 검색 결과가 하나뿐이면 고를 것이 없으므로 **바로 적용**합니다.
       (예전에는 목록에 1개만 남아도 value 가 그 항목으로 바뀔 뿐
        change 이벤트가 나지 않아 분석이 갱신되지 않았습니다.) */
    if (list.length === 1) {
      var onlyIdx = d.sectors.indexOf(list[0]);
      sel.replaceChildren.apply(sel, list.map(function (s) {
        var opt = document.createElement('option');
        opt.value = String(d.sectors.indexOf(s));
        opt.textContent = s.code + ' · ' + s.name;
        return opt;
      }));
      sel.value = String(onlyIdx);
      $('picker-note').textContent =
        '"' + q + '" 검색 결과가 1개라 ' + list[0].name + '을(를) 바로 적용했습니다.';
      selectSector(onlyIdx);
      return;
    }

    /* 결과가 2개 이상이면 **빈 안내 항목**을 첫 줄에 둡니다.
       이렇게 해야 사용자가 실제 항목을 눌렀을 때 value 가 빈 값에서 바뀌면서
       change 가 확실히 발생합니다. (첫 항목이 이미 선택돼 있으면
        그 항목을 눌러도 change 가 나지 않아 "클릭해도 반영 안 됨"이 됩니다.) */
    sel.replaceChildren.apply(sel, list.map(function (s) {
      var opt = document.createElement('option');
      opt.value = String(d.sectors.indexOf(s));
      opt.textContent = s.code + ' · ' + s.name;
      return opt;
    }));

    var hint = document.createElement('option');
    hint.value = '';
    hint.textContent = '— 산업을 고르세요 (' + list.length + '개) —';
    sel.insertBefore(hint, sel.firstChild);
    sel.value = '';

    $('picker-note').textContent =
      list.length + '개 산업이 검색되었습니다. 목록에서 하나를 고르면 분석이 시작됩니다.';
  }

  function selectSector(index) {
    var d = state.data;
    if (!d || !d.sectors[index]) return;
    state.selectedIndex = index;

    var sector = d.sectors[index];
    var influence = d.influence[index];
    var sensitivity = d.sensitivity[index];
    var influenceRank = IO.rankOf(d.influence, index);
    var sensitivityRank = IO.rankOf(d.sensitivity, index);

    $('result').hidden = false;
    $('result-code').textContent = sector.code;
    $('result-name').textContent = sector.name;
    $('result-sub').textContent = buildResultSub(sector, d);

    renderMetricCards(influence, influenceRank, sensitivity, sensitivityRank);

    var back = IO.breakdown(d, index, 'back', 5);
    var fwd = IO.breakdown(d, index, 'fwd', 5);
    renderTrack('track-back-list', back, 'track-back-sum');
    renderTrack('track-fwd-list', fwd, 'track-fwd-sum');

    $('fn-count').textContent = d.n - 1;

    renderCompanies(sector);
    $('analysis-note').textContent =
      '위 목록은 생산유발계수 크기 순서이며, 자기 자신은 제외했습니다. ' +
      '항목을 클릭하면 그 산업으로 분석 대상을 바꿉니다.';

    highlightRawRow(sector);
    prepareDrill(sector);
  }

  function buildResultSub(sector, d) {
    var parts = [d.label + ' · ' + d.n + '부문 중 비교 · 국산 생산유발계수 기준'];
    if (state.tree) {
      var line = state.tree.lineage(state.level, sector.code)
        .filter(function (n) { return n.code !== sector.code; })
        .map(function (n) { return n.label + ' ' + n.name; });
      if (line.length) parts.push('상위 분류: ' + line.join(' › '));
    }
    return parts.join(' · ');
  }

  /** STEP 2 원자료 표에서 선택 산업의 행을 강조합니다. */
  function highlightRawRow(sector) {
    if (!state.workbook || !sector) return;
    var rows = $('sheet-tbody').querySelectorAll('tr');
    var hit = null;
    Array.prototype.forEach.call(rows, function (tr) {
      tr.classList.remove('is-selected-row');
      var first = tr.cells[1];
      if (first && first.textContent.trim() === sector.code) hit = tr;
    });
    if (hit) hit.classList.add('is-selected-row');
    state.selectedRowCode = sector.code;
  }

  /** STEP 2를 열고 선택 산업의 행으로 이동합니다. */
  function gotoRaw() {
    var d = state.data;
    if (!d || state.selectedIndex < 0) return;
    var sector = d.sectors[state.selectedIndex];

    if ($('sheet-current').textContent !== '생산유발계수') {
      showSheet('생산유발계수');
    }
    highlightRawRow(sector);

    var hit = $('sheet-tbody').querySelector('.is-selected-row');
    if (hit && hit.scrollIntoView) {
      hit.scrollIntoView({ block: 'center' });
    }
    $('explore').scrollIntoView({ block: 'start' });
  }

  function renderMetricCards(influence, influenceRank, sensitivity, sensitivityRank) {
    var box = $('metric-cards');
    var cards = [
      {
        label: '영향력계수 (후방연쇄)',
        value: fmtCoef(influence),
        meta: '전체 ' + state.data.n + '부문 중 ' + influenceRank + '위',
        high: influence > 1
      },
      {
        label: '감응도계수 (전방연쇄)',
        value: fmtCoef(sensitivity),
        meta: '전체 ' + state.data.n + '부문 중 ' + sensitivityRank + '위',
        high: sensitivity > 1
      },
      {
        label: '판정',
        value: (influence > 1 ? '후방 강함' : '후방 약함') + ' · ' + (sensitivity > 1 ? '전방 강함' : '전방 약함'),
        meta: '1.0 초과 = 전체 평균보다 큼',
        high: influence > 1 || sensitivity > 1,
        small: true
      }
    ];

    box.replaceChildren.apply(box, cards.map(function (c) {
      var div = el('div', 'metric');
      div.appendChild(el('p', 'metric-label', c.label));
      var v = el('p', 'metric-value', c.value);
      if (c.small) v.style.fontSize = '1.02rem';
      div.appendChild(v);
      div.appendChild(el('p', 'metric-meta', c.meta));
      var badge = el('span', 'badge ' + (c.high ? 'badge-high' : 'badge-low'),
        c.high ? '평균 이상' : '평균 이하');
      if (c.small) badge.textContent = c.high ? '평균 초과 항목 있음' : '모두 평균 이하';
      div.appendChild(badge);
      return div;
    }));
  }

  function renderTrack(targetId, result, sumId, opts) {
    opts = opts || {};
    var items = result.items;
    var ol = $(targetId);
    var sumEl = $(sumId);

    if (!items.length) {
      var emptyLi = el('li');
      emptyLi.appendChild(el('span', 'rank-name', '표시할 항목이 없습니다.'));
      ol.replaceChildren(emptyLi);
      if (sumEl) sumEl.textContent = '';
      return;
    }

    var max = items[0].value || 1;

    ol.replaceChildren.apply(ol, items.map(function (item) {
      var li = document.createElement('li');

      var nameDiv = el('span', 'rank-name');
      nameDiv.appendChild(el('span', 'rank-code', item.code));

      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'rank-link';
      btn.textContent = item.name;
      btn.title = item.name + ' 부문을 분석 대상으로 선택';
      btn.addEventListener('click', function () {
        var idx = state.data.sectors.findIndex(function (s) { return s.code === item.code; });
        if (idx < 0) return;
        // 검색어로 목록이 좁혀져 있으면 선택 표시가 비어 보이므로, 전체 목록을 복원한 뒤 선택합니다.
        $('sector-search').value = '';
        buildSectorSelect();
        $('sector-select').value = String(idx);
        selectSector(idx);
        $('analysis').scrollIntoView({ block: 'start' });
      });
      nameDiv.appendChild(btn);

      var val = el('span', 'rank-val', fmtCoef(item.value));

      var bar = el('span', 'rank-bar');
      var fill = document.createElement('span');
      fill.style.width = Math.max(2, (item.value / max) * 100) + '%';
      bar.appendChild(fill);

      var share = el('span', 'rank-share');
      share.appendChild(el('span', 'share-bit', '비중 ' + pct(item.share)));
      share.appendChild(el('span', 'share-bit share-cum', '누적 ' + pct(item.cumulative)));

      li.appendChild(nameDiv);
      li.appendChild(bar);
      li.appendChild(val);
      li.appendChild(share);
      return li;
    }));

    if (sumEl) {
      var unitWord = opts.unitWord || '산업';
      var rest = result.otherCount > 0
        ? ' 나머지 ' + result.otherCount + '개 ' + unitWord + '이 ' + pct(1 - result.topShare) + '를 차지합니다.'
        : '';
      sumEl.replaceChildren();
      sumEl.appendChild(el('strong', null, '상위 5개 누적 ' + pct(result.topShare)));
      sumEl.appendChild(document.createTextNode(
        ' · ' + (opts.totalLabel || '자기 제외 파급 합') + ' ' + fmtCoef(result.total) + '배 기준.' + rest));
    }
  }

  /* ── 6. 드릴다운 (상위 분류 → 하위 분류 그룹 평균) ────── */
  function drillTargetLevel(level) {
    var ds = state.units[level];
    if (ds && ds.drillTo) return ds.drillTo;
    var i = IO_TREE.LEVELS.indexOf(level);
    return i >= 0 && i < IO_TREE.LEVELS.length - 1 ? IO_TREE.LEVELS[i + 1] : null;
  }

  function prepareDrill(sector) {
    var panel = $('drill');
    var tree = state.tree;
    var childLevel = drillTargetLevel(state.level);

    state.drill = null;
    $('drill-body').hidden = true;
    setStatus($('drill-status'), '', null);

    if (!tree || !childLevel || !state.units[childLevel]) {
      panel.hidden = true;
      return;
    }

    var kids = tree.childrenOf(state.level, sector.code);
    if (kids.length < 1) {
      panel.hidden = true;
      return;
    }

    var path = tree.lineage(state.level, sector.code).map(function (n) {
      return n.label + ' ' + n.name;
    });

    panel.hidden = false;
    $('drill-path').textContent = path.join(' › ');
    $('drill-title').textContent = state.units[childLevel].label + '로 내려보기';

    if (kids.length === 1) {
      $('drill-lead').textContent =
        '이 부문은 ' + state.units[childLevel].label + ' 1개로만 나뉩니다 (' +
        kids[0].code + ' ' + kids[0].name + '). 집계해도 같은 값이므로 참고용입니다.';
    } else {
      $('drill-lead').textContent =
        '이 부문은 ' + state.units[childLevel].label + ' ' + kids.length + '개로 나뉩니다 (' +
        kids.map(function (k) { return k.code + ' ' + k.name; }).join(', ') + '). ' +
        '이들을 한 그룹으로 보고 평균 계수로 집계합니다.';
    }

    $('drill-run').textContent = state.units[childLevel].label + ' ' + kids.length + '개 집계하기';
    $('drill-run').disabled = false;

    state.drill = {
      parentLevel: state.level,
      parentCode: sector.code,
      parentName: sector.name,
      childLevel: childLevel,
      children: kids
    };
  }

  function clearDrill() {
    state.drill = null;
    $('drill').hidden = true;
    $('drill-body').hidden = true;
    setStatus($('drill-status'), '', null);
  }

  async function runDrill() {
    var job = state.drill;
    if (!job) return;

    var statusEl = $('drill-status');
    var btn = $('drill-run');
    var childDs = state.units[job.childLevel];
    if (!childDs) { setStatus(statusEl, '하위 분류 자료를 사용할 수 없습니다.', 'is-error'); return; }

    btn.disabled = true;
    setStatus(statusEl, childDs.label + ' 자료를 준비하는 중…', 'is-busy');

    var childData = state.dataCache[job.childLevel];
    if (!childData) {
      try {
        var res = await fetch(childDs.file);
        if (!res.ok) throw new Error('HTTP ' + res.status);
        var buf = await res.arrayBuffer();
        setStatus(statusEl, childDs.label + ' 파일을 해석하는 중… (' +
          Math.round(buf.byteLength / 1024) + ' KB)', 'is-busy');
        var wb = XLSX.read(buf, { type: 'array' });
        var ws = wb.Sheets[childDs.sheet];
        if (!ws) throw new Error('「' + childDs.sheet + '」 시트를 찾을 수 없습니다.');
        var grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        childData = IO.parse(grid, childDs.label);
        state.dataCache[job.childLevel] = childData;
      } catch (e) {
        setStatus(statusEl, '하위 분류 자료를 읽지 못했습니다: ' + e.message, 'is-error');
        btn.disabled = false;
        return;
      }
    }

    // 그룹 구성 코드가 실제 파일에 있는지 확인
    var present = job.children.filter(function (k) { return childData.colIndexByCode[k.code] !== undefined; });
    if (present.length !== job.children.length) {
      setStatus(statusEl,
        '분류표의 하위 부문 ' + job.children.length + '개 중 ' + present.length + '개만 파일에서 찾았습니다.',
        'is-busy');
    }

    var codes = present.map(function (k) { return k.code; });
    var back, fwd;
    try {
      back = IO.aggregateGroup(childData, codes, 'back', 5);
      fwd = IO.aggregateGroup(childData, codes, 'fwd', 5);
    } catch (e) {
      setStatus(statusEl, '집계 실패: ' + e.message, 'is-error');
      btn.disabled = false;
      return;
    }

    $('drill-k').textContent = codes.length;

    renderTrack('drill-back-list', {
      items: back.items,
      total: back.externalTotal,
      topShare: back.topShare,
      otherCount: back.otherCount
    }, 'drill-back-sum', { totalLabel: '그룹 외부 파급 합', unitWord: childDs.label + ' 부문' });

    renderTrack('drill-fwd-list', {
      items: fwd.items,
      total: fwd.externalTotal,
      topShare: fwd.topShare,
      otherCount: fwd.otherCount
    }, 'drill-fwd-sum', { totalLabel: '그룹 외부 파급 합', unitWord: childDs.label + ' 부문' });

    renderInternal(back, fwd, job, childDs);

    var maxResidual = Math.max(back.residual, fwd.residual);
    setStatus(statusEl,
      childDs.label + ' ' + codes.length + '개 평균으로 집계했습니다. ' +
      '그룹 전체 파급(내부+외부) 교차검증 잔차 ' + maxResidual.toExponential(2) + '.',
      'is-ok');

    $('drill-body').hidden = false;
    btn.disabled = false;
  }

  /* ── 그룹 내부 연쇄: 3분해 (자기 자신 / 타 부문 거래 / 그룹 외부) ──
   *
   * 왜 나누는가
   *   생산유발계수의 대각값 b(m,m)은 정의상 거의 항상 1.0을 넘습니다. 그래서 그룹 내부
   *   합계를 그대로 보면 89~99%가 자기 자신이고, "부문끼리 서로 사고판다"와
   *   "각자 자기 생산분 1.0만 갖고 있다"가 같은 숫자로 보입니다(실측: 비철금속 1차제품
   *   내부 합의 99.4%가 대각). 그래서 자기 자신을 뺀 '타 부문 거래'를 함께 냅니다.
   *
   * 구간 경계의 근거
   *   2023년 연장표에서 만들 수 있는 그룹 281개(대→중 33 · 중→소 83 · 소→기본 165)를
   *   전부 계산해 사분위수를 경계로 삼았습니다. 절대 기준이 아니라 이 자료 안에서의
   *   상대 위치입니다.
   *     내부 비중(자기 포함)  p25 48.7 · 중앙값 55.3 · p75 62.2 · p90 69.6
   *     타 부문 거래          p25 0.00 · 중앙값 0.05 · p75 0.78 · p90 2.75
   *     집중도(k·HHI, 1=균등) p25 1.36 · 중앙값 1.71 · p75 2.12 · p90 2.96
   */
  var INTERNAL_BANDS = [
    { max: 0.49, cls: 'band-1', label: '낮음', note: '하위 25%' },
    { max: 0.55, cls: 'band-2', label: '보통', note: '중앙값 아래' },
    { max: 0.62, cls: 'band-2', label: '다소 높음', note: '중앙값~상위 25%' },
    { max: 0.70, cls: 'band-3', label: '높음', note: '상위 25%' },
    { max: Infinity, cls: 'band-4', label: '매우 높음', note: '상위 10%' }
  ];

  /* 타 부문 거래 구간 — 실측 누적분포로 경계를 잡았습니다.
       0%          : 13개 그룹 (부문끼리 거래가 계수에 없음)
       0~0.5%      : 196/281 = 69.8%   → 하위 70%
       0.5~2%      :  46/281 = 16.4%   → 상위 14~30%
       2~5%        :  26/281 =  9.3%   → 상위 5~14%
       5% 이상     :  13/281 =  4.6%   → 상위 5% 이내
     값 자체가 작은 것은 계산 오류가 아니라 계수의 성질입니다(도움말 3-1 참고). */
  var OFF_BANDS = [
    { max: 0.0000000001, cls: 'band-0', label: '없음', note: '부문끼리 거래가 잡히지 않음' },
    { max: 0.005, cls: 'band-1', label: '미미', note: '하위 70% 수준' },
    { max: 0.02, cls: 'band-2', label: '있음', note: '상위 14~30%' },
    { max: 0.05, cls: 'band-3', label: '뚜렷함', note: '상위 5~14%' },
    { max: Infinity, cls: 'band-4', label: '매우 뚜렷', note: '상위 5% 이내' }
  ];

  function pickBand(bands, value) {
    for (var i = 0; i < bands.length; i++) {
      if (value < bands[i].max) return bands[i];
    }
    return bands[bands.length - 1];
  }

  function bandNode(band) {
    var wrap = el('span');
    wrap.appendChild(el('span', 'band-pill ' + band.cls, band.label));
    wrap.appendChild(document.createTextNode(band.note));
    return wrap;
  }

  function legendBit(cls, label, value) {
    var s = el('span', cls);
    s.appendChild(el('i'));
    s.appendChild(document.createTextNode(label + ' ' + value));
    return s;
  }

  function concentrationLabel(c) {
    if (!isFinite(c)) return '—';
    if (c < 1.5) return '고르게 퍼짐';
    if (c < 2.5) return '중간';
    return '한 곳에 집중';
  }

  /* 받침 유무에 따라 조사를 고릅니다 — "열간압연강재이" 같은 어색한 문장을 막습니다.
     한글 음절은 (코드 − 0xAC00) % 28 이 종성 번호이고, 0 이면 받침이 없습니다. */
  function josa(word, withFinal, withoutFinal) {
    var s = String(word || '');
    if (!s.length) return withFinal;
    var code = s.charCodeAt(s.length - 1);
    if (code < 0xac00 || code > 0xd7a3) return withFinal;
    return (code - 0xac00) % 28 > 0 ? withFinal : withoutFinal;
  }

  function renderInternal(back, fwd, job, childDs) {
    var body = $('drill-internal-body');
    var k = back.size;
    var sh = back.shares;
    var hasOff = back.offTotal > 0;
    var topOff = hasOff ? back.offTopMember : null;
    var conc = back.offConcentration;
    var topExt = (back.items && back.items.length) ? back.items[0] : null;
    var bandA = pickBand(INTERNAL_BANDS, sh.internal);
    var bandB = pickBand(OFF_BANDS, hasOff ? sh.off : 0);

    function t(text) { return document.createTextNode(text); }
    function b(text) { return el('strong', null, text); }

    // ── 이 절이 무엇을 보여주는지 ──
    $('drill-internal-lead').textContent =
      job.parentName + ' 안의 ' + childDs.label + ' ' + k + '개를 한 묶음으로 보고, ' +
      '이 그룹에 최종수요가 생길 때의 파급(후방)이 어디로 흘러가는지 세 갈래로 나눴습니다. ' +
      '아래 두 지표는 분모가 같아서(그룹 전체 파급) 서로 더하면 100%가 됩니다.';
    $('help-k').textContent = k;

    /* ── 지표 A: 내부 비중 (자기 자신 포함) ── */
    $('drill-m-internal').textContent = pct(sh.internal);
    $('drill-m-internal-band').replaceChildren(bandNode(bandA));

    $('bar-diag').style.width = (sh.diag * 100).toFixed(2) + '%';
    $('bar-off').style.width = (sh.off * 100).toFixed(2) + '%';
    $('bar-ext').style.width = (sh.external * 100).toFixed(2) + '%';

    var legend = $('drill-m-legend');
    legend.replaceChildren.apply(legend, [
      legendBit('lg-diag', '자기 자신', pct(sh.diag)),
      legendBit('lg-off', '타 부문 거래', pct(sh.off)),
      legendBit('lg-ext', '그룹 외부', pct(sh.external))
    ]);

    /* ── 지표 B: 타 부문 거래 비중 (자기 자신 제외) ── */
    $('drill-m-off').textContent = pct(hasOff ? sh.off : 0);
    $('drill-m-off-band').replaceChildren(bandNode(bandB));

    $('drill-m-off-note').textContent = hasOff
      ? '그룹 안 ' + k + '개 중 ' + topOff.name + '(' + topOff.code + ') 한 곳이 ' +
        pct(topOff.offShare) + '를 차지합니다. 균등 기준선 ' + pct(1 / k) +
        ', 집중도 ' + conc.toFixed(2) + ' (1=완전 균등, ' + k + '=한 곳 집중).'
      : '이 그룹의 부문들은 서로 사고파는 관계가 계수에 잡히지 않습니다. ' +
        '내부 비중 ' + pct(sh.internal) + '는 사실상 전부 자기 자신(' + pct(sh.diag) + ')입니다.';
    /* ── 자동 판정 문장 ──
       숫자 3개(내부·자기 자신·타 부문)와 구간 이름을 문장에 직접 넣어,
       해석을 모르는 사용자도 읽기만 하면 이해되게 합니다. */
    var vn = [];
    vn.push(t('내부 비중 '));
    vn.push(b(pct(sh.internal)));
    vn.push(t('는 그룹 281개 중 '));
    vn.push(b(bandA.label));
    vn.push(t(' 구간(' + bandA.note + ')입니다. 그중 '));
    vn.push(b('자기 자신이 ' + pct(sh.diag)));
    vn.push(t('이고, 실제 '));
    vn.push(b('타 부문 거래는 ' + pct(sh.off)));
    vn.push(t('입니다. '));

    if (hasOff) {
      vn.push(t('타 부문 거래 중 가장 큰 '));
      vn.push(b(topOff.name));
      vn.push(t(josa(topOff.name, '이 ', '가 ')));
      vn.push(el('span', 'v-em', pct(topOff.offShare)));
      vn.push(t('를 차지해 '));
      vn.push(b(concentrationLabel(conc)));
      vn.push(t(' 상태입니다(균등 기준선 ' + pct(1 / k) + '). '));
    } else {
      vn.push(t('타 부문 거래가 0이라 '));
      vn.push(b('부문끼리 교류는 없는 셈'));
      vn.push(t('입니다 — 이 경우 내부 비중은 자기 자신의 크기일 뿐, ' +
        '그룹이 서로 얽혀 있다는 뜻은 아닙니다. '));
    }

    vn.push(t('그룹 밖으로 나간 파급은 '));
    vn.push(b(pct(sh.external)));
    vn.push(t('입니다.'));
    $('drill-verdict').replaceChildren.apply($('drill-verdict'), vn);

    /* ── 종목 탐색 관점의 해석 ──
       이 앱의 목적이 '산업 → 연관산업 → 종목'이므로, 그룹 내부 결합도가
       종목 선택에 무엇을 뜻하는지로 번역합니다. */
    var pn = [];
    pn.push(t('이 페이지의 목적은 산업에서 연관산업으로, 다시 종목으로 후보를 좁히는 것입니다. ' +
      '그룹 내부 연쇄는 "이 그룹에 속한 종목들을 한 묶음으로 봐도 되는지"를 알려줍니다. '));

    if (!hasOff) {
      pn.push(t('이 그룹은 부문끼리 거래가 0이므로, '));
      pn.push(b('같은 분류에 묶여 있다는 이유만으로 종목들이 함께 움직인다고 보기는 어렵습니다'));
      pn.push(t('. 실적이 같이 갈 이유가 약하니, 그룹 안을 훑기보다 아래 ①후방 1위 같은 ' +
        '그룹 밖 연관산업의 종목을 먼저 확인하는 편이 낫습니다.'));
    } else if (conc >= 2.5) {
      pn.push(t('사슬이 한 줄로 이어져 있습니다. 수요 충격이 '));
      pn.push(b(topOff.name));
      pn.push(t('을 거쳐 그룹 전체로 전달되므로, 종목 하나를 고르면 그 부문만이 아니라 '));
      pn.push(b('이 사슬 전체의 업황에 노출'));
      pn.push(t('됩니다. ' + topOff.name + '과 직접 거래하는 부문을 우선 보세요.'));
    } else {
      pn.push(t('부문끼리 비교적 고르게 얽혀 있습니다. 그룹에 속한 종목들은 업황을 공유하므로 '));
      pn.push(b('여러 종목을 담아도 분산 효과가 작습니다'));
      pn.push(t('. 대신 대표 종목 하나로 그룹 전체를 대신해 볼 수 있습니다.'));
    }

    if (topExt) {
      pn.push(t(' 참고로 이 그룹이 커질 때 가장 크게 함께 움직이는 그룹 밖 산업은 '));
      pn.push(b(topExt.name));
      pn.push(t('입니다(위 ①후방 1위). 아래 표의 부문 이름을 누르면 그 부문으로 바로 이동합니다.'));
    }
    $('drill-purpose').replaceChildren.apply($('drill-purpose'), pn);

    /* ── 구성원별 표 ── */
    var rows = back.memberDetail;
    var baseline = 1 / k;

    $('drill-internal-caption').textContent =
      job.parentName + ' · 그룹 내부 ' + k + '개 (타 부문 거래 큰 순)';

    body.replaceChildren.apply(body, rows.map(function (r) {
      var tr = el('tr');
      tr.appendChild(el('td', null, r.code + ' · ' + r.name));
      tr.appendChild(el('td', 'num', fmtCoef(r.value)));
      tr.appendChild(el('td', 'num', pct(r.internalShare)));
      tr.appendChild(el('td', 'uniform-cell', pct(r.uniformBaseline)));

      var offTd = el('td', 'num off-cell', r.off > 0 ? pct(r.offShare) : '0.0%');
      if (r.off <= 0) offTd.classList.add('is-zero');
      else if (r.offShare >= baseline) offTd.classList.add('is-strong');
      tr.appendChild(offTd);
      return tr;
    }));

    $('drill-internal-foot').textContent =
      '균등 기준선은 1/n = ' + pct(baseline) + '입니다. 내부 비중이 이 값에 가까우면 ' +
      '고르게 퍼진 것이고, 크게 벗어나면 한 부문에 쏠린 것입니다. 다만 내부 비중은 ' +
      '자기 자신(1.0)이 대부분이라 대개 기준선 근처에 붙습니다 — 실제 편차는 ' +
      '마지막 열 "타 부문 거래"로 판단하세요.';

    // 내부 표의 부문도 클릭하면 그 하위 분류로 이동합니다.
    Array.prototype.forEach.call(body.querySelectorAll('tr'), function (tr, i) {
      tr.style.cursor = 'pointer';
      tr.title = '이 부문을 분석 대상으로 선택';
      tr.addEventListener('click', function () {
        jumpToUnit(rows[i].code, job.childLevel);
      });
    });
  }

  /** 특정 분류 단위의 특정 부문으로 이동합니다(필요하면 파일을 새로 불러옵니다). */
  function jumpToUnit(code, level, scroll) {
    var ds = state.units[level];
    if (!ds) {
      setStatus($('load-status'), (IO_TREE.LABELS[level] || level) + ' 자료를 사용할 수 없습니다.', 'is-error');
      return;
    }

    /* 같은 단위여도 아직 그 파일을 불러오지 않았으면 state.data 가 null 입니다.
       (STEP 1 검색에서 결과를 누르는 경로가 여기에 해당합니다.) */
    if (state.level === level && state.data) {
      var idx = state.data.sectors.findIndex(function (s) { return s.code === code; });
      if (idx >= 0) {
        $('sector-search').value = '';
        buildSectorSelect();
        $('sector-select').value = String(idx);
        selectSector(idx);
        if (scroll) $('analysis').scrollIntoView({ block: 'start' });
        return;
      }
    }

    // 단위를 바꾼(또는 처음 불러오는) 뒤 그 코드를 선택하도록 예약합니다.
    pendingSelection = { level: level, code: code };
    state.keepScrollTo = !!scroll;
    loadUnit(level);
  }

  var pendingSelection = null;

  /* ── 7. 대표기업 ─────────────────────────────────── */
  function renderCompanyPlaceholder() {
    var empty = $('company-empty');
    empty.hidden = false;
    empty.textContent = 'STEP 1 자료를 불러오고 산업을 선택하면 대표기업이 표시됩니다.';
    $('company-table').hidden = true;
  }

  function renderCompanies(sector) {
    var label = CSV_LEVEL[state.level];
    var rows = state.companies.filter(function (c) {
      return c.sector_code === sector.code && (!c.level || c.level === label);
    });

    var basisNote = '';

    // 하위 분류(소분류·기본부문)는 대표기업을 직접 기입하지 않았으므로,
    // 상위 대분류의 대표기업을 "계통 기준"으로 보여줍니다.
    if (!rows.length && state.tree) {
      var path = state.tree.lineage(state.level, sector.code);
      var largeNode = path.filter(function (n) { return n.level === 'large'; })[0];
      if (largeNode && largeNode.code !== sector.code) {
        var inherited = state.companies.filter(function (c) {
          return c.sector_code === largeNode.code && (!c.level || c.level === '대분류');
        });
        if (inherited.length) {
          rows = inherited;
          basisNote = '상위 대분류 ' + largeNode.name + '의 대표기업입니다 (이 하위 부문에 직접 대응하는 종목은 따로 기입하지 않았습니다).';
        }
      }
    }

    var table = $('company-table');
    var empty = $('company-empty');
    var tbody = $('company-tbody');

    var headText = sector.code + ' · ' + sector.name + ' 부문의 대표기업';
    if (label) headText = label + ' ' + headText;
    $('company-sector').textContent = headText + (basisNote ? ' — ' + basisNote : '');

    if (!rows.length) {
      table.hidden = true;
      empty.hidden = false;
      empty.textContent = '이 부문은 대표기업 데이터가 아직 등록되지 않았습니다. data/companies.csv에 추가할 수 있습니다.';
      return;
    }

    var noteOnly = rows.length === 1 && !rows[0].ticker;

    if (noteOnly) {
      table.hidden = true;
      empty.hidden = false;
      empty.textContent = rows[0].note || '이 부문은 상장 대표기업이 확인되지 않았습니다.';
      return;
    }

    empty.hidden = true;
    table.hidden = false;
    $('company-caption').textContent =
      sector.name + ' · 대표기업 ' + rows.filter(function (r) { return r.ticker; }).length + '개' +
      (basisNote ? ' (상위 대분류 기준)' : '');

    tbody.replaceChildren.apply(tbody, rows.filter(function (r) { return r.ticker; }).map(function (r) {
      var tr = el('tr');
      tr.appendChild(el('td', null, r.company_name));
      var tcode = el('td');
      tcode.style.fontFamily = 'var(--font-mono)';
      tcode.textContent = r.ticker;
      tr.appendChild(tcode);
      tr.appendChild(el('td', null, r.market));
      tr.appendChild(el('td', null, r.note || ''));
      return tr;
    }));
  }

  /* ── 8. 이벤트 연결 ───────────────────────────────── */
  $('load-btn').addEventListener('click', function () {
    loadUnit(state.level);
  });

  Array.prototype.forEach.call($('unit-switch').querySelectorAll('button'), function (b) {
    b.addEventListener('click', function () {
      if (b.disabled || b.dataset.unit === state.level) return;
      loadUnit(b.dataset.unit);
    });
  });

  var sheetSearchTimer = null;
  $('sheet-search').addEventListener('input', function () {
    clearTimeout(sheetSearchTimer);
    sheetSearchTimer = setTimeout(renderSheetTable, 160);
  });

  /* STEP 1 이름 검색 — 입력이 멈춘 뒤 그립니다. */
  var finderTimer = null;
  if ($('finder-search')) {
    $('finder-search').addEventListener('input', function () {
      clearTimeout(finderTimer);
      finderTimer = setTimeout(renderFinderResults, 140);
    });
  }

  $('sector-search').addEventListener('input', filterSectors);

  $('sector-select').addEventListener('change', function () {
    var idx = Number(this.value);
    if (Number.isFinite(idx)) selectSector(idx);
  });

  if ($('goto-raw')) $('goto-raw').addEventListener('click', gotoRaw);
  if ($('drill-run')) $('drill-run').addEventListener('click', runDrill);
  if ($('drill-close')) $('drill-close').addEventListener('click', function () {
    $('drill-body').hidden = true;
    $('drill').hidden = true;
  });

  /* 물음표 버튼 — 모바일에는 hover 가 없으므로 '열기' 동작으로 연결합니다.
     판정문은 항상 보이게 두고, 이 버튼은 긴 상세 설명을 펼치는 역할만 합니다. */
  if ($('drill-help-open')) $('drill-help-open').addEventListener('click', function () {
    var d = $('drill-help');
    if (!d) return;
    d.open = !d.open;
    if (d.open) d.scrollIntoView({ block: 'nearest' });
  });

  init();
})();
