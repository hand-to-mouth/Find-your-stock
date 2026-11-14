/* io.js — 한국은행 투입산출표 「생산유발계수」 시트 파서
 *
 * 파일 구조(대·중·소·기본부문 4종 공통, 실제 파일로 검증함)
 *   R1  투입산출표_생산유발계수표(2023년 가격)
 *   R2  2023년
 *   R3  생산자가격
 *   R4  (빈 행)
 *   R5  [A1] "상품/상품"  [C~] 열 코드,  끝에 9090(행합) · 9802(감응도계수)
 *   R6  [C~] 열 부문명
 *   R7~ [A] 행 코드  [B] 행 부문명  [C~] 생산유발계수 n×n
 *   끝-1행 9590 열합 / 끝행 9801 영향력계수
 *
 * 계수 정의
 *   B(i,j) = j산업에 최종수요 1단위가 생길 때 i산업 생산 증가분
 *   영향력계수(j) = 열합(j) ÷ (전체합 ÷ n)   → 후방연쇄
 *   감응도계수(i) = 행합(i) ÷ (전체합 ÷ n)   → 전방연쇄
 */
(function (global) {
  'use strict';

  var SPECIAL_ROW_CODES = { '9590': true, '9801': true };

  function str(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\s+/g, ' ').trim();
  }

  function num(v) {
    var n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }

  /**
   * 생산유발계수 시트의 2차원 배열을 받아 분석용 객체로 변환합니다.
   * @param {Array<Array>} grid  XLSX.utils.sheet_to_json(ws, {header:1}) 결과
   * @param {string} label       분류 단위 이름 (예: '대분류')
   */
  function parse(grid, label) {
    if (!Array.isArray(grid) || grid.length < 10) {
      throw new Error('시트 내용이 예상보다 짧습니다. 올바른 투입산출표 파일인지 확인해 주세요.');
    }

    // 1) 영향력계수(9801) 행 찾기 — 이 행을 기준으로 데이터 구간을 확정합니다.
    var impactRow = -1;
    for (var i = 0; i < grid.length; i++) {
      if (str((grid[i] || [])[0]) === '9801') { impactRow = i; break; }
    }
    if (impactRow < 0) {
      throw new Error('「생산유발계수」 시트에서 영향력계수(9801) 행을 찾지 못했습니다.');
    }

    // 2) 부문 행(코드와 이름이 모두 있는 행) 수집
    var rowIdx = [];
    for (var r = 0; r < impactRow; r++) {
      var row = grid[r] || [];
      var code = str(row[0]);
      var name = str(row[1]);
      if (!code || !name) continue;
      if (SPECIAL_ROW_CODES[code]) continue;
      rowIdx.push(r);
    }
    var n = rowIdx.length;
    if (n < 3) {
      throw new Error('부문 행을 충분히 찾지 못했습니다 (n=' + n + '). 파일 구조를 확인해 주세요.');
    }

    var sectors = rowIdx.map(function (ri) {
      return { code: str(grid[ri][0]), name: str(grid[ri][1]) };
    });

    // 3) 열 헤더 (코드 행 = 첫 부문행-2, 이름 행 = 첫 부문행-1)
    var firstRow = rowIdx[0];
    var codeRow = grid[firstRow - 2] || [];
    var nameRow = grid[firstRow - 1] || [];

    var cols = [];
    for (var j = 0; j < n; j++) {
      cols.push({ code: str(codeRow[2 + j]), name: str(nameRow[2 + j]) });
    }

    // 행 순서와 열 순서가 같은지 확인 (4개 파일 모두 동일했지만 방어적으로 검사)
    var mismatched = 0;
    for (var k = 0; k < n; k++) {
      if (cols[k].code && cols[k].code !== sectors[k].code) mismatched++;
    }

    // 4) 특수 열(행합 9090 / 감응도계수 9802) 위치
    var sumCol = -1, respCol = -1;
    for (var c = 0; c < codeRow.length; c++) {
      var cv = str(codeRow[c]);
      if (cv === '9090') sumCol = c;
      if (cv === '9802') respCol = c;
    }

    // 5) 계수 행렬과 총합
    var matrix = [];
    var total = 0;
    for (var mi = 0; mi < n; mi++) {
      var src = grid[rowIdx[mi]] || [];
      var arr = new Array(n);
      for (var mj = 0; mj < n; mj++) {
        var v = num(src[2 + mj]);
        arr[mj] = v;
        total += v;
      }
      matrix.push(arr);
    }

    // 6) 계수는 표에 적힌 발표값을 그대로 읽습니다(직접 읽기).
    //    동시에 행렬로 재계산해 두 값이 일치하는지 교차 검증합니다.
    var impactRowVals = grid[impactRow] || [];
    var influence = [];     // 영향력계수 (열 기준 · 후방연쇄)
    var sensitivity = [];   // 감응도계수 (행 기준 · 전방연쇄)
    var rowSum = new Array(n).fill(0);   // 행렬 행합
    var colSum = new Array(n).fill(0);   // 행렬 열합
    var reportedRowSum = [];             // 시트의 9090(행합) 열 값

    for (var a = 0; a < n; a++) {
      influence.push(num(impactRowVals[2 + a]));
      var srow = grid[rowIdx[a]] || [];
      reportedRowSum.push(sumCol >= 0 ? num(srow[sumCol]) : 0);
      sensitivity.push(respCol >= 0 ? num(srow[respCol]) : 0);
      for (var b = 0; b < n; b++) {
        rowSum[a] += matrix[a][b];
        colSum[b] += matrix[a][b];
      }
    }

    // 7) 교차 검증
    var avg = n > 0 ? total / n : 0;
    var maxResidual = 0;     // 계수 재계산값과 발표값의 최대 차이
    var maxRowResidual = 0;  // 행합 재계산값과 표기값의 최대 차이

    if (avg > 0) {
      for (var vi = 0; vi < n; vi++) {
        if (influence[vi] !== 0) {
          maxResidual = Math.max(maxResidual, Math.abs(colSum[vi] / avg - influence[vi]));
        }
        if (sensitivity[vi] !== 0) {
          maxResidual = Math.max(maxResidual, Math.abs(rowSum[vi] / avg - sensitivity[vi]));
        }
        if (reportedRowSum[vi] !== 0) {
          maxRowResidual = Math.max(maxRowResidual, Math.abs(rowSum[vi] - reportedRowSum[vi]));
        }
      }
    }

    var sumInfluence = influence.reduce(function (s, x) { return s + x; }, 0);
    var sumSensitivity = sensitivity.reduce(function (s, x) { return s + x; }, 0);

    // 부문 코드 → 열 인덱스
    var colIndexByCode = {};
    for (var ci = 0; ci < n; ci++) {
      if (cols[ci].code) colIndexByCode[cols[ci].code] = ci;
    }

    return {
      label: label || '',
      n: n,
      sectors: sectors,
      cols: cols,
      matrix: matrix,
      influence: influence,
      sensitivity: sensitivity,
      rowSum: rowSum,
      colSum: colSum,
      colIndexByCode: colIndexByCode,
      columnOrderMatches: mismatched === 0,
      validation: {
        total: total,
        average: avg,
        maxResidual: maxResidual,
        sumInfluence: sumInfluence,
        sumSensitivity: sumSensitivity,
        expectedSum: n
      }
    };
  }

  /** 선택 산업의 열(column) = 후방연쇄: 그 산업에 최종수요가 생길 때 생산이 늘어나는 산업 */
  function topBackward(data, index, limit) {
    var target = data.sectors[index];
    var colIdx = data.colIndexByCode[target.code];
    if (colIdx === undefined) colIdx = index;
    var out = [];
    for (var i = 0; i < data.n; i++) {
      if (data.sectors[i].code === target.code) continue;
      out.push({ code: data.sectors[i].code, name: data.sectors[i].name, value: data.matrix[i][colIdx] });
    }
    out.sort(function (a, b) { return b.value - a.value; });
    return out.slice(0, limit || 5);
  }

  /** 선택 산업의 행(row) = 전방연쇄: 그 산업의 생산을 가장 크게 끌어올리는 최종수요 산업 */
  function topForward(data, index, limit) {
    var target = data.sectors[index];
    var out = [];
    for (var j = 0; j < data.n; j++) {
      if (data.cols[j].code === target.code) continue;
      out.push({ code: data.cols[j].code, name: data.cols[j].name, value: data.matrix[index][j] });
    }
    out.sort(function (a, b) { return b.value - a.value; });
    return out.slice(0, limit || 5);
  }

  /**
   * 두 트랙의 상위 산업과 파급 비중을 함께 계산합니다.
   *
   * 비중의 정의
   *   후방(back): 선택 산업 A에 최종수요 1단위가 생길 때 A를 제외한 나머지 산업으로
   *               흘러간 파급의 합 T = Σ b(i, A) (i ≠ A) 를 100%로 봅니다.
   *   전방(fwd) : 대칭으로 T = Σ b(A, j) (j ≠ A).
   *   자기 자신(대각값)은 정의상 대부분 1.0을 넘어 절반 이상을 차지하므로 분모에서 제외합니다.
   *   그래야 "다른 산업과의 연관"이라는 의미가 유지됩니다.
   *
   * @param {object} data   IO.parse 결과
   * @param {number} index  선택 산업 인덱스
   * @param {string} dir    'back' | 'fwd'
   * @param {number} limit  상위 몇 개 (기본 5)
   */
  function breakdown(data, index, dir, limit) {
    var target = data.sectors[index];
    var items = [];
    var total = 0;
    var selfValue = 0;

    if (dir === 'back') {
      var colIdx = data.colIndexByCode[target.code];
      if (colIdx === undefined) colIdx = index;
      selfValue = data.matrix[index][colIdx];
      for (var i = 0; i < data.n; i++) {
        if (data.sectors[i].code === target.code) continue;
        var v = data.matrix[i][colIdx];
        total += v;
        items.push({ code: data.sectors[i].code, name: data.sectors[i].name, value: v });
      }
    } else {
      selfValue = data.matrix[index][index];
      for (var j = 0; j < data.n; j++) {
        if (data.cols[j].code === target.code) continue;
        var w = data.matrix[index][j];
        total += w;
        items.push({ code: data.cols[j].code, name: data.cols[j].name, value: w });
      }
    }

    items.sort(function (a, b) { return b.value - a.value; });

    var cum = 0;
    var top = items.slice(0, limit || 5).map(function (item) {
      cum += item.value;
      return {
        code: item.code,
        name: item.name,
        value: item.value,
        share: total > 0 ? item.value / total : 0,
        cumulative: total > 0 ? cum / total : 0
      };
    });

    return {
      items: top,
      total: total,          // 자기 제외 파급 합
      selfValue: selfValue,  // 자기 자신(대각값)
      topSum: cum,
      topShare: total > 0 ? cum / total : 0,
      otherCount: Math.max(0, items.length - top.length)
    };
  }

  /**
   * 그룹 평균 집계 — 상위 분류 하나를, 그 자식 부문들의 **평균**으로 봅니다.
   *
   * 자식 부문 집합 M (|M| = k) 에 대하여, 부문 크기가 제각각이므로 합계가 아니라 1/k 평균을 씁니다.
   * 그래야 자식이 많은 그룹(예: 화학제품 8개 중분류)이 자식이 적은 그룹보다
   * 자동으로 커지지 않습니다.
   *
   *   후방(back): v_i = (1/k) · Σ_{j∈M} b(i, j)   M 전체에 최종수요 1단위가 생길 때 i의 생산 증가
   *   전방(fwd) : u_j = (1/k) · Σ_{i∈M} b(i, j)   M 전체의 산출이 필요해질 때 j의 생산 증가
   *
   * i ∈ M 이면 '그룹 내부'(같은 상위 분류 안의 거래), 아니면 '그룹 외부'입니다.
   * 비중·누적비중은 1단계와 같은 원칙으로 **그룹 외부 파급의 합을 100%**로 계산합니다.
   * (그룹 내부는 정의상 값이 크게 나오므로 분모에서 빼야 "다른 산업과의 연관"이 남습니다.)
   *
   * ── 내부를 다시 셋으로 나눕니다 (3분해) ──
   * 그룹 내부의 값은 거의 전부 '자기 자신' b(m,m)≈1.0 입니다. 실측하면 내부 합의
   * 89~99%가 대각값이라, 내부 비중만 보면 "그룹 안에서 서로 사고판다"와
   * "각자 자기 생산만 한다"가 구별되지 않습니다. 그래서 구성원마다
   *
   *   내부 기여 = (1/k) Σ_{q∈M} b(m, q)     ← 자기 자신 포함
   *   자기 자신 = b(m, m) / k
   *   타 부문 거래 = 내부 기여 − 자기 자신
   *
   * 로 나눠서 돌려줍니다. 세 값의 합은 내부 기여 총합이 되고, 여기에 그룹 외부
   * 총합을 더하면 그룹 전체 파급(= expectedTotal)이 되어 교차 검증이 됩니다.
   *
   * @param {object} data         집계에 쓸 계수 행렬 (자식 부문이 들어 있는 분류 단위)
   * @param {Array<string>} memberCodes  그룹에 속한 자식 부문 코드
   * @param {string} dir          'back' | 'fwd'
   * @param {number} limit        상위 몇 개 (기본 5)
   */
  function aggregateGroup(data, memberCodes, dir, limit) {
    if (!memberCodes || !memberCodes.length) {
      throw new Error('그룹 구성 부문이 비어 있습니다.');
    }

    var memberSet = {};
    memberCodes.forEach(function (c) { memberSet[c] = true; });

    /* 구성원의 열 인덱스와 행 인덱스를 함께 확보합니다.
       자기 자신(대각)과 타 부문 거래를 가르려면 행·열 두 좌표가 모두 필요합니다. */
    var refs = [];
    memberCodes.forEach(function (code) {
      var ci = data.colIndexByCode[code];
      if (ci === undefined) return;
      var ri = -1;
      for (var r = 0; r < data.n; r++) {
        if (data.sectors[r].code === code) { ri = r; break; }
      }
      if (ri < 0) ri = ci;   // 행·열 순서가 같으므로 안전한 대체값
      refs.push({
        code: code,
        name: data.cols[ci].name || data.sectors[ri].name,
        col: ci,
        row: ri
      });
    });

    if (!refs.length) {
      throw new Error('그룹 구성 부문을 이 분류 단위에서 찾지 못했습니다.');
    }
    var k = refs.length;

    var items = [];
    var i, j, sum;

    if (dir === 'back') {
      for (i = 0; i < data.n; i++) {
        sum = 0;
        for (j = 0; j < k; j++) sum += data.matrix[i][refs[j].col];
        var codeB = data.sectors[i].code;
        items.push({
          code: codeB,
          name: data.sectors[i].name,
          value: sum / k,
          inside: !!memberSet[codeB]
        });
      }
    } else {
      for (j = 0; j < data.n; j++) {
        sum = 0;
        for (i = 0; i < k; i++) sum += data.matrix[refs[i].row][j];
        var codeF = data.cols[j].code;
        items.push({
          code: codeF,
          name: data.cols[j].name,
          value: sum / k,
          inside: !!memberSet[codeF]
        });
      }
    }

    items.sort(function (a, b) { return b.value - a.value; });

    var external = items.filter(function (x) { return !x.inside; });
    var internal = items.filter(function (x) { return x.inside; });

    var externalTotal = external.reduce(function (s, x) { return s + x.value; }, 0);
    var internalTotal = internal.reduce(function (s, x) { return s + x.value; }, 0);

    var cum = 0;
    var top = external.slice(0, limit || 5).map(function (x) {
      cum += x.value;
      return {
        code: x.code,
        name: x.name,
        value: x.value,
        share: externalTotal > 0 ? x.value / externalTotal : 0,
        cumulative: externalTotal > 0 ? cum / externalTotal : 0
      };
    });

    var internalList = internal.map(function (x) {
      return {
        code: x.code,
        name: x.name,
        value: x.value,
        share: internalTotal > 0 ? x.value / internalTotal : 0
      };
    });

    // ── 구성원별 3분해 ──────────────────────────────
    var memberDetail = refs.map(function (m) {
      var sumIn = 0;
      for (var q = 0; q < k; q++) {
        var other = refs[q];
        sumIn += (dir === 'back')
          ? data.matrix[m.row][other.col]
          : data.matrix[other.row][m.col];
      }
      var diag = data.matrix[m.row][m.col];
      return {
        code: m.code,
        name: m.name,
        value: sumIn / k,          // 내부 기여(자기 자신 포함) = 그룹 평균 계수
        diag: diag / k,            // 자기 자신(대각)
        off: (sumIn - diag) / k    // 타 부문 거래(비대각)
      };
    });

    var diagTotal = memberDetail.reduce(function (s, m) { return s + m.diag; }, 0);
    var offTotal = memberDetail.reduce(function (s, m) { return s + m.off; }, 0);

    var groupTotal = externalTotal + internalTotal;
    var shares = {
      internal: groupTotal > 0 ? internalTotal / groupTotal : 0,
      diag: groupTotal > 0 ? diagTotal / groupTotal : 0,
      off: groupTotal > 0 ? offTotal / groupTotal : 0,
      external: groupTotal > 0 ? externalTotal / groupTotal : 0
    };

    /* 타 부문 거래가 그룹 안에서 고르게 퍼졌는지 — 허핀달 지수.
       k배 해 두면 1 = 완전 균등, k = 한 부문에 전부 집중이라 읽기 쉽습니다. */
    var hhi = 0;
    if (offTotal > 0) {
      memberDetail.forEach(function (m) {
        var q = m.off / offTotal;
        hhi += q * q;
      });
    }
    var maxOff = memberDetail.reduce(function (acc, m) {
      return m.off > acc ? m.off : acc;
    }, 0);

    memberDetail.sort(function (a, b) { return b.off - a.off; });
    memberDetail.forEach(function (m) {
      m.internalShare = internalTotal > 0 ? m.value / internalTotal : 0;
      m.offShare = offTotal > 0 ? m.off / offTotal : 0;
      m.uniformBaseline = 1 / k;
    });

    // 교차 검증용 기대값: (1/k)·Σ 열합(후방) 또는 (1/k)·Σ 행합(전방) = 그룹 전체 파급
    var expectedTotal = 0;
    var cj, rj;
    if (dir === 'back') {
      for (cj = 0; cj < k; cj++) expectedTotal += data.colSum[refs[cj].col];
      expectedTotal = expectedTotal / k;
    } else {
      for (rj = 0; rj < k; rj++) expectedTotal += data.rowSum[refs[rj].row];
      expectedTotal = expectedTotal / k;
    }

    return {
      dir: dir,
      size: k,
      items: top,                       // 그룹 외부 상위
      externalTotal: externalTotal,
      internalItems: internalList,      // 그룹 내부 전부
      internalTotal: internalTotal,
      diagTotal: diagTotal,             // 그중 자기 자신
      offTotal: offTotal,               // 그중 타 부문 거래
      shares: shares,                   // 내부/자기/타부문/외부 비중 (분모: 그룹 전체 파급)
      memberDetail: memberDetail,       // 구성원별 3분해 (타 부문 거래 큰 순)
      offConcentration: offTotal > 0 ? k * hhi : NaN,  // 1=균등, k=집중
      maxOffShare: offTotal > 0 ? maxOff / offTotal : NaN,
      offTopMember: memberDetail.length ? memberDetail[0] : null,
      topSum: cum,
      topShare: externalTotal > 0 ? cum / externalTotal : 0,
      otherCount: Math.max(0, external.length - top.length),
      groupTotal: groupTotal,
      expectedTotal: expectedTotal,
      residual: Math.abs(groupTotal - expectedTotal)
    };
  }

  /** 값이 큰 순서의 순위(1위 = 1) */
  function rankOf(values, index) {
    var target = values[index];
    var rank = 1;
    for (var i = 0; i < values.length; i++) {
      if (i === index) continue;
      if (values[i] > target) rank++;
    }
    return rank;
  }

  global.IO = {
    parse: parse,
    topBackward: topBackward,
    topForward: topForward,
    breakdown: breakdown,
    aggregateGroup: aggregateGroup,
    rankOf: rankOf,
    str: str
  };
})(window);
