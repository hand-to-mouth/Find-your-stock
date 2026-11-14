/* tree.js — 「2020 기준년 상품분류(투입산출표)」 계층 파서
 *
 * sector-tree.xlsx 「부문분류」 시트 구조 (실제 파일로 검증함)
 *   R1  제목            2020 기준년 상품분류(투입산출표)
 *   R2  단위 표기        기본부문(380) | 소분류(165) | 중분류(83) | 대분류(33)
 *   R3  열 머리글        코드 | 상품명  (×4쌍)
 *   R4~ 데이터
 *       A/B 기본부문   C/D 소분류   E/F 중분류   G/H 대분류
 *       상위 분류는 병합 셀이라 첫 행에만 적혀 있으므로 아래로 채워(fill-down) 읽습니다.
 *
 * 코드 체계
 *   기본부문 4자리 → 소분류 3자리 → 중분류 2자리 → 대분류 1~3자리
 *   (예: 0111 벼 → 011 곡물 및 식량작물 → 01 작물 → A 농림수산품)
 *   9로 시작하는 코드(9090·9801·9811 …)는 계(計) 행이므로 제외합니다.
 *
 * 검증 결과: 대분류 33 / 중분류 83 / 소분류 165 / 기본부문 380 이
 *          4개 투입산출표 파일의 부문 코드와 완전히 일치했습니다(누락 0).
 */
(function (global) {
  'use strict';

  var LEVELS = ['large', 'medium', 'small', 'basic'];   // 큰 분류 → 작은 분류
  var LABELS = { large: '대분류', medium: '중분류', small: '소분류', basic: '기본부문' };
  var WIDTH = { basic: 4, small: 3, medium: 2, large: 0 };  // 코드 자리수(대분류는 문자 포함)

  // 열 배치: [코드열, 이름열]
  var COLUMNS = { basic: [0, 1], small: [2, 3], medium: [4, 5], large: [6, 7] };

  function str(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\s+/g, ' ').trim();
  }

  /** 숫자로 읽혀 앞자리 0이 사라진 코드를 원래 자리수로 되돌립니다. */
  function padCode(code, width) {
    if (!width) return code;
    if (!/^\d+$/.test(code)) return code;
    while (code.length < width) code = '0' + code;
    return code;
  }

  /** 계(計) 행 여부 — 9090·9111·9590·9801·9811 등 */
  function isAggregate(code) {
    return code.charAt(0) === '9';
  }

  /**
   * @param {Array<Array>} grid  XLSX.utils.sheet_to_json(ws, {header:1}) 결과
   */
  function parse(grid) {
    if (!Array.isArray(grid) || grid.length < 10) {
      throw new Error('분류표 시트가 예상보다 짧습니다.');
    }

    var nodes = { large: [], medium: [], small: [], basic: [] };
    var seen = { large: {}, medium: {}, small: {}, basic: {} };
    var cursor = {
      large: { code: '', name: '' },
      medium: { code: '', name: '' },
      small: { code: '', name: '' }
    };

    for (var r = 3; r < grid.length; r++) {
      var row = grid[r] || [];
      var basicCode = padCode(str(row[COLUMNS.basic[0]]), WIDTH.basic);
      if (!basicCode || basicCode === '코드') continue;
      if (isAggregate(basicCode)) continue;

      var raw = {};
      LEVELS.forEach(function (level) {
        var c = COLUMNS[level];
        raw[level] = {
          code: padCode(str(row[c[0]]), WIDTH[level]),
          name: str(row[c[1]])
        };
      });

      // 병합 셀 복원: 상위 분류는 첫 행에만 값이 있습니다.
      ['small', 'medium', 'large'].forEach(function (level) {
        if (raw[level].code) {
          cursor[level] = raw[level];
        } else {
          raw[level] = { code: cursor[level].code, name: cursor[level].name };
        }
      });

      if (!raw.small.code || !raw.medium.code || !raw.large.code) {
        throw new Error('분류표 ' + (r + 1) + '행에서 상위 분류를 찾지 못했습니다.');
      }

      // 부모 → 자식 순서로 등록
      addNode('large', raw.large, null);
      addNode('medium', raw.medium, raw.large.code);
      addNode('small', raw.small, raw.medium.code);
      addNode('basic', { code: basicCode, name: raw.basic.name || str(row[COLUMNS.basic[1]]) }, raw.small.code);
    }

    function addNode(level, item, parentCode) {
      if (seen[level][item.code]) return;
      seen[level][item.code] = true;
      nodes[level].push({
        level: level,
        label: LABELS[level],
        code: item.code,
        name: item.name,
        parentCode: parentCode
      });
    }

    // 코드 → 노드
    var index = { large: {}, medium: {}, small: {}, basic: {} };
    LEVELS.forEach(function (level) {
      nodes[level].forEach(function (node) { index[level][node.code] = node; });
    });

    // 부모 코드 → 자식 목록
    var childMap = { large: {}, medium: {}, small: {} };
    ['medium', 'small', 'basic'].forEach(function (level) {
      nodes[level].forEach(function (node) {
        if (!node.parentCode) return;
        var bucket = childMap[parentOfLevel(level)];
        if (!bucket) return;
        if (!bucket[node.parentCode]) bucket[node.parentCode] = [];
        bucket[node.parentCode].push(node);
      });
    });

    function parentOfLevel(level) {
      var i = LEVELS.indexOf(level);
      return i > 0 ? LEVELS[i - 1] : null;
    }

    function childrenOf(level, code) {
      var bucket = childMap[level];
      if (!bucket || !bucket[code]) return [];
      return bucket[code].slice();
    }

    function parentOf(level, code) {
      var node = index[level] && index[level][code];
      if (!node || !node.parentCode) return null;
      var pLevel = parentOfLevel(level);
      return index[pLevel] ? index[pLevel][node.parentCode] || null : null;
    }

    /** 위에서 아래로 이어지는 분류 경로(대분류 → … → 현재) */
    function lineage(level, code) {
      var out = [];
      var cur = index[level] && index[level][code];
      while (cur) {
        out.unshift(cur);
        cur = parentOf(cur.level, cur.code);
      }
      return out;
    }

    /** 대분류 코드 하나에 속한 중분류 목록 (드릴다운용) */
    function childrenOfLarge(largeCode) {
      return childrenOf('large', largeCode);
    }

    return {
      levels: LEVELS.slice(),
      labels: LABELS,
      nodes: nodes,
      index: index,
      childrenOf: childrenOf,
      childrenOfLarge: childrenOfLarge,
      parentOf: parentOf,
      parentOfLevel: parentOfLevel,
      lineage: lineage,
      counts: {
        large: nodes.large.length,
        medium: nodes.medium.length,
        small: nodes.small.length,
        basic: nodes.basic.length
      }
    };
  }

  global.IO_TREE = {
    parse: parse,
    LEVELS: LEVELS,
    LABELS: LABELS
  };
})(window);
