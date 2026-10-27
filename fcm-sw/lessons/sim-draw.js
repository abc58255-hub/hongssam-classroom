// 닮음 보물찾기 — 삼각형 SVG 공용 모듈 (교사 빌더·학생앱 공용)
// 세 변(a=BC, b=CA, c=AB)으로 좌표 계산 → show 플래그에 따라 변 길이·각도 라벨 표시 → 회전/반전.
(function (global) {
  function valid(a, b, c) { return a > 0 && b > 0 && c > 0 && a + b > c && b + c > a && a + c > b; }

  // 꼭짓점 좌표: B=(0,0), C=(a,0), A=위쪽
  function vertices(a, b, c) {
    var ax = (c * c - b * b + a * a) / (2 * a);
    var ay2 = c * c - ax * ax;
    var ay = ay2 > 0 ? Math.sqrt(ay2) : 0;
    return [{ x: ax, y: ay }, { x: 0, y: 0 }, { x: a, y: 0 }]; // [A,B,C]
  }
  // 각도(도): A는 b,c 사이 / B는 c,a 사이 / C는 a,b 사이
  function angles(a, b, c) {
    function ang(op, s1, s2) { var v = (s1 * s1 + s2 * s2 - op * op) / (2 * s1 * s2); return Math.acos(Math.max(-1, Math.min(1, v))) * 180 / Math.PI; }
    return [ang(a, b, c), ang(b, c, a), ang(c, a, b)]; // [A,B,C]
  }
  function num(n) { var r = Math.round(n * 10) / 10; return (r % 1 === 0) ? String(r) : r.toFixed(1); }

  // spec: {a,b,c, show:{sa,sb,sc,Aa,Ab,Ac}, rotate(도), flip, size, labelLetters}
  function draw(spec) {
    var a = +spec.a, b = +spec.b, c = +spec.c;
    var show = spec.show || {}, rot = (+spec.rotate || 0) * Math.PI / 180, flip = !!spec.flip;
    var size = spec.size || 220, pad = 34;
    if (!valid(a, b, c)) return '<div style="color:#ef4444;font-size:12px;padding:8px;">변 길이가 삼각형을 만들 수 없어요 (삼각부등식)</div>';
    var P = vertices(a, b, c), ang = angles(a, b, c);
    // 중심 기준 변환: flip(x) → 회전
    var cx = (P[0].x + P[1].x + P[2].x) / 3, cy = (P[0].y + P[1].y + P[2].y) / 3;
    var T = P.map(function (p) {
      var x = (flip ? -1 : 1) * (p.x - cx), y = (p.y - cy);
      return { x: x * Math.cos(rot) - y * Math.sin(rot), y: x * Math.sin(rot) + y * Math.cos(rot) };
    });
    // SVG는 y가 아래로 증가 → 수학 좌표 뒤집기
    T = T.map(function (p) { return { x: p.x, y: -p.y }; });
    // 바운딩 박스 → size에 맞게 스케일/이동
    var xs = T.map(function (p) { return p.x; }), ys = T.map(function (p) { return p.y; });
    var minX = Math.min.apply(null, xs), maxX = Math.max.apply(null, xs);
    var minY = Math.min.apply(null, ys), maxY = Math.max.apply(null, ys);
    var w = maxX - minX || 1, h = maxY - minY || 1;
    var s = Math.min((size - 2 * pad) / w, (size - 2 * pad) / h);
    var ox = (size - w * s) / 2 - minX * s, oy = (size - h * s) / 2 - minY * s;
    var Q = T.map(function (p) { return { x: p.x * s + ox, y: p.y * s + oy }; });
    var G = { x: (Q[0].x + Q[1].x + Q[2].x) / 3, y: (Q[0].y + Q[1].y + Q[2].y) / 3 };

    var poly = Q.map(function (p) { return num(p.x) + ',' + num(p.y); }).join(' ');
    var svg = '<svg viewBox="0 0 ' + size + ' ' + size + '" width="100%" style="max-width:' + size + 'px;display:block;margin:0 auto;">';
    svg += '<polygon points="' + poly + '" fill="rgba(37,99,235,.08)" stroke="#2563eb" stroke-width="2.5" stroke-linejoin="round"/>';

    // 변 라벨 (변 a=BC=인덱스 1-2, b=CA=2-0, c=AB=0-1)
    var sides = [{ k: 'sa', v: a, i: 1, j: 2 }, { k: 'sb', v: b, i: 2, j: 0 }, { k: 'sc', v: c, i: 0, j: 1 }];
    sides.forEach(function (sd) {
      if (!show[sd.k]) return;
      var mx = (Q[sd.i].x + Q[sd.j].x) / 2, my = (Q[sd.i].y + Q[sd.j].y) / 2;
      var dx = mx - G.x, dy = my - G.y, d = Math.hypot(dx, dy) || 1;
      var lx = mx + dx / d * 16, ly = my + dy / d * 16;
      svg += '<text x="' + num(lx) + '" y="' + num(ly) + '" font-size="15" font-weight="800" fill="#1e3a8a" text-anchor="middle" dominant-baseline="middle">' + num(sd.v) + '</text>';
    });
    // 각 라벨 (A,B,C = 인덱스 0,1,2)
    var verts = [{ k: 'Aa', v: ang[0], i: 0 }, { k: 'Ab', v: ang[1], i: 1 }, { k: 'Ac', v: ang[2], i: 2 }];
    verts.forEach(function (vt) {
      if (!show[vt.k]) return;
      var p = Q[vt.i], dx = G.x - p.x, dy = G.y - p.y, d = Math.hypot(dx, dy) || 1;
      var lx = p.x + dx / d * 22, ly = p.y + dy / d * 22;
      svg += '<text x="' + num(lx) + '" y="' + num(ly) + '" font-size="13" font-weight="800" fill="#b45309" text-anchor="middle" dominant-baseline="middle">' + num(vt.v) + '°</text>';
    });
    // 꼭짓점 글자(선택)
    if (spec.labelLetters) {
      ['A', 'B', 'C'].forEach(function (L, i) {
        var p = Q[i], dx = p.x - G.x, dy = p.y - G.y, d = Math.hypot(dx, dy) || 1;
        svg += '<text x="' + num(p.x + dx / d * 14) + '" y="' + num(p.y + dy / d * 14) + '" font-size="12" fill="#64748b" text-anchor="middle" dominant-baseline="middle">' + L + '</text>';
      });
    }
    svg += '</svg>';
    return svg;
  }

  // 닮음 조건 자동 추정: show 패턴 → 'SSS'|'SAS'|'AA'|'' (판정 불가)
  function guessCond(show) {
    show = show || {};
    var sN = (show.sa ? 1 : 0) + (show.sb ? 1 : 0) + (show.sc ? 1 : 0);
    var aN = (show.Aa ? 1 : 0) + (show.Ab ? 1 : 0) + (show.Ac ? 1 : 0);
    if (sN === 3) return 'SSS';
    if (aN >= 2) return 'AA';
    if (sN === 2 && aN >= 1) return 'SAS'; // 두 변 + (낀)각
    return '';
  }

  // 입력 방식 → 세 변 {a,b,c}
  //  sss: v1=a, v2=b, v3=c
  //  sas: v1=변b, v2=낀각A(°), v3=변c  → a=√(b²+c²-2bc·cosA)
  //  asa: v1=변a, v2=각B(°), v3=각C(°) → A=180-B-C, b=a·sinB/sinA, c=a·sinC/sinA
  function round1(n) { return Math.round(n * 100) / 100; }
  function computeABC(mode, v1, v2, v3) {
    v1 = +v1 || 0; v2 = +v2 || 0; v3 = +v3 || 0;
    if (mode === 'sas') { var A = v2 * Math.PI / 180; var a = Math.sqrt(Math.max(0, v1 * v1 + v3 * v3 - 2 * v1 * v3 * Math.cos(A))); return { a: round1(a), b: v1, c: v3 }; }
    if (mode === 'asa') { var B = v2 * Math.PI / 180, C = v3 * Math.PI / 180, Aa = Math.PI - B - C; if (Aa <= 0.0001) return { a: v1, b: 0, c: 0 }; return { a: v1, b: round1(v1 * Math.sin(B) / Math.sin(Aa)), c: round1(v1 * Math.sin(C) / Math.sin(Aa)) }; }
    return { a: v1, b: v2, c: v3 };
  }

  // 닮음비(정수비, 약분): 작은변 기준 배율 k를 간단한 분수로. 반환 {p,q} = t1 : t2
  function gcd(x, y) { x = Math.abs(Math.round(x)); y = Math.abs(Math.round(y)); while (y) { var t = y; y = x % y; x = t; } return x || 1; }
  function ratio(t1, t2) {
    var s1 = [t1.a, t1.b, t1.c].sort(function (m, n) { return m - n; });
    var s2 = [t2.a, t2.b, t2.c].sort(function (m, n) { return m - n; });
    if (!s2[0]) return { p: 1, q: 1, consistent: false };
    var k = s1[0] / s2[0];
    var consistent = Math.abs(s1[1] / s2[1] - k) < 0.03 && Math.abs(s1[2] / s2[2] - k) < 0.03;
    var best = { p: Math.max(1, Math.round(k)), q: 1, err: Math.abs(k - Math.round(k)) };
    for (var q = 1; q <= 12; q++) { var p = Math.round(k * q); if (p < 1) continue; var err = Math.abs(p / q - k); if (err < best.err) best = { p: p, q: q, err: err }; }
    var g = gcd(best.p, best.q);
    return { p: best.p / g, q: best.q / g, consistent: consistent };
  }
  function similar(t1, t2) { return ratio(t1, t2).consistent; }

  global.SimDraw = { draw: draw, valid: valid, angles: angles, guessCond: guessCond, ratio: ratio, similar: similar, computeABC: computeABC };
})(window);
