/* graph-plot.js — 그래프 탐구 공용 좌표평면 렌더러
 * 교사(graph-teacher.html)·학생(graph.html) 공유.
 * 진짜 좌표평면: 격자·축(화살표)·눈금 라벨·원점 O
 * 상호작용: 드래그 이동(마우스/터치), 확대·축소(휠/핀치/버튼), 자동 맞춤
 * 개형: kind별 곡선(현재 line) + 공개 애니메이션
 * 점: 정답(초록)/오답(회색) + 이름 표시 옵션
 *
 * 사용:
 *   var plot = GraphPlot.attach(canvasEl, { interactive:true });
 *   plot.setRoom({a,b,c,kind,revealed,show_wrong,show_names});
 *   plot.setPoints([{x,y,correct,name}]);
 *   plot.render();  plot.fit();  plot.resetView();  plot.animateReveal();
 */
(function (global) {
  'use strict';

  var COL = {
    bg: '#fffdf7', grid: '#efe6d4', gridMinor: '#f6efe0',
    axis: '#8a6034', tick: '#a68a63', label: '#7a5230',
    curve: '#dc2626', ok: '#16a34a', wrong: '#c4b5a0', name: '#7a5230'
  };

  function niceStep(minUnits) {
    var pow = Math.pow(10, Math.floor(Math.log10(minUnits)));
    var cands = [1, 2, 5, 10];
    for (var i = 0; i < cands.length; i++) if (cands[i] * pow >= minUnits) return cands[i] * pow;
    return 10 * pow;
  }
  function fmt(n) {
    if (Math.abs(n) < 1e-9) return '0';
    var r = Math.round(n * 1000) / 1000;
    return String(r);
  }

  function attach(canvas, opts) {
    opts = opts || {};
    var ctx = canvas.getContext('2d');
    var room = null, points = [], reveal = 1; // reveal: 0~1 애니메이션 진행도
    var view = { ox: 0, oy: 0, scale: 24 };    // ox,oy=원점 화면좌표(px), scale=단위당 px
    var W = 0, H = 0, dpr = 1, inited = false;

    function sizeCanvas() {
      dpr = global.devicePixelRatio || 1;
      W = canvas.clientWidth; H = canvas.clientHeight;
      canvas.width = Math.round(W * dpr); canvas.height = Math.round(H * dpr);
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    }
    function centerView(halfRange) {
      sizeCanvas();
      var r = halfRange || 10;
      view.scale = Math.min(W, H) / (r * 2 + 1);
      view.ox = W / 2; view.oy = H / 2;
      inited = true;
    }
    function px(x) { return view.ox + x * view.scale; }
    function py(y) { return view.oy - y * view.scale; }
    function wx(sx) { return (sx - view.ox) / view.scale; }
    function wy(sy) { return (view.oy - sy) / view.scale; }

    // ── 자동 맞춤: 점 + (공개된) 직선 절편이 다 보이게 ──
    function fit() {
      sizeCanvas();
      var xs = [], ys = [];
      points.forEach(function (p) { if (room && !room.show_wrong && !p.correct) return; xs.push(p.x); ys.push(p.y); });
      if (!xs.length) { centerView(10); render(); return; }
      var minX = Math.min.apply(null, xs), maxX = Math.max.apply(null, xs);
      var minY = Math.min.apply(null, ys), maxY = Math.max.apply(null, ys);
      minX = Math.min(minX, 0); maxX = Math.max(maxX, 0);
      minY = Math.min(minY, 0); maxY = Math.max(maxY, 0);
      var spanX = Math.max(maxX - minX, 4), spanY = Math.max(maxY - minY, 4);
      var pad = 1.15;
      var sx = W / (spanX * pad), sy = H / (spanY * pad);
      view.scale = Math.max(6, Math.min(sx, sy));
      view.ox = W / 2 - ((minX + maxX) / 2) * view.scale;
      view.oy = H / 2 + ((minY + maxY) / 2) * view.scale;
      inited = true; render();
    }
    function resetView() { centerView(10); render(); }
    function setHalfRange(r) { centerView(r); render(); }

    function zoomAt(sx, sy, factor) {
      var bx = wx(sx), by = wy(sy);
      view.scale = Math.max(3, Math.min(400, view.scale * factor));
      view.ox = sx - bx * view.scale;
      view.oy = sy + by * view.scale;
      render();
    }
    function zoomBy(factor) { zoomAt(W / 2, H / 2, factor); }
    function panBy(dx, dy) { view.ox += dx; view.oy += dy; render(); }

    // ── 렌더 ──
    function render() {
      if (!inited) centerView(10);
      if (W !== canvas.clientWidth || H !== canvas.clientHeight) sizeCanvas();
      ctx.clearRect(0, 0, W, H);
      ctx.fillStyle = COL.bg; ctx.fillRect(0, 0, W, H);

      var step = niceStep(64 / view.scale);         // 큰 격자 간격(단위)
      var minx = wx(0), maxx = wx(W), miny = wy(H), maxy = wy(0);

      // 격자
      ctx.lineWidth = 1;
      ctx.strokeStyle = COL.grid; ctx.beginPath();
      var i0 = Math.ceil(minx / step), i1 = Math.floor(maxx / step);
      for (var i = i0; i <= i1; i++) { var X = px(i * step); ctx.moveTo(X, 0); ctx.lineTo(X, H); }
      var j0 = Math.ceil(miny / step), j1 = Math.floor(maxy / step);
      for (var j = j0; j <= j1; j++) { var Y = py(j * step); ctx.moveTo(0, Y); ctx.lineTo(W, Y); }
      ctx.stroke();

      // 축 (화면 안에 원점이 없으면 가장자리에 클램프)
      var ax = Math.max(0, Math.min(W, view.ox));
      var ay = Math.max(0, Math.min(H, view.oy));
      ctx.strokeStyle = COL.axis; ctx.lineWidth = 2;
      ctx.beginPath(); ctx.moveTo(0, ay); ctx.lineTo(W, ay); ctx.stroke();   // x축
      ctx.beginPath(); ctx.moveTo(ax, 0); ctx.lineTo(ax, H); ctx.stroke();   // y축
      // 화살표
      arrow(W - 2, ay, 1, 0); arrow(ax, 2, 0, -1);
      ctx.fillStyle = COL.axis; ctx.font = 'italic 13px serif';
      ctx.fillText('x', W - 14, ay - 8); ctx.fillText('y', ax + 8, 14);

      // 눈금 숫자
      ctx.fillStyle = COL.tick; ctx.font = '11px sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      for (var k = i0; k <= i1; k++) { if (k === 0) continue; var lx = px(k * step); ctx.fillText(fmt(k * step), lx, clamp(ay + 4, 4, H - 14)); }
      ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      for (var m = j0; m <= j1; m++) { if (m === 0) continue; var ly = py(m * step); ctx.fillText(fmt(m * step), clamp(ax - 6, 22, W - 4), ly); }
      // 원점 O
      if (view.ox >= -20 && view.ox <= W + 20 && view.oy >= -20 && view.oy <= H + 20) {
        ctx.textAlign = 'right'; ctx.textBaseline = 'top'; ctx.fillText('O', ax - 5, ay + 4);
      }

      // 개형(곡선)
      if (room && room.revealed && reveal > 0) drawCurve(minx, maxx);

      // 점
      points.forEach(function (p) {
        if (room && !room.show_wrong && !p.correct) return;
        var X = px(p.x), Y = py(p.y);
        if (X < -12 || X > W + 12 || Y < -12 || Y > H + 12) return;
        ctx.beginPath(); ctx.arc(X, Y, 6, 0, Math.PI * 2);
        ctx.fillStyle = p.correct ? COL.ok : COL.wrong; ctx.fill();
        ctx.strokeStyle = '#fff'; ctx.lineWidth = 2; ctx.stroke();
        if (room && room.show_names && p.name) {
          ctx.fillStyle = COL.name; ctx.font = '700 11px sans-serif';
          ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
          ctx.fillText(p.name, X, Y - 8);
        }
      });
    }

    function arrow(x, y, dx, dy) {
      var s = 7; ctx.fillStyle = COL.axis; ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x - dx * s - dy * s * 0.6, y - dy * s + dx * s * 0.6);
      ctx.lineTo(x - dx * s + dy * s * 0.6, y - dy * s - dx * s * 0.6);
      ctx.closePath(); ctx.fill();
    }
    function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

    // kind별 곡선 — 현재 line만, 확장 시 여기 case 추가
    function drawCurve(minx, maxx) {
      var kind = (room.kind || 'line');
      ctx.strokeStyle = COL.curve; ctx.lineWidth = 3; ctx.lineJoin = 'round';
      var sx, sy, ex, ey; // 시작/끝 (화면)
      if (kind === 'line') {
        var a = Number(room.a), b = Number(room.b), c = Number(room.c);
        if (b === 0) { if (a === 0) return; var xv = c / a; sx = px(xv); sy = 0; ex = px(xv); ey = H; }
        else { sx = px(minx); sy = py((c - a * minx) / b); ex = px(maxx); ey = py((c - a * maxx) / b); }
      } else return;
      // reveal 애니메이션: 시작→끝 부분만
      var tx = sx + (ex - sx) * reveal, ty = sy + (ey - sy) * reveal;
      ctx.beginPath(); ctx.moveTo(sx, sy); ctx.lineTo(tx, ty); ctx.stroke();
    }

    function animateReveal() {
      reveal = 0; var t0 = null;
      function frame(ts) {
        if (t0 === null) t0 = ts;
        var p = Math.min(1, (ts - t0) / 550);
        reveal = p < 1 ? (1 - Math.pow(1 - p, 3)) : 1; // easeOutCubic
        render();
        if (p < 1) global.requestAnimationFrame(frame);
      }
      global.requestAnimationFrame(frame);
    }

    // ── 상호작용 (드래그·휠·핀치) ──
    if (opts.interactive !== false) {
      var ptrs = {}, lastPan = null, pinch = null;
      canvas.style.touchAction = 'none'; canvas.style.cursor = 'grab';
      canvas.addEventListener('pointerdown', function (e) {
        canvas.setPointerCapture(e.pointerId);
        ptrs[e.pointerId] = { x: e.clientX, y: e.clientY };
        var ids = Object.keys(ptrs);
        if (ids.length === 1) { lastPan = { x: e.clientX, y: e.clientY }; canvas.style.cursor = 'grabbing'; }
        else if (ids.length === 2) { pinch = pinchState(ids); lastPan = null; }
      });
      canvas.addEventListener('pointermove', function (e) {
        if (!ptrs[e.pointerId]) return;
        ptrs[e.pointerId] = { x: e.clientX, y: e.clientY };
        var ids = Object.keys(ptrs);
        if (ids.length >= 2 && pinch) {
          var np = pinchState(ids);
          var f = np.dist / (pinch.dist || 1);
          var rect = canvas.getBoundingClientRect();
          zoomAt(np.cx - rect.left, np.cy - rect.top, f);
          pinch = np;
        } else if (ids.length === 1 && lastPan) {
          panBy(e.clientX - lastPan.x, e.clientY - lastPan.y);
          lastPan = { x: e.clientX, y: e.clientY };
        }
      });
      function up(e) {
        delete ptrs[e.pointerId];
        var ids = Object.keys(ptrs);
        if (ids.length < 2) pinch = null;
        if (ids.length === 1) lastPan = { x: ptrs[ids[0]].x, y: ptrs[ids[0]].y };
        if (ids.length === 0) { lastPan = null; canvas.style.cursor = 'grab'; }
      }
      canvas.addEventListener('pointerup', up);
      canvas.addEventListener('pointercancel', up);
      canvas.addEventListener('wheel', function (e) {
        e.preventDefault();
        var rect = canvas.getBoundingClientRect();
        zoomAt(e.clientX - rect.left, e.clientY - rect.top, e.deltaY < 0 ? 1.12 : 1 / 1.12);
      }, { passive: false });
      function pinchState(ids) {
        var a = ptrs[ids[0]], b = ptrs[ids[1]];
        return { dist: Math.hypot(a.x - b.x, a.y - b.y), cx: (a.x + b.x) / 2, cy: (a.y + b.y) / 2 };
      }
    }

    global.addEventListener('resize', function () { if (inited) { sizeCanvas(); render(); } });

    return {
      setRoom: function (r) { room = r; return this; },
      setPoints: function (arr) { points = arr || []; return this; },
      addPoint: function (p) { points.push(p); return this; },
      render: render, fit: fit, resetView: resetView, setHalfRange: setHalfRange,
      zoomBy: zoomBy, animateReveal: animateReveal,
      get room() { return room; }
    };
  }

  global.GraphPlot = { attach: attach };
})(window);
