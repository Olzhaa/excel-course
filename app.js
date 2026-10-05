/* Excel & Google Sheets course platform — single-page app (no build step). */
(function () {
  'use strict';

  var CONFIG = window.COURSE_CONFIG || {};
  var S = { course: null, user: null, creds: null, progress: {}, settings: { openLessons: 0, quizAttempts: 1 }, quizDraft: {}, quizResult: {}, quizReview: {}, practiceResult: {}, teacher: null, menuOpen: false };
  var app = document.getElementById('app');
  var CRED_KEY = 'xlcourse-creds';

  // ---------- helpers ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }
  function md(text) {
    if (!window.marked) return '<p>' + esc(text) + '</p>';
    var html = window.marked.parse(String(text || ''), { gfm: true, breaks: false });
    return html.replace(/<table>/g, '<div class="table-wrap"><table>').replace(/<\/table>/g, '</table></div>');
  }
  function store(k, v) { try { if (v === undefined) return JSON.parse(localStorage.getItem(k) || 'null'); if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, JSON.stringify(v)); } catch (e) { return null; } }
  function pct(a, b) { return b ? Math.round(100 * a / b) : 0; }
  function toast(msg) {
    var t = document.createElement('div'); t.className = 'toast'; t.setAttribute('role', 'status'); t.textContent = msg; document.body.appendChild(t);
    setTimeout(function () { t.remove(); }, 2600);
  }
  function go(hash) { if (location.hash === hash) render(); else location.hash = hash; }
  function colLetter(n) { var s = ''; n = n + 1; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  // ---------- API ----------
  function api(action, payload) {
    var body = Object.assign({ action: action, login: S.creds && S.creds.login, password: S.creds && S.creds.password }, payload || {});
    if (CONFIG.demo) {
      return new Promise(function (res) { setTimeout(function () { res(window.DemoBackend.call(body)); }, 150); });
    }
    if (!CONFIG.apiUrl) return Promise.resolve({ ok: false, error: 'The platform is not connected to its Google Sheet yet (apiUrl is empty in config.js).' });
    return fetch(CONFIG.apiUrl, { method: 'POST', body: JSON.stringify(body) })
      .then(function (r) { return r.json(); })
      .catch(function () { return { ok: false, network: true, error: 'No connection to the server. Check your internet and try again.' }; });
  }

  // ---------- course helpers ----------
  function lessons() { return S.course.lessons; }
  function lessonById(id) { return lessons().filter(function (l) { return l.id === id; })[0]; }
  function isTeacher() { return S.user && S.user.role === 'teacher'; }
  function lessonComplete(lesson, progress) {
    progress = progress || S.progress;
    for (var i = 0; i < lesson.practice.length; i++) { var r = progress[lesson.practice[i].id]; if (!r || !r.max || r.score < r.max) return false; }
    var q = progress[lesson.id + '-Q'];
    if (!q || !q.attempts || !q.max) return false;
    var pass = S.settings.passPercent == null || S.settings.passPercent === '' ? 50 : Number(S.settings.passPercent);
    return 100 * q.score / q.max >= pass || q.attempts >= Number(S.settings.quizAttempts);
  }
  function unlockedUpTo(progress) {
    progress = progress || S.progress;
    var cap = Number(S.settings.openLessons), manual = progress.UNLOCK ? Number(progress.UNLOCK.score) || 0 : 0, n = 1;
    while (n < cap && (lessonComplete(lessons()[n - 1], progress) || n + 1 <= manual)) n++;
    return Math.min(n, cap);
  }
  function isOpen(lesson) { return isTeacher() || lesson.num <= unlockedUpTo(); }
  function lockReason(lesson) {
    if (lesson.num > Number(S.settings.openLessons)) return 'Your teacher has not opened this lesson yet.';
    var prev = lessons()[lesson.num - 2];
    return 'Finish lesson ' + prev.num + ' first: solve all its practice tasks and score at least ' + (S.settings.passPercent == null ? 50 : S.settings.passPercent) + '% in its quiz (or use all quiz attempts).';
  }
  function projectOpen(p) { return isTeacher() || (Number(S.settings.openLessons) >= p.openAfter && lessonComplete(lessons()[p.openAfter - 1])); }
  function lessonStats(lesson, progress) {
    progress = progress || S.progress;
    var q = progress[lesson.id + '-Q'];
    var solved = lesson.practice.filter(function (p) { var r = progress[p.id]; return r && r.score >= r.max && r.max > 0; }).length;
    var quizDone = !!(q && q.attempts > 0);
    var status = quizDone && solved === lesson.practice.length ? 'done' : (quizDone || solved > 0 || (q && q.attempts)) ? 'part' : 'none';
    return { quiz: q, quizPct: q && q.max ? pct(q.score, q.max) : null, solved: solved, total: lesson.practice.length, status: status };
  }
  function overall(progress) {
    progress = progress || S.progress;
    var items = 0, done = 0, qsum = 0, qn = 0, psolved = 0, ptotal = 0;
    lessons().forEach(function (l) {
      var st = lessonStats(l, progress);
      items += 1 + st.total; done += (st.quiz && st.quiz.attempts ? 1 : 0) + st.solved;
      if (st.quizPct != null) { qsum += st.quizPct; qn++; }
      psolved += st.solved; ptotal += st.total;
    });
    var projects = S.course.projects.map(function (p) { return progress[p.id] || null; });
    return { pct: pct(done, items), quizAvg: qn ? Math.round(qsum / qn) : null, quizCount: qn, psolved: psolved, ptotal: ptotal, projects: projects };
  }
  function fileLink(f) {
    if (!f) return '';
    return '<a class="file-link" href="' + esc(f.url) + '" download="' + esc(f.filename) + '"><span class="ext">XLSX</span>' + esc(f.filename) + '</a>';
  }

  // ---------- shell ----------
  function shell(location_, title, inner, opts) {
    opts = opts || {};
    var demo = CONFIG.demo ? '<div class="demo-banner">Preview mode: progress is saved only in this browser. Try <b>student / student123</b> or <b>teacher / teacher123</b>.</div>' : '';
    return demo +
      '<header class="topbar">' +
      '<button class="btn ghost menu-btn" data-act="menu" aria-label="Open menu" aria-expanded="false" aria-controls="side-nav">☰</button>' +
      '<a class="brand" href="#/home" aria-label="Home"><span class="brand-mark">' + '<i></i>'.repeat(9) + '</span><span class="brand-text">' + esc(S.course.title) + '</span></a>' +
      '<div class="namebox" title="Where you are">' + esc(location_) + '</div>' +
      '<div class="fx"><em>fx</em><span>' + esc(title) + '</span></div>' +
      '<div class="top-actions"><span class="small muted user-name" style="white-space:nowrap">' + esc(S.user.name) + '</span>' +
      '<a class="btn small ghost acct" href="#/account">Account</a><button class="btn small" data-act="logout">Log out</button></div>' +
      '</header>' +
      '<div class="shell"><nav id="side-nav" aria-label="Course" class="side' + (S.menuOpen ? ' open' : '') + '">' + sideNav(opts.active) + '</nav>' +
      '<main class="main"><div class="wrap' + (opts.wide ? ' wide' : '') + '">' + inner + '</div></main></div>';
  }
  function sideNav(active) {
    var h = '';
    if (isTeacher()) {
      h += '<div class="side-section">Teacher</div>';
      h += navItem('#/teacher', 'T1', 'Gradebook', '', active === 'teacher');
      h += navItem('#/teacher/projects', 'T2', 'Grade projects', '', active === 'tprojects');
      h += navItem('#/teacher/accounts', 'T3', 'Student accounts', '', active === 'taccounts');
    }
    h += '<div class="side-section">Course</div>';
    h += navItem('#/home', 'A1', 'Overview', '', active === 'home');
    h += navItem('#/projects', 'P', 'Projects', '', active === 'projects');
    h += navItem('#/account', 'ME', 'My account', '', active === 'account');
    var week = 0;
    lessons().forEach(function (l) {
      if (l.week !== week) { week = l.week; h += '<div class="side-section">Week ' + week + '</div>'; }
      var st = lessonStats(l);
      var open = isOpen(l);
      h += '<a class="nav' + (active === l.id ? ' active' : '') + (open ? '' : ' locked') + '" href="#/lesson/' + l.id + '">' +
        '<span class="cellref">' + esc(l.id) + '</span><span>' + esc(l.title) + '</span>' +
        (open ? '<span class="dot ' + st.status + '" title="' + st.status + '"></span>' : '<span class="small" title="Not open yet">🔒</span>') + '</a>';
    });
    return h;
  }
  function navItem(href, ref, label, extra, on) {
    return '<a class="nav' + (on ? ' active' : '') + '" href="' + href + '"><span class="cellref">' + ref + '</span><span>' + esc(label) + '</span>' + (extra || '<span></span>') + '</a>';
  }

  // ---------- views ----------
  function viewLogin(err) {
    var google = !!CONFIG.googleClientId;
    app.innerHTML =
      '<div class="login-page"><form class="card login-card" id="login-form" autocomplete="on">' +
      '<div class="row"><span class="brand-mark">' + '<i></i>'.repeat(9) + '</span><span class="eyebrow">' + esc(S.course ? S.course.group : '') + '</span></div>' +
      '<h1>' + esc(S.course ? S.course.title : 'Course') + '</h1>' +
      '<p class="muted" style="margin:0">' + esc(S.course ? S.course.tagline : '') + '</p>' +
      (google ? '<div id="g-btn" style="min-height:44px"></div><div class="or-line"><span>or use your login and password</span></div>' : '') +
      '<label class="field" for="lg-login">' + (google ? 'Login or college email' : 'Login') + '<input id="lg-login" name="username" type="text" autocomplete="username" required autocapitalize="none" spellcheck="false"></label>' +
      '<label class="field" for="lg-pass">Password<input id="lg-pass" name="password" type="password" autocomplete="current-password" required></label>' +
      '<div class="error" id="lg-err" role="alert">' + esc(err || '') + '</div>' +
      '<button class="btn primary" type="submit" id="lg-btn">Log in</button>' +
      (CONFIG.demo ? '<p class="small muted" style="margin:0">Preview accounts: <span class="mono">student / student123</span> or <span class="mono">teacher / teacher123</span></p>' : '<p class="small muted" style="margin:0">' + (google ? 'Use your college Google account, or the login and password from your teacher.' : 'Your login and password are on the card from your teacher.') + '</p>') +
      '<div class="sheet-tabs"><span class="on">Login</span><span>Lessons</span><span>Practice</span><span>Quiz</span><span>Projects</span></div>' +
      '</form></div>';
    document.getElementById('login-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var btn = document.getElementById('lg-btn'); btn.disabled = true; btn.textContent = 'Checking…';
      S.creds = { login: document.getElementById('lg-login').value.trim().toLowerCase(), password: document.getElementById('lg-pass').value.trim() };
      api('login').then(function (r) {
        if (!r.ok) {
          S.creds = null;
          var er = document.getElementById('lg-err'); er.textContent = r.error;
          btn.disabled = false; btn.textContent = 'Log in';
          var pw = document.getElementById('lg-pass'); pw.value = ''; pw.focus();
          return;
        }
        S.creds.login = r.user.login;
        afterLogin(r);
      });
    });
    if (google) renderGoogleButton();
  }
  function afterLogin(r) {
    S.quizRun = null; S.quizDraft = {}; S.quizResult = {}; S.practiceResult = {}; S.quizReview = {};
    S.user = r.user; S.progress = r.progress || {}; S.settings = r.settings;
    store(CRED_KEY, S.creds);
    loadVariant().then(function () {
      if (!location.hash || location.hash === '#/login') location.hash = isTeacher() ? '#/teacher' : '#/home'; else render();
    });
  }
  // Every student has personal data in some tasks and projects: overlay data/v/vNN.json on the common course.
  function loadVariant() {
    if (!S.baseCourse) S.baseCourse = S.course;
    S.course = S.baseCourse;
    var v = S.user && !isTeacher() && Number(S.user.variant);
    if (!v) return Promise.resolve();
    var id = (v < 10 ? '0' : '') + v;
    var src = window.COURSE_VARIANTS ? Promise.resolve(window.COURSE_VARIANTS[String(v)] || null)
      : fetch('data/v/v' + id + '.json', { cache: 'no-cache' }).then(function (r) { return r.ok ? r.json() : null; });
    return src.then(function (ov) {
      if (!ov) return;
      var c = JSON.parse(JSON.stringify(S.baseCourse));
      c.lessons.forEach(function (l) {
        l.practice = l.practice.map(function (p) { var o = ov.practice && ov.practice[p.id]; return o ? Object.assign({}, p, o) : p; });
      });
      c.projects.forEach(function (p) {
        var o = ov.projects && ov.projects[p.id];
        if (o) { p.personal = o.personal; if (o.file) p.file = o.file; }
      });
      c.variant = v;
      S.course = c;
    }).catch(function () {});
  }
  function renderGoogleButton() {
    function draw() {
      var el = document.getElementById('g-btn');
      if (!el || !window.google || !google.accounts) return;
      google.accounts.id.initialize({
        client_id: CONFIG.googleClientId,
        callback: function (resp) {
          var e = document.getElementById('lg-err'); if (e) e.textContent = 'Signing in…';
          api('googleLogin', { idToken: resp.credential }).then(function (r) {
            if (!r.ok) { viewLogin(r.error); return; }
            S.creds = r.creds;
            afterLogin(r);
          });
        }
      });
      google.accounts.id.renderButton(el, { theme: 'outline', size: 'large', text: 'signin_with', shape: 'rectangular', width: Math.min(360, el.clientWidth || 320) });
    }
    if (window.google && window.google.accounts) { draw(); return; }
    if (!document.getElementById('gsi-script')) {
      var sc = document.createElement('script');
      sc.src = 'https://accounts.google.com/gsi/client'; sc.async = true; sc.id = 'gsi-script';
      sc.onload = draw;
      document.head.appendChild(sc);
    } else document.getElementById('gsi-script').addEventListener('load', draw);
  }

  function viewHome() {
    var o = overall();
    var next = lessons().filter(function (l) { return isOpen(l) && lessonStats(l).status !== 'done'; })[0];
    var subs = o.projects.filter(Boolean).length;
    var h = '<section class="hero"><div class="eyebrow">' + esc(S.course.group) + ' · ' + esc(S.course.lessons.length) + ' lessons · 10 weeks</div>' +
      '<h1>Hello, ' + esc(S.user.name.split(' ')[0]) + '</h1><p class="lead">' + esc(S.course.tagline) + '</p></section>';
    h += '<section class="card stats">' +
      stat('Course progress', o.pct + '%', '<div class="bar"><i style="width:' + o.pct + '%"></i></div>') +
      stat('Quiz average', o.quizAvg == null ? '—' : o.quizAvg + '%', '<span class="small muted">' + o.quizCount + ' quizzes taken</span>') +
      stat('Practice tasks solved', o.psolved + ' / ' + o.ptotal, '') +
      stat('Projects submitted', subs + ' / ' + S.course.projects.length, '') + '</section>';
    if (next) {
      h += '<section class="card"><div class="eyebrow">Continue</div><div class="row" style="justify-content:space-between;margin-top:6px">' +
        '<div style="min-width:0"><h2 style="font-size:22px">' + esc(next.id) + ' · ' + esc(next.title) + '</h2><p class="muted" style="margin:4px 0 0">' + esc(next.summary) + '</p></div>' +
        '<a class="btn primary" href="#/lesson/' + next.id + '">Open lesson</a></div></section>';
    }
    h += '<section class="card"><h2 style="font-size:22px;margin-bottom:12px">Course plan</h2><div class="week-grid">';
    var weeks = {};
    lessons().forEach(function (l) { (weeks[l.week] = weeks[l.week] || []).push(l); });
    Object.keys(weeks).forEach(function (w) {
      h += '<div class="week-row"><div class="wk">Week ' + w + '</div>';
      weeks[w].forEach(function (l) { h += lessonTile(l); });
      h += '</div>';
    });
    h += '</div></section>';
    h += projectsList();
    return shell('A1', 'Overview', h, { active: 'home' });
  }
  function stat(label, value, extra) { return '<div class="stat"><span class="small muted">' + esc(label) + '</span><b>' + value + '</b>' + extra + '</div>'; }
  function lessonTile(l) {
    var st = lessonStats(l), open = isOpen(l);
    var meta = !open ? '<span class="pill">🔒 Locked</span>' :
      '<span class="pill ' + (st.quizPct == null ? '' : st.quizPct >= 80 ? 'good' : st.quizPct >= 50 ? 'warn' : 'bad') + '">Quiz ' + (st.quizPct == null ? '—' : st.quizPct + '%') + '</span>' +
      '<span class="pill ' + (st.solved === st.total ? 'good' : st.solved ? 'warn' : '') + '">Practice ' + st.solved + '/' + st.total + '</span>';
    return '<a class="lesson-tile' + (open ? '' : ' locked') + '" href="#/lesson/' + l.id + '"><span class="cellref">' + l.id + '</span><span class="t">' + esc(l.title) + '</span><span class="meta">' + meta + '</span></a>';
  }
  function projectsList() {
    var h = '<section class="card"><h2 style="font-size:22px;margin-bottom:12px">Projects</h2><div class="grid2">';
    S.course.projects.forEach(function (p) {
      var r = S.progress[p.id];
      var open = projectOpen(p);
      var badge = !open ? '<span class="pill">🔒 Opens after you finish lesson ' + p.openAfter + '</span>' :
        r && r.detail && r.detail.grade !== undefined && r.detail.grade !== '' ? '<span class="pill good">Graded: ' + esc(r.detail.grade) + '/100</span>' :
        r ? '<span class="pill warn">Submitted, waiting for grade</span>' : '<span class="pill accent">Due: ' + esc(p.due) + '</span>';
      h += '<a class="lesson-tile' + (open ? '' : ' locked') + '" href="#/project/' + p.id + '"><span class="cellref">' + p.id + '</span><span class="t">' + esc(p.title) + '</span><span class="small muted">' + esc(p.summary) + '</span><span class="meta">' + badge + '</span></a>';
    });
    return h + '</div></section>';
  }

  function viewLesson(id, tab) {
    var l = lessonById(id);
    if (!l) return shell('#REF!', 'Lesson not found', '<div class="card">This lesson does not exist. <a href="#/home">Back to overview</a></div>');
    tab = tab || 'read';
    var st = lessonStats(l);
    var head = '<section class="lesson-head"><div class="eyebrow">Week ' + l.week + ' · Lesson ' + l.num + ' of ' + lessons().length + '</div><h1>' + esc(l.title) + '</h1><p class="lead">' + esc(l.summary) + '</p></section>';
    if (!isOpen(l)) {
      return shell(l.id, l.title, head + '<div class="card notice">🔒 ' + esc(lockReason(l)) + '</div>', { active: l.id });
    }
    var tabs = '<div class="tabs" role="tablist">' +
      tabBtn(l.id, 'read', 'Read', tab) +
      tabBtn(l.id, 'practice', 'Practice · ' + st.solved + '/' + st.total, tab) +
      tabBtn(l.id, 'quiz', 'Quiz' + (S.quizRun && S.quizRun[l.id] ? ' · in progress' : st.quiz && st.quiz.attempts && st.quizPct != null ? ' · ' + st.quizPct + '%' : ''), tab) + '</div>';
    var body = tab === 'practice' ? practiceView(l) : tab === 'quiz' ? quizView(l) : readView(l);
    var idx = lessons().indexOf(l);
    var nav = '<div class="row" style="justify-content:space-between">' +
      (idx > 0 ? '<a class="btn" href="#/lesson/' + lessons()[idx - 1].id + '">← ' + esc(lessons()[idx - 1].id) + '</a>' : '<span></span>') +
      (tab === 'read' ? '<a class="btn primary" href="#/lesson/' + l.id + '/practice">Go to practice →</a>' : tab === 'practice' ? '<a class="btn primary" href="#/lesson/' + l.id + '/quiz">Go to quiz →</a>' :
        (idx < lessons().length - 1 ? '<a class="btn" href="#/lesson/' + lessons()[idx + 1].id + '">' + esc(lessons()[idx + 1].id) + ' →</a>' : '<span></span>')) + '</div>';
    return shell(l.id + (tab === 'read' ? '' : tab === 'practice' ? ':P' : ':Q'), l.title, head + tabs + body + nav, { active: l.id });
  }
  function tabBtn(id, key, label, cur) { return '<button class="' + (key === cur ? 'on' : '') + '"' + (key === cur ? ' aria-current="page"' : '') + ' data-go="#/lesson/' + id + '/' + key + '">' + esc(label) + '</button>'; }

  function readView(l) {
    return '<section class="card"><div class="eyebrow">In this lesson you will</div><ul class="goals">' + l.goals.map(function (g) { return '<li>' + esc(g) + '</li>'; }).join('') + '</ul></section>' +
      '<article class="card prose">' + md(l.theory) + '</article>';
  }

  function practiceView(l) {
    var h = '<p class="muted" style="margin:0">Download the starter file, do the steps in Excel or Google Sheets (File → Import → Upload), then type your answers below and press <b>Check</b>. You can check as many times as you need.</p>';
    l.practice.forEach(function (p, i) {
      var rec = S.progress[p.id];
      var solved = rec && rec.max && rec.score >= rec.max;
      var res = S.practiceResult[p.id] || {};
      var correct = rec && rec.detail && rec.detail.correct || {};
      h += '<section class="card task" id="task-' + p.id + '"><div class="task-head"><h3><span class="cellref">' + colLetter(i) + (i + 1) + '</span> ' + esc(p.title) + '</h3>' +
        '<div class="row"><span class="pill ' + (p.difficulty === 'easy' ? 'good' : p.difficulty === 'medium' ? 'warn' : 'bad') + '">' + esc(p.difficulty) + '</span>' +
        (solved ? '<span class="pill good">✓ Solved</span>' : rec ? '<span class="pill warn">' + rec.score + '/' + rec.max + ' correct</span>' : '') + '</div></div>' +
        (p.file ? '<div>' + fileLink(p.file) + '</div>' : '') +
        '<div class="prose">' + md(p.steps) + '</div>' +
        '<form class="checks" data-task="' + p.id + '"><div class="eyebrow">Check your work</div>';
      p.checks.forEach(function (c) {
        var val = res.answers ? res.answers[c.id] : correct[c.id];
        var mark = res.results ? (res.results[c.id] ? '<span class="mark ok" role="img" aria-label="Correct" title="Correct">✓</span>' : '<span class="mark no" role="img" aria-label="Not correct yet" title="Not correct yet">✗</span>') : (correct[c.id] != null ? '<span class="mark ok">✓</span>' : '<span class="mark"></span>');
        var fid = 'ck-' + p.id + '-' + c.id;
        if (c.type === 'choice') {
          h += '<div class="check"><label class="q" for="' + fid + '">' + esc(c.prompt) + '</label><div class="check-row"><select id="' + fid + '" name="' + c.id + '"><option value="">Choose…</option>' +
            c.options.map(function (o, k) { return '<option value="' + k + '"' + (String(val) === String(k) ? ' selected' : '') + '>' + esc(o) + '</option>'; }).join('') + '</select>' + mark + '</div></div>';
        } else {
          h += '<div class="check"><label class="q" for="' + fid + '">' + esc(c.prompt) + '</label><div class="check-row"><input type="text" id="' + fid + '" name="' + c.id + '" value="' + esc(val == null ? '' : val) + '" autocomplete="off" spellcheck="false"' + (c.type === 'number' ? ' inputmode="decimal"' : '') + ' placeholder="' + (c.type === 'number' ? 'Type a number' : 'Type your answer') + '">' + mark + '</div></div>';
        }
      });
      h += '<div class="row"><button class="btn primary" type="submit">Check</button>' +
        (res.score != null ? '<span class="' + (res.score === res.max ? 'pill good' : 'pill warn') + '">' + (res.score === res.max ? 'All correct. Well done!' : res.score + ' of ' + res.max + ' correct. Fix the ✗ answers and check again.') + '</span>' : '') + '</div></form>' +
        (p.hint ? '<details class="hint"><summary>Need a hint?</summary><p>' + esc(p.hint) + '</p></details>' : '') + '</section>';
    });
    return h;
  }

  function seeded(str) {
    var h = 2166136261;
    for (var i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); }
    return function () { h += 0x6D2B79F5; var t = h; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
  }
  function shuffled(n, rnd) { var a = []; for (var i = 0; i < n; i++) a.push(i); for (var j = n - 1; j > 0; j--) { var k = Math.floor(rnd() * (j + 1)); var t = a[j]; a[j] = a[k]; a[k] = t; } return a; }
  function fmtSec(sec) { if (sec == null) return '—'; sec = Math.max(0, Math.round(sec)); return Math.floor(sec / 60) + ':' + ('0' + sec % 60).slice(-2); }

  function quizView(l) {
    var rec = S.progress[l.id + '-Q'];
    var limit = Number(S.settings.quizAttempts), minutes = Number(S.settings.quizMinutes || 15);
    var attempts = rec ? rec.attempts : 0;
    var result = S.quizResult[l.id];
    if (rec && rec.detail && rec.detail.startedAt && !(S.quizRun && S.quizRun[l.id]) && !isTeacher()) {
      var dl = Date.parse(rec.detail.startedAt) / 1000 + Number(S.settings.quizMinutes || 15) * 60;
      if (dl > Date.now() / 1000) { S.quizRun = S.quizRun || {}; S.quizRun[l.id] = { deadline: dl, order: quizOrder(l, attempts + 1), away: 0 }; }
    }
    var run = S.quizRun && S.quizRun[l.id];
    var h = '<section class="card result-banner">' +
      (rec && rec.attempts ? '<b class="num">' + pct(rec.score, rec.max) + '%</b><div><div>Best score: ' + rec.score + ' / ' + rec.max + '</div><div class="small muted">Attempts used: ' + attempts + ' of ' + limit + '</div></div>'
        : '<div><div><b style="font-size:20px">' + l.quiz.length + ' questions · ' + minutes + ' minutes</b></div><div class="small muted">You have ' + limit + ' attempts. Your best score counts.</div></div>') + '</section>';
    if (run) {
      return '<div class="quiz-timer" id="quiz-timer" role="timer">⏱ <span id="quiz-left">' + fmtSec(run.deadline - Date.now() / 1000) + '</span> left</div>' +
        '<form id="quiz-form" class="no-copy" data-lesson="' + l.id + '">' + quizQuestions(l, S.quizDraft[l.id] || {}, null, null, false, run.order) +
        '<div class="row" style="margin-top:16px"><button class="btn primary" type="submit">Submit answers</button><span class="small muted" id="quiz-msg" role="status"></span></div></form>';
    }
    if (result) {
      h += '<section class="card"><div class="result-banner"><b class="num">' + result.score + ' / ' + result.max + '</b><div>' +
        (result.late ? 'Time was over, so this attempt counts as 0.' : result.score === result.max ? 'Perfect score!' : result.attemptsLeft > 0 ? 'You have ' + result.attemptsLeft + ' attempt' + (result.attemptsLeft === 1 ? '' : 's') + ' left. Questions marked ✗ were wrong.' : 'No attempts left.') +
        '</div>' + (result.attemptsLeft > 0 && result.score < result.max ? '<button class="btn primary" data-act="quizstart" data-lesson="' + l.id + '">Try again</button>' : '') + '</div></section>';
      return h + quizQuestions(l, result.answers, result.results, result.reveal, true, result.order);
    }
    var finished = rec && (attempts >= limit || (rec.detail && rec.detail.last === rec.max && rec.attempts > 0));
    var expired = rec && rec.detail && rec.detail.startedAt && !run && !isTeacher();
    if (expired && attempts + 1 >= limit) {
      return h + '<section class="card" style="display:grid;gap:10px"><div>The time for this quiz ran out before your answers were sent, so this attempt counts as 0.</div>' +
        '<div><button class="btn primary" data-act="quizclose" data-lesson="' + l.id + '">OK, continue</button></div></section>';
    }
    if (finished && !isTeacher()) {
      if (Number(S.settings.showAnswers)) {
        var rv = S.quizReview[l.id];
        if (!rv) {
          S.quizReview[l.id] = 'loading';
          api('quizReview', { lessonId: l.id }).then(function (r) { S.quizReview[l.id] = r.ok && r.reveal ? r.reveal : 'none'; render(); });
        }
        if (rv && rv !== 'loading' && rv !== 'none') {
          var last = rec.detail && rec.detail.lastAnswers || {};
          var res = l.quiz.map(function (_, i) { return String(last[i]) === String(rv.answers[i]); });
          return h + '<div class="card muted">This quiz is finished. Here are your last answers and the correct ones.</div>' + quizQuestions(l, last, res, rv, true, quizOrder(l, attempts));
        }
      }
      return h + '<div class="card muted">This quiz is finished.' + (Number(S.settings.showAnswers) ? '' : ' Correct answers are not shown, so that classmates cannot copy them. Ask your teacher if you want to discuss a question.') + '</div>';
    }
    h += '<section class="card" style="display:grid;gap:10px"><h2 style="font-size:20px">Before you start</h2><ul style="margin:0;padding-left:20px">' +
      '<li>You have <b>' + minutes + ' minutes</b>. When the time is over, your answers are sent automatically.</li>' +
      '<li>Questions and answers are in a different order for every student.</li>' +
      '<li>Stay on this page. Leaving it (other tabs or apps) is recorded and your teacher can see it.</li>' +
      '<li>Starting the quiz uses one attempt, even if you close the page.</li></ul>' +
      '<div><button class="btn primary" data-act="quizstart" data-lesson="' + l.id + '">Start quiz' + (attempts ? ' (attempt ' + (attempts + 1) + ' of ' + limit + ')' : '') + '</button></div></section>';
    return h;
  }
  function quizQuestions(l, answers, results, reveal, locked, order) {
    answers = answers || {};
    order = order || { q: l.quiz.map(function (_, i) { return i; }), o: l.quiz.map(function (q) { return q.options.map(function (_, k) { return k; }); }) };
    var h = '<div style="display:grid;gap:16px">';
    order.q.forEach(function (i, pos) {
      var q = l.quiz[i];
      var mark = results ? (results[i] ? '<span class="mark ok" role="img" aria-label="Correct">✓</span>' : '<span class="mark no" role="img" aria-label="Wrong">✗</span>') : '';
      h += '<fieldset class="card q-card" style="margin:0"><legend class="small muted" style="padding:0 4px">Question ' + (pos + 1) + ' of ' + l.quiz.length + '</legend><div class="qtext">' + mark + ' ' + mdInline(q.q) + '</div><div class="opts">';
      order.o[i].forEach(function (k) {
        var o = q.options[k];
        var chosen = String(answers[i]) === String(k);
        var cls = '';
        if (reveal) { if (reveal.answers[i] === k) cls = ' correct'; else if (chosen) cls = ' wrong'; }
        else if (results && chosen) cls = results[i] ? ' correct' : ' wrong';
        h += '<label class="opt' + cls + '"><input type="radio" name="q' + i + '" value="' + k + '"' + (chosen ? ' checked' : '') + (locked ? ' disabled' : '') + '><span>' + mdInline(o) + '</span></label>';
      });
      h += '</div>' + (reveal && reveal.explains[i] ? '<div class="explain"><b>Why:</b> ' + mdInline(reveal.explains[i]) + '</div>' : '') + '</fieldset>';
    });
    return h + '</div>';
  }
  function quizOrder(l, attemptNo) {
    var rnd = seeded(S.user.login + '|' + l.id + '|' + attemptNo);
    return { q: shuffled(l.quiz.length, rnd), o: l.quiz.map(function (q) { return shuffled(q.options.length, rnd); }) };
  }
  function startQuiz(lid) {
    var l = lessonById(lid);
    api('quizStart', { lessonId: lid }).then(function (r) {
      if (!r.ok) { if (r.progress) S.progress = r.progress; toast(r.error); render(); return; }
      var skew = Date.now() / 1000 - Date.parse(r.now) / 1000;
      var rec = S.progress[lid + '-Q'] || { score: 0, max: l.quiz.length, attempts: 0, detail: {} };
      rec.attempts = r.attempts != null ? r.attempts : rec.attempts;
      rec.detail = Object.assign({}, rec.detail, { startedAt: r.startedAt });
      S.progress[lid + '-Q'] = rec;
      S.quizRun = S.quizRun || {};
      S.quizRun[lid] = { deadline: Date.parse(r.startedAt) / 1000 + skew + r.minutes * 60, order: quizOrder(l, rec.attempts + 1), away: 0 };
      delete S.quizResult[lid];
      render(); window.scrollTo(0, 0);
    });
  }
  function submitQuiz(lid, auto) {
    var l = lessonById(lid), run = S.quizRun[lid], form = document.getElementById('quiz-form'), ans = [];
    l.quiz.forEach(function (q, i) { var el = form && form.querySelector('input[name="q' + i + '"]:checked'); ans.push(el ? Number(el.value) : (S.quizDraft[lid] && S.quizDraft[lid][i] != null ? Number(S.quizDraft[lid][i]) : null)); });
    var btn = form && form.querySelector('button[type=submit]'); if (btn) { btn.disabled = true; btn.textContent = 'Sending…'; }
    run.sending = true;
    api('quiz', { lessonId: lid, answers: ans, away: run.away }).then(function (r) {
      run.sending = false;
      if (!r.ok) {
        if (r.network) {
          // No internet: try again in 10 seconds, show the message once.
          run.retryAt = Date.now() + 10000;
          if (!run.warned) { run.warned = true; toast(r.error); }
          if (btn) { btn.disabled = false; btn.textContent = 'Submit answers'; }
          return;
        }
        delete S.quizRun[lid]; delete S.quizDraft[lid];
        if (r.progress) S.progress = r.progress;
        toast(r.error);
        api('state').then(function (st) { if (st.ok) { S.progress = st.progress || {}; S.settings = st.settings; } render(); });
        return;
      }
      S.quizResult[lid] = { score: r.score, max: r.max, results: r.results, reveal: r.reveal, attemptsLeft: r.attemptsLeft, answers: ans, order: run.order, late: r.late };
      var prev = S.progress[lid + '-Q'] || { detail: {} };
      var d = Object.assign({}, prev.detail, { last: r.score, lastAnswers: ans }); delete d.startedAt;
      S.progress[lid + '-Q'] = { kind: 'quiz', score: r.best, max: r.max, attempts: r.attempts, detail: d };
      delete S.quizDraft[lid]; delete S.quizRun[lid];
      if (auto) toast('Time is over. Your answers were sent.');
      render(); window.scrollTo(0, 0);
    });
  }
  function mdInline(s) {
    if (window.marked && window.marked.parseInline) return window.marked.parseInline(String(s));
    return esc(s);
  }

  function viewProjects() {
    return shell('P', 'Projects', '<section class="lesson-head"><div class="eyebrow">Graded by your teacher</div><h1>Projects</h1><p class="lead">Four bigger tasks that join many skills together. Each is graded out of 100. The last one, the capstone, you defend in person in class.</p></section>' + projectsList(), { active: 'projects' });
  }
  function viewProject(id) {
    var p = S.course.projects.filter(function (x) { return x.id === id; })[0];
    if (!p) return shell('#REF!', 'Not found', '<div class="card">Project not found.</div>');
    var open = projectOpen(p);
    var h = '<section class="lesson-head"><div class="eyebrow">' + esc(p.id) + ' · Due: ' + esc(p.due) + '</div><h1>' + esc(p.title) + '</h1><p class="lead">' + esc(p.summary) + '</p>' +
      '<div class="row">' + p.skills.map(function (s) { return '<span class="pill accent">' + esc(s) + '</span>'; }).join('') + '</div></section>';
    if (!open) return shell(p.id, p.title, h + '<div class="card notice">This project opens after you finish lesson ' + p.openAfter + '.</div>', { active: 'projects' });
    if (p.file) h += '<div>' + fileLink(p.file) + '</div>';
    if (p.personal) h += '<section class="card prose personal"><div class="eyebrow">Your personal task · only you have these values</div>' + md(p.personal) + '</section>';
    h += '<article class="card prose">' + md(p.brief) + '</article>';
    h += '<section class="card"><h2 style="font-size:20px;margin-bottom:10px">How it is graded</h2><div class="table-wrap"><table class="grid"><thead><tr><th>Criterion</th><th>Points</th></tr></thead><tbody>' +
      p.rubric.map(function (r) { return '<tr><td>' + esc(r.criterion) + '</td><td class="num">' + r.points + '</td></tr>'; }).join('') + '<tr><th>Total</th><th class="num">100</th></tr></tbody></table></div></section>';
    var r = S.progress[p.id];
    var graded = r && r.detail && r.detail.grade !== undefined && r.detail.grade !== '';
    h += '<section class="card"><h2 style="font-size:20px;margin-bottom:10px">Your submission</h2>';
    if (graded) {
      h += '<div class="result-banner"><b>' + esc(r.detail.grade) + ' / 100</b><div><div>Graded by your teacher</div>' + (r.detail.feedback ? '<div class="muted">' + esc(r.detail.feedback) + '</div>' : '') + '</div></div><p><a href="' + esc(r.detail.link) + '" target="_blank" rel="noopener">Your submitted file</a></p>';
    } else {
      if (r) h += '<p class="pill warn">Submitted ' + esc(String(r.detail.submittedAt || '').slice(0, 10)) + '. You can update the link until your teacher grades it.</p>';
      h += '<form id="project-form" data-project="' + p.id + '" style="display:grid;gap:12px">' +
        '<label class="field" for="pj-link">Link to your Google Sheet (Share → Anyone with the link → Viewer)<input id="pj-link" type="url" required placeholder="https://docs.google.com/spreadsheets/d/…" value="' + esc(r ? r.detail.link : '') + '"></label>' +
        '<label class="field" for="pj-note">Note for your teacher (optional)<textarea id="pj-note" rows="3">' + esc(r ? r.detail.note : '') + '</textarea></label>' +
        '<div class="row"><button class="btn primary" type="submit">' + (r ? 'Update submission' : 'Submit project') + '</button><span id="pj-msg" class="error" role="alert"></span></div></form>';
    }
    h += '</section>';
    return shell(p.id, p.title, h, { active: 'projects' });
  }

  function viewAccount() {
    var h = '<section class="lesson-head"><h1>Account</h1><p class="lead">Logged in as <span class="mono">' + esc(S.user.login) + '</span></p></section>' +
      '<form class="card" id="pw-form" style="display:grid;gap:12px;max-width:420px"><h2 style="font-size:20px">Change password</h2>' +
      '<label class="field" for="pw-new">New password<input id="pw-new" type="password" autocomplete="new-password" minlength="6" required></label>' +
      '<label class="field" for="pw-new2">Repeat new password<input id="pw-new2" type="password" autocomplete="new-password" minlength="6" required></label>' +
      '<div class="row"><button class="btn primary" type="submit">Save password</button><span id="pw-msg" class="error" role="alert"></span></div></form>';
    return shell('ME', 'Account', h, { active: 'account' });
  }

  // ---------- teacher ----------
  function loadTeacher(force) {
    if (S.teacher && !force) return Promise.resolve(S.teacher);
    return api('overview').then(function (r) {
      if (!r.ok) throw new Error(r.error);
      var byUser = {};
      r.progress.forEach(function (p) { (byUser[p.login] = byUser[p.login] || {})[p.item] = p; });
      S.teacher = { users: r.users, students: r.users.filter(function (u) { return u.role !== 'teacher'; }), prog: byUser };
      S.settings = r.settings;
      return S.teacher;
    });
  }
  function teacherView(fn) {
    if (!isTeacher()) { go('#/home'); return null; }
    if (!S.teacher) {
      loadTeacher().then(render).catch(function (e) { app.innerHTML = shell('T1', 'Teacher', '<div class="card error">' + esc(e.message) + '</div>', {}); bindShell(); });
      return shell('T1', 'Teacher', '<div class="loading">Loading the gradebook…</div>', { active: 'teacher', wide: true });
    }
    return fn(S.teacher);
  }
  function cellClass(v) { return v == null ? 'c-none' : v >= 80 ? 'c-good' : v >= 50 ? 'c-warn' : 'c-bad'; }

  function viewTeacher() {
    return teacherView(function (T) {
      var opts = ''; for (var i = 0; i <= lessons().length; i++) opts += '<option value="' + i + '"' + (Number(S.settings.openLessons) === i ? ' selected' : '') + '>' + (i === 0 ? 'None' : 'Lessons 1–' + i) + '</option>';
      var aopts = ''; for (var a = 1; a <= 5; a++) aopts += '<option value="' + a + '"' + (Number(S.settings.quizAttempts) === a ? ' selected' : '') + '>' + a + '</option>';
      var mopts = ''; [5, 10, 15, 20, 30, 45].forEach(function (m) { mopts += '<option value="' + m + '"' + (Number(S.settings.quizMinutes || 15) === m ? ' selected' : '') + '>' + m + ' min</option>'; });
      var popts = ''; [0, 30, 40, 50, 60, 70, 80].forEach(function (m) { popts += '<option value="' + m + '"' + (Number(S.settings.passPercent == null ? 50 : S.settings.passPercent) === m ? ' selected' : '') + '>' + (m ? m + '%' : 'any score') + '</option>'; });
      var sopts = '<option value="0"' + (Number(S.settings.showAnswers) ? '' : ' selected') + '>No (safer)</option><option value="1"' + (Number(S.settings.showAnswers) ? ' selected' : '') + '>Yes, after last attempt</option>';
      var h = '<section class="lesson-head"><div class="eyebrow">Teacher · ' + esc(S.course.group) + ' · ' + T.students.length + ' students</div><h1>Gradebook</h1></section>' +
        '<section class="card row" style="justify-content:space-between">' +
        '<div class="row"><label class="field" for="set-open" style="min-width:170px">Lessons available<select id="set-open">' + opts + '</select></label>' +
        '<label class="field" for="set-att" style="min-width:110px">Quiz attempts<select id="set-att">' + aopts + '</select></label>' +
        '<label class="field" for="set-min" style="min-width:110px">Quiz time<select id="set-min">' + mopts + '</select></label>' +
        '<label class="field" for="set-pass" style="min-width:150px">Quiz score to go on<select id="set-pass">' + popts + '</select></label>' +
        '<label class="field" for="set-show" style="min-width:190px">Show correct answers<select id="set-show">' + sopts + '</select></label></div>' +
        '<div class="row"><button class="btn" data-act="refresh">Refresh</button><button class="btn" data-act="csv">Download CSV</button></div></section>';
      var subs = 0, waiting = 0;
      T.students.forEach(function (u) { S.course.projects.forEach(function (p) { var r = (T.prog[u.login] || {})[p.id]; if (r) { subs++; if (r.detail.grade === undefined || r.detail.grade === '') waiting++; } }); });
      if (waiting) h += '<a class="card notice" href="#/teacher/projects" style="text-decoration:none;color:var(--ink)"><b>' + waiting + ' project submission' + (waiting === 1 ? '' : 's') + '</b> waiting for your grade →</a>';
      h += '<p class="small muted" style="margin:0">Each student moves to the next lesson only after finishing the previous one (all practice tasks + quiz). "Lessons available" is the furthest anyone can go.</p>';
      h += '<div class="legend"><span><span class="pill good">80%+</span></span><span><span class="pill warn">50–79%</span></span><span><span class="pill bad">below 50%</span></span><span>Q = best quiz score, P = practice tasks solved, ⚠ = quiz done very fast or left the page 2+ times</span></div>';
      h += '<div class="gradebook-wrap" tabindex="0" role="region" aria-label="Gradebook table, scroll sideways"><table class="gradebook"><thead><tr><th class="name" rowspan="2">Student</th><th rowspan="2">Now on</th><th rowspan="2">Progress</th><th rowspan="2">Quiz avg</th>';
      S.course.projects.forEach(function (p) { h += '<th rowspan="2">' + p.id + '</th>'; });
      lessons().forEach(function (l) { h += '<th colspan="2" title="' + esc(l.title) + '">' + l.id + '</th>'; });
      h += '<th rowspan="2">Last login</th></tr><tr>';
      lessons().forEach(function () { h += '<th>Q</th><th>P</th>'; });
      h += '</tr></thead><tbody>';
      T.students.forEach(function (u) {
        var pr = T.prog[u.login] || {};
        var o = overall(pr);
        h += '<tr><td class="name"><a href="#/teacher/s/' + encodeURIComponent(u.login) + '">' + esc(u.name) + '</a> <span class="cellref">' + esc(u.login) + '</span></td>' +
          '<td class="mono">L' + ('0' + unlockedUpTo(pr)).slice(-2) + '</td>' +
          '<td class="' + cellClass(o.pct >= 1 ? o.pct : null) + '">' + o.pct + '%</td><td class="' + cellClass(o.quizAvg) + '">' + (o.quizAvg == null ? '—' : o.quizAvg + '%') + '</td>';
        S.course.projects.forEach(function (p) {
          var r = pr[p.id];
          h += r ? (r.detail.grade !== undefined && r.detail.grade !== '' ? '<td class="' + cellClass(Number(r.detail.grade)) + '">' + esc(r.detail.grade) + '</td>' : '<td class="c-warn" title="Submitted, not graded">sent</td>') : '<td class="c-none">—</td>';
        });
        lessons().forEach(function (l) {
          var st = lessonStats(l, pr);
          var flag = suspicious(st.quiz);
          h += '<td class="' + cellClass(st.quizPct) + '"' + (flag ? ' title="' + esc(flag) + '"' : '') + '>' + (st.quizPct == null ? '·' : st.quizPct) + (flag ? ' ⚠' : '') + '</td>' +
            '<td class="' + (st.solved === 0 ? 'c-none' : cellClass(pct(st.solved, st.total))) + '">' + (st.solved ? st.solved + '/' + st.total : '·') + '</td>';
        });
        h += '<td class="small">' + esc(fmtDate(u.lastLogin)) + '</td></tr>';
      });
      h += '</tbody></table></div>';
      return shell('T1', 'Gradebook', h, { active: 'teacher', wide: true });
    });
  }
  function suspicious(q) {
    var log = q && q.detail && q.detail.log || [];
    var out = [];
    log.forEach(function (e, i) {
      if (e.away >= 2) out.push('attempt ' + (i + 1) + ': left the page ' + e.away + ' times');
      if (e.sec != null && e.sec < 90 && e.score > 0) out.push('attempt ' + (i + 1) + ': finished in ' + fmtSec(e.sec));
    });
    return out.join('; ');
  }
  function fmtDate(s) { if (!s) return '—'; var d = new Date(s); return isNaN(d) ? String(s).slice(0, 16) : d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }); }

  function viewTeacherStudent(login) {
    return teacherView(function (T) {
      var u = T.users.filter(function (x) { return x.login === login; })[0];
      if (!u) return shell('#N/A', 'Student', '<div class="card">Student not found.</div>', { active: 'teacher' });
      var pr = T.prog[login] || {};
      var o = overall(pr);
      var h = '<section class="lesson-head"><div class="eyebrow"><a href="#/teacher">← Gradebook</a></div><h1>' + esc(u.name) + '</h1><p class="lead mono">' + esc(u.login) + (u.variant ? ' · variant ' + u.variant : '') + '</p></section>' +
        '<section class="card stats">' + stat('Progress', o.pct + '%', '') + stat('Quiz average', o.quizAvg == null ? '—' : o.quizAvg + '%', '') + stat('Practice solved', o.psolved + ' / ' + o.ptotal, '') +
        '<div class="stat"><span class="small muted">Password</span><b class="mono" style="font-size:20px">' + esc(u.password) + '</b><button class="btn small" data-act="resetpw" data-student="' + esc(login) + '">Make new password</button></div></section>';
      var upTo = unlockedUpTo(pr);
      h += '<section class="card row" style="justify-content:space-between"><div>Now on <b>lesson ' + upTo + '</b>' + (upTo >= Number(S.settings.openLessons) ? ' <span class="small muted">(the limit you set)</span>' : '') + '</div>' +
        (upTo < Number(S.settings.openLessons) ? '<button class="btn small" data-act="unlock" data-to="' + (upTo + 1) + '" data-student="' + esc(login) + '">Open lesson ' + (upTo + 1) + ' for this student</button>' : '') + '</section>';
      h += '<section class="card"><h2 style="font-size:20px;margin-bottom:10px">Lessons</h2><div class="table-wrap"><table class="grid"><thead><tr><th>Lesson</th><th>Quiz best</th><th>Quiz attempts (score · time · left page)</th><th>Practice</th><th></th></tr></thead><tbody>';
      lessons().forEach(function (l) {
        var st = lessonStats(l, pr);
        var log = st.quiz && st.quiz.detail && st.quiz.detail.log || [];
        var logHtml = log.map(function (e) { var bad = e.away >= 2 || (e.sec != null && e.sec < 90 && e.score > 0); return '<div class="small' + (bad ? ' error' : '') + '">' + e.score + ' · ' + fmtSec(e.sec) + ' · ' + (e.away == null ? '—' : e.away + '×') + (e.note ? ' · ' + esc(e.note) : '') + '</div>'; }).join('') || (st.quiz ? st.quiz.attempts : '—');
        h += '<tr><td><span class="cellref">' + l.id + '</span> ' + esc(l.title) + '</td><td class="num">' + (st.quiz ? st.quiz.score + '/' + st.quiz.max : '—') + '</td><td class="num">' + logHtml + '</td><td class="num">' + st.solved + '/' + st.total + '</td><td>' +
          (st.quiz && st.quiz.attempts ? '<button class="btn small" data-act="resetquiz" data-student="' + esc(login) + '" data-lesson="' + l.id + '">Give new attempts</button>' : '') + '</td></tr>';
      });
      h += '</tbody></table></div></section>';
      h += '<section class="card"><h2 style="font-size:20px;margin-bottom:6px">Projects</h2>' + S.course.projects.map(function (p) { return submissionRow(u, p, pr[p.id]); }).join('') + '</section>';
      return shell(login, u.name, h, { active: 'teacher' });
    });
  }
  function submissionRow(u, p, r) {
    if (!r) return '<div class="sub-row"><div><b>' + esc(p.title) + '</b><div class="small muted">' + esc(u.name) + ' · not submitted</div></div></div>';
    var key = u.login + '|' + p.id;
    return '<div class="sub-row"><div style="min-width:0"><b>' + esc(p.id) + ' · ' + esc(u.name) + '</b> <span class="cellref">' + esc(u.login) + '</span>' +
      '<div style="overflow-wrap:anywhere"><a href="' + esc(r.detail.link) + '" target="_blank" rel="noopener">' + esc(r.detail.link) + '</a></div>' +
      (r.detail.note ? '<div class="small muted">Note: ' + esc(r.detail.note) + '</div>' : '') +
      '<div class="small muted">Submitted ' + esc(fmtDate(r.detail.submittedAt)) + '</div></div>' +
      '<form class="grade-form" data-grade="' + esc(key) + '"><input type="number" min="0" max="100" placeholder="0–100" aria-label="Grade" name="grade" value="' + esc(r.detail.grade == null ? '' : r.detail.grade) + '">' +
      '<input type="text" name="feedback" placeholder="Feedback for the student" aria-label="Feedback" value="' + esc(r.detail.feedback || '') + '"><button class="btn primary small" type="submit">Save grade</button></form></div>';
  }
  function viewTeacherProjects() {
    return teacherView(function (T) {
      var h = '<section class="lesson-head"><div class="eyebrow">Teacher</div><h1>Grade projects</h1><p class="lead">Open each link, grade with the rubric, and write a short comment. Students see the grade and comment.</p></section>';
      S.course.projects.forEach(function (p) {
        var rows = T.students.map(function (u) { return { u: u, r: (T.prog[u.login] || {})[p.id] }; }).filter(function (x) { return x.r; });
        rows.sort(function (a, b) { var ga = a.r.detail.grade === undefined || a.r.detail.grade === '' ? 0 : 1, gb = b.r.detail.grade === undefined || b.r.detail.grade === '' ? 0 : 1; return ga - gb; });
        h += '<section class="card"><div class="row" style="justify-content:space-between"><h2 style="font-size:20px">' + esc(p.title) + '</h2><span class="pill">' + rows.length + ' / ' + T.students.length + ' submitted</span></div>' +
          '<details class="hint" style="margin:8px 0"><summary>Rubric</summary><ul>' + p.rubric.map(function (r) { return '<li>' + esc(r.criterion) + ' — <b>' + r.points + '</b></li>'; }).join('') + '</ul></details>' +
          (rows.length ? rows.map(function (x) { return submissionRow(x.u, p, x.r); }).join('') : '<p class="muted">No submissions yet.</p>') + '</section>';
      });
      return shell('T2', 'Grade projects', h, { active: 'tprojects' });
    });
  }
  function viewTeacherAccounts() {
    return teacherView(function (T) {
      var h = '<section class="lesson-head"><div class="eyebrow">Teacher</div><h1>Student accounts</h1><p class="lead">To change names or add college emails, edit the <b>Students</b> tab of your Google Sheet (name in column C, email in column F). Students with an email can use “Sign in with Google”.</p></section>' +
        '<div class="gradebook-wrap" tabindex="0" role="region" aria-label="Gradebook table, scroll sideways"><table class="gradebook"><thead><tr><th>#</th><th class="name">Name</th><th>Login</th><th>College email</th><th>Password</th><th>Last login</th><th></th></tr></thead><tbody>' +
        T.students.map(function (u, i) {
          return '<tr><td>' + (i + 1) + '</td><td class="name">' + esc(u.name) + '</td><td class="mono">' + esc(u.login) + '</td><td class="mono">' + (u.email ? esc(u.email) : '<span class="muted">—</span>') + '</td><td class="mono">' + esc(u.password) + '</td><td class="small">' + esc(fmtDate(u.lastLogin)) + '</td><td><button class="btn small" data-act="resetpw" data-student="' + esc(u.login) + '">New password</button></td></tr>';
        }).join('') + '</tbody></table></div>';
      return shell('T3', 'Student accounts', h, { active: 'taccounts', wide: true });
    });
  }
  function exportCsv() {
    var T = S.teacher;
    var head = ['Name', 'Login', 'Progress %', 'Quiz avg %', 'Practice solved'].concat(S.course.projects.map(function (p) { return p.id; }));
    lessons().forEach(function (l) { head.push(l.id + ' quiz %', l.id + ' practice'); });
    var rows = [head];
    T.students.forEach(function (u) {
      var pr = T.prog[u.login] || {}, o = overall(pr);
      var r = [u.name, u.login, o.pct, o.quizAvg == null ? '' : o.quizAvg, o.psolved];
      S.course.projects.forEach(function (p) { var x = pr[p.id]; r.push(x ? (x.detail.grade === undefined ? 'submitted' : x.detail.grade) : ''); });
      lessons().forEach(function (l) { var st = lessonStats(l, pr); r.push(st.quizPct == null ? '' : st.quizPct, st.solved + '/' + st.total); });
      rows.push(r);
    });
    var csv = '﻿' + rows.map(function (r) { return r.map(function (v) { v = String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(','); }).join('\n');
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
    a.download = 'gradebook-' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.appendChild(a); a.click(); a.remove();
  }

  // ---------- routing ----------
  function render() {
    if (!S.course) return;
    if (!S.user) { if (!document.getElementById('login-form')) viewLogin(); return; }
    var parts = (location.hash || '#/home').replace(/^#\/?/, '').split('/');
    var html;
    S.menuOpen = false;
    switch (parts[0]) {
      case 'lesson': html = viewLesson(parts[1], parts[2]); break;
      case 'projects': html = viewProjects(); break;
      case 'project': html = viewProject(parts[1]); break;
      case 'account': html = viewAccount(); break;
      case 'teacher':
        html = parts[1] === 's' ? viewTeacherStudent(decodeURIComponent(parts[2] || '')) : parts[1] === 'projects' ? viewTeacherProjects() : parts[1] === 'accounts' ? viewTeacherAccounts() : viewTeacher();
        break;
      default: html = viewHome();
    }
    if (html == null) return;
    app.innerHTML = html;
    bindShell();
    var keep = S.scrollTo; S.scrollTo = null;
    if (keep) { var el = document.getElementById(keep); if (el) el.scrollIntoView({ block: 'start' }); }
    else if (S.lastHash !== location.hash) window.scrollTo(0, 0);
    if (S.focusedHash !== location.hash) {
      // Move keyboard / screen-reader focus to the new page's heading (once it has loaded).
      var h1 = app.querySelector('main h1, h1');
      if (h1) { h1.setAttribute('tabindex', '-1'); h1.focus({ preventScroll: true }); S.focusedHash = location.hash; }
    }
    S.lastHash = location.hash;
  }

  function bindShell() {
    var quiz = document.getElementById('quiz-form');
    if (quiz) {
      var lid = quiz.getAttribute('data-lesson');
      quiz.addEventListener('change', function () {
        var d = {}; new FormData(quiz).forEach(function (v, k) { d[k.slice(1)] = v; }); S.quizDraft[lid] = d;
      });
      ['copy', 'cut', 'contextmenu'].forEach(function (ev) { quiz.addEventListener(ev, function (e) { e.preventDefault(); }); });
      quiz.addEventListener('submit', function (e) {
        e.preventDefault();
        var l = lessonById(lid), missing = 0;
        l.quiz.forEach(function (q, i) { if (!quiz.querySelector('input[name="q' + i + '"]:checked')) missing++; });
        var msg = document.getElementById('quiz-msg');
        if (missing && !quiz.dataset.confirm) { quiz.dataset.confirm = '1'; msg.textContent = missing + ' question(s) have no answer. Press Submit again to send anyway.'; return; }
        submitQuiz(lid, false);
      });
    }
    document.querySelectorAll('form.checks').forEach(function (f) {
      f.addEventListener('submit', function (e) {
        e.preventDefault();
        var tid = f.getAttribute('data-task'), answers = {};
        new FormData(f).forEach(function (v, k) { answers[k] = String(v).trim(); });
        var btn = f.querySelector('button[type=submit]'); btn.disabled = true; btn.textContent = 'Checking…';
        api('practice', { taskId: tid, answers: answers }).then(function (r) {
          if (!r.ok) { btn.disabled = false; btn.textContent = 'Check'; toast(r.error); return; }
          S.practiceResult[tid] = { results: r.results, score: r.score, max: r.max, answers: answers };
          var prev = S.progress[tid] || { score: 0, attempts: 0, detail: {} };
          var correct = Object.assign({}, prev.detail.correct || {});
          Object.keys(r.results).forEach(function (k) { if (r.results[k]) correct[k] = answers[k]; });
          S.progress[tid] = { kind: 'practice', score: Math.max(prev.score || 0, r.score, Object.keys(correct).length), max: r.max, attempts: (prev.attempts || 0) + 1, detail: { correct: correct } };
          S.scrollTo = 'task-' + tid;
          render();
        });
      });
    });
    var pf = document.getElementById('project-form');
    if (pf) pf.addEventListener('submit', function (e) {
      e.preventDefault();
      var pid = pf.getAttribute('data-project');
      api('project', { projectId: pid, link: document.getElementById('pj-link').value, note: document.getElementById('pj-note').value }).then(function (r) {
        if (!r.ok) { document.getElementById('pj-msg').textContent = r.error; return; }
        S.progress = r.progress; toast('Project submitted'); render();
      });
    });
    var pw = document.getElementById('pw-form');
    if (pw) pw.addEventListener('submit', function (e) {
      e.preventDefault();
      var a = document.getElementById('pw-new').value.trim(), b = document.getElementById('pw-new2').value.trim(), m = document.getElementById('pw-msg');
      if (a !== b) { m.textContent = 'The two passwords are different.'; return; }
      api('changePassword', { newPassword: a }).then(function (r) {
        if (!r.ok) { m.textContent = r.error; return; }
        S.creds.password = a; store(CRED_KEY, S.creds); toast('Password changed'); pw.reset(); m.textContent = '';
      });
    });
    document.querySelectorAll('form.grade-form').forEach(function (f) {
      f.addEventListener('submit', function (e) {
        e.preventDefault();
        var key = f.getAttribute('data-grade').split('|');
        var grade = f.elements.grade.value, feedback = f.elements.feedback.value;
        api('grade', { student: key[0], projectId: key[1], grade: grade, feedback: feedback }).then(function (r) {
          if (!r.ok) { toast(r.error); return; }
          var rec = S.teacher.prog[key[0]][key[1]];
          rec.detail.grade = grade === '' ? '' : Math.max(0, Math.min(100, Number(grade))); rec.detail.feedback = feedback;
          toast('Grade saved'); render();
        });
      });
    });
    var setKeys = { 'set-open': 'openLessons', 'set-att': 'quizAttempts', 'set-min': 'quizMinutes', 'set-pass': 'passPercent', 'set-show': 'showAnswers' };
    Object.keys(setKeys).forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.addEventListener('change', function () {
        var payload = {}; payload[setKeys[id]] = el.value;
        api('settings', payload).then(function (r) { if (!r.ok) { toast(r.error); return; } S.settings = r.settings; toast('Saved'); render(); });
      });
    });
  }

  document.addEventListener('click', function (e) {
    var t = e.target.closest('[data-act],[data-go]');
    if (!t) { if (S.menuOpen && !e.target.closest('.side')) { S.menuOpen = false; var s = document.querySelector('.side'); if (s) s.classList.remove('open'); } return; }
    if (t.dataset.go) { go(t.dataset.go); return; }
    var act = t.dataset.act;
    if (act === 'menu') { S.menuOpen = !S.menuOpen; document.querySelector('.side').classList.toggle('open', S.menuOpen); t.setAttribute('aria-expanded', String(S.menuOpen)); if (S.menuOpen) { var f = document.querySelector('.side a, .side button'); if (f) f.focus(); } }
    if (act === 'logout') { store(CRED_KEY, null); S.user = null; S.creds = null; S.progress = {}; if (S.baseCourse) S.course = S.baseCourse; S.teacher = null; S.quizResult = {}; S.practiceResult = {}; S.quizRun = null; S.quizDraft = {}; S.quizReview = {}; document.querySelectorAll('.toast').forEach(function (x) { x.remove(); }); go('#/login'); }
    if (act === 'quizclose') {
      // Records the attempt whose time ran out, so the next lesson can open.
      t.disabled = true;
      api('quizStart', { lessonId: t.dataset.lesson }).then(function () { return api('state'); }).then(function (st) { if (st.ok) { S.progress = st.progress || {}; S.settings = st.settings; } render(); });
    }
    if (act === 'quizstart') {
      if (t.dataset.confirm !== '1') { t.dataset.confirm = '1'; t.textContent = 'The timer starts now. Click again to begin'; return; }
      t.disabled = true; startQuiz(t.dataset.lesson);
    }
    if (act === 'unlock') {
      if (t.disabled) return; t.disabled = true;
      api('unlockNext', { student: t.dataset.student, to: Number(t.dataset.to) }).then(function (r) {
        if (!r.ok) { t.disabled = false; toast(r.error); return; }
        loadTeacher(true).then(function () { toast('Opened up to lesson ' + r.unlockedUpTo); render(); });
      });
    }
    if (act === 'refresh') { loadTeacher(true).then(function () { toast('Updated'); render(); }); }
    if (act === 'csv') exportCsv();
    if (act === 'resetpw') {
      if (t.dataset.confirm !== '1') { t.dataset.confirm = '1'; t.textContent = 'Click again to confirm'; return; }
      api('resetPassword', { student: t.dataset.student }).then(function (r) {
        if (!r.ok) { toast(r.error); return; }
        S.teacher.users.forEach(function (u) { if (u.login === t.dataset.student) u.password = r.password; });
        toast('New password: ' + r.password); render();
      });
    }
    if (act === 'resetquiz') {
      api('resetQuiz', { student: t.dataset.student, lessonId: t.dataset.lesson }).then(function (r) {
        if (!r.ok) { toast(r.error); return; }
        loadTeacher(true).then(function () { toast('Attempts reset'); render(); });
      });
    }
  });
  window.addEventListener('hashchange', render);
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && S.menuOpen) { S.menuOpen = false; var sd = document.querySelector('.side'); if (sd) sd.classList.remove('open'); var mb = document.querySelector('.menu-btn'); if (mb) { mb.setAttribute('aria-expanded', 'false'); mb.focus(); } }
  });
  function activeRun() { if (!S.quizRun) return null; var m = /^#\/lesson\/(L\d\d)\/quiz/.exec(location.hash); return m && S.quizRun[m[1]] ? m[1] : null; }
  setInterval(function () {
    if (!S.quizRun) return;
    Object.keys(S.quizRun).forEach(function (lid) {
      var run = S.quizRun[lid];
      var left = run.deadline - Date.now() / 1000;
      var el = document.getElementById('quiz-left');
      if (el && activeRun() === lid) { el.textContent = fmtSec(left); document.getElementById('quiz-timer').classList.toggle('low', left < 60); }
      if (left <= 0 && !run.sending && !(run.retryAt && Date.now() < run.retryAt)) { if (activeRun() !== lid) location.hash = '#/lesson/' + lid + '/quiz'; setTimeout(function () { if (S.quizRun && S.quizRun[lid] && !S.quizRun[lid].sending) submitQuiz(lid, true); }, 50); }
    });
  }, 500);
  function markAway() { var lid = activeRun(); if (lid && !S.quizRun[lid].sending) { S.quizRun[lid].away++; } }
  document.addEventListener('visibilitychange', function () { if (document.hidden) markAway(); });
  window.addEventListener('blur', function () { if (!document.hidden) markAway(); });

  // ---------- boot ----------
  (window.COURSE_DATA ? Promise.resolve(window.COURSE_DATA) : fetch(CONFIG.courseUrl || 'data/course.json').then(function (r) { return r.json(); })).then(function (course) {
    S.course = course;
    document.title = course.title;
    var saved = store(CRED_KEY);
    if (saved && saved.login) {
      S.creds = saved;
      app.innerHTML = '<div class="loading">Loading…</div>';
      api('login').then(function (r) {
        if (r.ok) { S.user = r.user; S.progress = r.progress || {}; S.settings = r.settings; loadVariant().then(render); }
        else if (r.network) { app.innerHTML = '<div class="loading">No connection to the server. Check your internet and refresh the page.</div>'; }
        else { S.creds = null; store(CRED_KEY, null); viewLogin(); }
      });
    } else viewLogin();
  }).catch(function () { app.innerHTML = '<div class="loading">Could not load the course. Refresh the page.</div>'; });
})();
