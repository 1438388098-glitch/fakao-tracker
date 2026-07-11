/* ═══════════════════════════════════════════
   Fakao Tracker — Global JS
   ═══════════════════════════════════════════ */

// ── Theme Toggle ──
(function () {
  const STORAGE_KEY = 'fakao-tracker:theme';
  const saved = localStorage.getItem(STORAGE_KEY) || 'light';
  document.documentElement.setAttribute('data-theme', saved);
  updateToggleIcon(saved);

  document.getElementById('themeToggle').addEventListener('click', function () {
    const current = document.documentElement.getAttribute('data-theme');
    const next = current === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem(STORAGE_KEY, next);
    updateToggleIcon(next);
  });

  function updateToggleIcon(theme) {
    document.getElementById('themeToggle').textContent = theme === 'dark' ? '☀️' : '🌙';
  }
})();

// ── Toast ──
function showToast(message, duration) {
  if (!duration) duration = 2000;
  var toast = document.getElementById('toast');
  toast.textContent = message;
  toast.classList.add('show');
  clearTimeout(toast._timeout);
  toast._timeout = setTimeout(function () {
    toast.classList.remove('show');
  }, duration);
}

// ── Checkin ──
function doCheckin(mood, notes) {
  if (!mood) mood = 0;
  if (!notes) notes = '';

  fetch('/api/checkin', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ mood: mood, notes: notes })
  })
  .then(function (r) { return r.json(); })
  .then(function (data) {
    if (data.ok) {
      showToast(data.repeated ? '今日已打卡 ✓' : '打卡成功！今天辛苦了 🎯');
      // Refresh UI state
      var done = document.getElementById('checkinDone');
      var wrapper = document.getElementById('checkinBtnWrapper');
      if (done && wrapper) {
        done.classList.add('show');
        wrapper.classList.add('hide');
      }
      // Refresh calendar markers
      var todayEl = document.querySelector('.week-day.today');
      if (todayEl) todayEl.classList.add('checked');
    } else {
      showToast('打卡失败: ' + (data.error || '未知错误'));
    }
  })
  .catch(function () {
    showToast('网络错误，请重试');
  });
}

function undoCheckin() {
  fetch('/api/checkin/undo', { method: 'POST' })
  .then(function (r) { return r.json(); })
  .then(function (data) {
    if (data.ok) {
      showToast('已取消打卡');
      var done = document.getElementById('checkinDone');
      var wrapper = document.getElementById('checkinBtnWrapper');
      if (done && wrapper) {
        done.classList.remove('show');
        wrapper.classList.remove('hide');
      }
      var todayEl = document.querySelector('.week-day.today');
      if (todayEl) todayEl.classList.remove('checked');
    }
  });
}

// ── Phase Tab Switching ──
function switchPhase(tabEl, phaseName) {
  // Update active tab
  document.querySelectorAll('.phase-tab').forEach(function (t) { t.classList.remove('active'); });
  tabEl.classList.add('active');
  // Update active content
  document.querySelectorAll('.phase-content').forEach(function (c) { c.classList.remove('active'); });
  var target = document.getElementById('phase-' + phaseName);
  if (target) target.classList.add('active');
}

// ── Mood Selector ──
function selectMood(btn, value) {
  document.querySelectorAll('.mood-btn').forEach(function (b) { b.classList.remove('selected'); });
  btn.classList.add('selected');
  btn.closest('.mood-selector').dataset.value = value;
}

// ── Page Transitions ──
(function(){
  document.querySelectorAll('.nav-links a').forEach(function(a){
    a.addEventListener('click',function(e){
      e.preventDefault();
      var main=document.querySelector('.main-container');
      if(main)main.classList.add('page-out');
      var href=this.getAttribute('href');
      setTimeout(function(){window.location=href},180);
    });
  });
})();
