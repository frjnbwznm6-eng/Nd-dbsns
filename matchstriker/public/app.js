// Small progressive enhancements. The site works without JavaScript.
(function () {
  // Confirm before destructive form posts.
  document.querySelectorAll('form[data-confirm]').forEach(function (form) {
    form.addEventListener('submit', function (e) {
      if (!window.confirm(form.dataset.confirm)) e.preventDefault();
    });
  });

  // Sign-up: only show matchmaker fields when that role is picked.
  var roleOnly = document.querySelectorAll('[data-role-only]');
  if (roleOnly.length) {
    var sync = function () {
      var picked = document.querySelector('input[name="role"]:checked');
      roleOnly.forEach(function (el) { el.hidden = !picked || picked.value !== el.dataset.roleOnly; });
    };
    document.querySelectorAll('input[name="role"]').forEach(function (r) { r.addEventListener('change', sync); });
    sync();
  }

  // Message thread: stay scrolled to the bottom and poll for new messages.
  var thread = document.getElementById('thread');
  if (thread) {
    var toBottom = function () { thread.scrollTop = thread.scrollHeight; };
    toBottom();
    var last = null;
    setInterval(function () {
      if (document.hidden) return;
      fetch(thread.dataset.poll, { credentials: 'same-origin' })
        .then(function (r) { return r.ok ? r.text() : null; })
        .then(function (htmlText) {
          if (htmlText && last !== null && htmlText !== last) {
            var atBottom = thread.scrollHeight - thread.scrollTop - thread.clientHeight < 60;
            thread.innerHTML = htmlText;
            if (atBottom) toBottom();
          }
          if (htmlText) last = htmlText;
        })
        .catch(function () {});
    }, 5000);
    var box = document.querySelector('.composer textarea');
    if (box) {
      box.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); if (box.value.trim()) box.form.submit(); }
      });
    }
  }
})();
