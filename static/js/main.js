/**
 * Training Management System — main.js
 * ------------------------------------------------------------------------
 * Small, dependency-free (besides Bootstrap's JS bundle) helpers that make
 * the server-rendered templates feel dynamic:
 *
 *   1. Mobile sidebar drawer (open/close + overlay + auto-close on nav)
 *   2. Auto-dismissing alert / message banners
 *   3. Reusable show/hide password toggle ([data-password-field] wrapper)
 *   4. Client-side "quick search" filtering for CRUD list tables
 *   5. Button loading state on form submit (prevents double-submits)
 *   6. Bootstrap-style client-side validation for forms with .needs-validation
 *   7. Lightweight toast() helper for one-off JS-triggered notifications
 *
 * Every helper checks for its target elements before doing anything, so
 * this single file can safely be included on every page (base.html and
 * the standalone login page) without errors on pages that don't use a
 * given feature.
 * ------------------------------------------------------------------------
 */

(function () {
  'use strict';

  document.addEventListener('DOMContentLoaded', function () {
    initSidebarDrawer();
    initAutoDismissAlerts();
    initPasswordToggles();
    initTableSearch();
    initLoadingSubmit();
    initClientValidation();
    initImagePreview();
  });

  // ------------------------------------------------------------------
  // 1. Mobile sidebar drawer
  // ------------------------------------------------------------------
  function initSidebarDrawer() {
    var openBtn = document.getElementById('sidebarToggleBtn');
    var closeBtn = document.getElementById('sidebarCloseBtn');
    var overlay = document.getElementById('sidebarOverlay');
    var sidebar = document.getElementById('appSidebar');

    if (!sidebar) return;

    function open() {
      document.body.classList.add('sidebar-open');
    }
    function close() {
      document.body.classList.remove('sidebar-open');
    }

    if (openBtn) openBtn.addEventListener('click', open);
    if (closeBtn) closeBtn.addEventListener('click', close);
    if (overlay) overlay.addEventListener('click', close);

    // Close the drawer automatically when a nav link is tapped on mobile.
    sidebar.querySelectorAll('.nav-link').forEach(function (link) {
      link.addEventListener('click', close);
    });

    // Close on Escape for keyboard users.
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') close();
    });
  }

  // ------------------------------------------------------------------
  // 2. Auto-dismissing alerts (Django messages, form errors, etc.)
  // ------------------------------------------------------------------
  function initAutoDismissAlerts() {
    var alerts = document.querySelectorAll('.alert.alert-dismissible:not(.alert-persistent)');
    alerts.forEach(function (alertEl) {
      setTimeout(function () {
        fadeOutAndRemove(alertEl);
      }, 5000);
    });
  }

  function fadeOutAndRemove(el) {
    if (!el || !el.parentNode) return;
    el.style.transition = 'opacity 0.35s ease, transform 0.35s ease';
    el.style.opacity = '0';
    el.style.transform = 'translateY(-6px)';
    setTimeout(function () {
      if (el.parentNode) el.parentNode.removeChild(el);
    }, 350);
  }

  // ------------------------------------------------------------------
  // 3. Password visibility toggle
  //    Markup contract: a wrapper with class "password-field" containing
  //    a password <input> and a "button.toggle-password".
  //    (Used on the login page; reusable anywhere the same markup appears.)
  // ------------------------------------------------------------------
  function initPasswordToggles() {
    document.querySelectorAll('.password-field').forEach(function (wrap) {
      var input = wrap.querySelector('input');
      var btn = wrap.querySelector('.toggle-password');
      if (!input || !btn) return;

      btn.addEventListener('click', function () {
        var icon = btn.querySelector('i');
        var showing = input.type === 'text';
        input.type = showing ? 'password' : 'text';
        btn.setAttribute('aria-label', showing ? 'Show password' : 'Hide password');
        if (icon) {
          icon.classList.toggle('fa-eye', showing);
          icon.classList.toggle('fa-eye-slash', !showing);
        }
      });
    });
  }

  // ------------------------------------------------------------------
  // 4. Quick-search filtering for CRUD list tables
  //    Markup contract: <input id="tableSearch"> + a table inside
  //    ".table-panel" whose rows should be filtered.
  // ------------------------------------------------------------------
  function initTableSearch() {
    var input = document.getElementById('tableSearch');
    var panel = document.querySelector('.table-panel');
    if (!input || !panel) return;

    var emptyState = document.getElementById('tableSearchEmpty');
    var groups = Array.prototype.slice.call(panel.querySelectorAll('.trainer-group'));

    if (groups.length) {
      initGroupedTableSearch(input, groups, emptyState);
      return;
    }

    var table = panel.querySelector('table');
    if (!table) return;

    var rows = Array.prototype.slice.call(table.querySelectorAll('tbody tr[data-searchable]'));
    if (!rows.length) {
      rows = Array.prototype.slice.call(table.querySelectorAll('tbody tr'));
    }

    input.addEventListener('input', debounce(function () {
      var term = input.value.trim().toLowerCase();
      var visibleCount = 0;

      rows.forEach(function (row) {
        var text = row.textContent.toLowerCase();
        var matches = term === '' || text.indexOf(term) !== -1;
        row.classList.toggle('table-row-hidden', !matches);
        if (matches) visibleCount += 1;
      });

      if (emptyState) {
        emptyState.classList.toggle('d-none', visibleCount !== 0);
      }
    }, 150));
  }

  // ------------------------------------------------------------------
  // 4b. Quick-search for trainer-grouped ("accordion") CRUD list tables
  //    (Modules / Learning Outcomes / Indicative Contents). Each trainer's
  //    <details class="trainer-group"> is collapsed by default; searching
  //    auto-expands only the groups that contain a match and restores each
  //    group's original open/closed state once the search is cleared.
  //
  //    Learning Outcomes and Indicative Contents nest a second level of
  //    <details class="module-group"> accordions inside each trainer group,
  //    with the actual records rendered as [data-searchable] cards instead
  //    of table rows. The search logic below is written generically so it
  //    filters whatever leaf items exist (table rows or cards) and then
  //    resolves visibility bottom-up: module groups first, then trainers.
  // ------------------------------------------------------------------
  function initGroupedTableSearch(input, trainerGroups, emptyState) {
    var panel = trainerGroups[0].closest('.table-panel') || document;
    var moduleGroups = Array.prototype.slice.call(panel.querySelectorAll('.module-group'));

    var trainerWasOpen = trainerGroups.map(function (g) { return g.hasAttribute('open'); });
    var moduleWasOpen = moduleGroups.map(function (g) { return g.hasAttribute('open'); });

    input.addEventListener('input', debounce(function () {
      var term = input.value.trim().toLowerCase();

      // 1. Filter every leaf item (table row or record card) directly.
      var items = Array.prototype.slice.call(panel.querySelectorAll('[data-searchable]'));
      items.forEach(function (item) {
        var text = item.textContent.toLowerCase();
        var matches = term === '' || text.indexOf(term) !== -1;
        item.classList.toggle('table-row-hidden', !matches);
      });

      // 2. Resolve module-level groups (nested inside a trainer group).
      moduleGroups.forEach(function (group, idx) {
        var total = group.querySelectorAll('[data-searchable]').length;
        var visible = group.querySelectorAll('[data-searchable]:not(.table-row-hidden)').length;
        group.classList.toggle('d-none', total > 0 && visible === 0);
        if (total) {
          group.open = term === '' ? moduleWasOpen[idx] : visible > 0;
        }
      });

      // 3. Resolve trainer-level groups.
      var totalVisible = 0;
      trainerGroups.forEach(function (group, idx) {
        var total = group.querySelectorAll('[data-searchable]').length;
        var visible = group.querySelectorAll('[data-searchable]:not(.table-row-hidden)').length;
        group.classList.toggle('d-none', total > 0 && visible === 0);
        if (total) {
          group.open = term === '' ? trainerWasOpen[idx] : visible > 0;
        }
        totalVisible += visible;
      });

      if (emptyState) {
        emptyState.classList.toggle('d-none', totalVisible !== 0);
      }
    }, 150));
  }

  function debounce(fn, wait) {
    var timeout;
    return function () {
      var args = arguments;
      var context = this;
      clearTimeout(timeout);
      timeout = setTimeout(function () {
        fn.apply(context, args);
      }, wait);
    };
  }

  // ------------------------------------------------------------------
  // 5. Button loading state on submit
  //    Markup contract: <form class="js-loading-submit">
  //    Disables the submit button and shows a spinner so users can't
  //    double-submit create/edit/delete forms.
  // ------------------------------------------------------------------
  function initLoadingSubmit() {
    document.querySelectorAll('form.js-loading-submit').forEach(function (form) {
      form.addEventListener('submit', function (e) {
        // Let the browser's native / Bootstrap validation run first.
        if (form.classList.contains('needs-validation') && form.checkValidity && !form.checkValidity()) {
          return;
        }
        var submitBtn = form.querySelector('[type="submit"]');
        if (submitBtn && !submitBtn.classList.contains('is-loading')) {
          submitBtn.classList.add('is-loading');
          submitBtn.disabled = true;
        }
      });
    });
  }

  // ------------------------------------------------------------------
  // 6. Bootstrap-style client-side validation
  //    Markup contract: <form class="needs-validation" novalidate>
  // ------------------------------------------------------------------
  function initClientValidation() {
    document.querySelectorAll('form.needs-validation').forEach(function (form) {
      form.addEventListener('submit', function (event) {
        if (!form.checkValidity()) {
          event.preventDefault();
          event.stopPropagation();
        }
        form.classList.add('was-validated');
      }, false);
    });
  }

  // ------------------------------------------------------------------
  // 6b. Live preview for any image file input
  //    Markup contract: any <input type="file" accept="image/..."> gets
  //    an auto-inserted preview box right after it, updated on change.
  //    (No template changes required — this is fully generic, so it also
  //    works for future image-upload fields beyond Logo.)
  // ------------------------------------------------------------------
  function initImagePreview() {
    var inputs = document.querySelectorAll('input[type="file"][accept*="image"]');
    inputs.forEach(function (input) {
      input.addEventListener('change', function () {
        var file = input.files && input.files[0];
        var box = input.parentNode.querySelector('.image-preview-box');

        if (!file) {
          if (box) box.remove();
          return;
        }

        if (!box) {
          box = document.createElement('div');
          box.className = 'image-preview-box';
          var labelEl = document.createElement('div');
          labelEl.className = 'image-preview-label';
          labelEl.textContent = 'New image preview';
          var imgEl = document.createElement('img');
          box.appendChild(labelEl);
          box.appendChild(imgEl);
          input.insertAdjacentElement('afterend', box);
        }

        var img = box.querySelector('img');
        var reader = new FileReader();
        reader.onload = function (e) {
          img.src = e.target.result;
        };
        reader.readAsDataURL(file);
      });
    });
  }

  // ------------------------------------------------------------------
  // 7. Lightweight toast helper
  //    Usage: TMS.toast('Saved successfully', 'success')
  // ------------------------------------------------------------------
  function toast(message, variant) {
    variant = variant || 'success';
    var stack = document.querySelector('.tms-toast-stack');
    if (!stack) {
      stack = document.createElement('div');
      stack.className = 'tms-toast-stack';
      document.body.appendChild(stack);
    }

    var el = document.createElement('div');
    el.className = 'alert alert-' + variant + ' alert-dismissible fade show shadow-sm mb-0';
    el.setAttribute('role', 'alert');
    el.style.minWidth = '260px';
    el.innerHTML = message + '<button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>';

    stack.appendChild(el);
    setTimeout(function () {
      fadeOutAndRemove(el);
    }, 4000);
  }

  // Expose a tiny public API in case inline template scripts need it.
  window.TMS = {
    toast: toast,
  };
})();
