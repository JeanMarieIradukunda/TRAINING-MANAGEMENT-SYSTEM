/**
 * Training Management System — bulk-form.js
 * ------------------------------------------------------------------------
 * Powers the "add multiple" screens (Learning Outcomes / Indicative
 * Contents): lets an admin add and remove rows of a Django formset
 * client-side before submitting the whole batch in one POST.
 *
 * Markup contract, per formset on the page:
 *   <div data-bulk-formset="<prefix>">      the prefix matches the Django
 *                                            formset's `prefix=` kwarg
 *     <div data-bulk-row-list>               holds the row elements
 *       <div data-bulk-row>                  one form's fields
 *         <span data-bulk-row-index>          (optional) "1", "2", ... label
 *         ...fields...
 *         <button data-bulk-remove>           removes this row
 *     <template data-bulk-empty>             a blank row using Django's
 *                                             formset.empty_form, with the
 *                                             literal "__prefix__" left in
 *                                             every name/id attribute
 *     <button data-bulk-add>                 appends a new row
 *
 * Because every row here is a brand-new (unsaved) record, rows are simply
 * removed from the DOM and everything is renumbered — there's no need to
 * preserve index gaps the way you would for a formset of already-saved
 * instances.
 * ------------------------------------------------------------------------
 */

(function () {
  'use strict';

  document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('[data-bulk-formset]').forEach(initBulkFormset);
  });

  function initBulkFormset(container) {
    var prefix = container.getAttribute('data-bulk-formset');
    var rowList = container.querySelector('[data-bulk-row-list]');
    var addBtn = container.querySelector('[data-bulk-add]');
    var emptyTemplate = container.querySelector('template[data-bulk-empty]');
    var totalInput = container.querySelector('input[name="' + prefix + '-TOTAL_FORMS"]');

    if (!rowList || !addBtn || !emptyTemplate || !totalInput) return;

    function rows() {
      return Array.prototype.slice.call(rowList.querySelectorAll('[data-bulk-row]'));
    }

    // Re-derives every row's form-<index>-<field> name/id (and matching
    // label "for") from its current DOM position, and keeps TOTAL_FORMS in
    // sync so Django's formset parses exactly the rows still on screen.
    function renumber() {
      var list = rows();
      var pattern = new RegExp(prefix + '-(\\d+|__prefix__)-');

      list.forEach(function (row, idx) {
        row.querySelectorAll('[name], [id], label[for]').forEach(function (el) {
          ['name', 'id', 'for'].forEach(function (attr) {
            var value = el.getAttribute(attr);
            if (value && pattern.test(value)) {
              el.setAttribute(attr, value.replace(pattern, prefix + '-' + idx + '-'));
            }
          });
        });

        var indexEl = row.querySelector('[data-bulk-row-index]');
        if (indexEl) indexEl.textContent = idx + 1;
      });

      totalInput.value = list.length;
      updateRemoveButtons(list);
    }

    // Always keep at least one row so there's something to fill in.
    function updateRemoveButtons(list) {
      list.forEach(function (row) {
        var btn = row.querySelector('[data-bulk-remove]');
        if (btn) btn.disabled = list.length <= 1;
      });
    }

    function addRow() {
      var wrapper = document.createElement('div');
      wrapper.innerHTML = emptyTemplate.innerHTML.trim();
      var newRow = wrapper.firstElementChild;
      if (!newRow) return;

      rowList.appendChild(newRow);
      renumber();

      var firstField = newRow.querySelector('textarea, input:not([type="hidden"]):not([type="checkbox"]), select');
      if (firstField) firstField.focus();
    }

    addBtn.addEventListener('click', addRow);

    rowList.addEventListener('click', function (e) {
      var removeBtn = e.target.closest('[data-bulk-remove]');
      if (!removeBtn || removeBtn.disabled) return;

      var row = removeBtn.closest('[data-bulk-row]');
      if (row) row.remove();
      renumber();
    });

    renumber();
  }
})();
