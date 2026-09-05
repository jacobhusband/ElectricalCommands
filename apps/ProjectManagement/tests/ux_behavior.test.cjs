const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'script.js'), 'utf8');
function extract(name) {
  const start = source.indexOf(`function ${name}(`);
  assert.ok(start >= 0, name);
  const open = source.indexOf('{', start);
  let depth = 1;
  let end = open + 1;
  while (depth && end < source.length) {
    if (source[end] === '{') depth++;
    if (source[end] === '}') depth--;
    end++;
  }
  return source.slice(start, end);
}
function context(extra = {}) {
  const scope = vm.createContext({
    Date,
    isFinished: (d) => (d.statuses || []).some(s => ['Complete', 'Delivered'].includes(s)),
    ...extra,
  });
  for (const name of ['parseDueStr', 'getWeekStartDate', 'getEffectiveDueStr', 'getHardDueStr',
    'deliverableNeedsAttention', 'filterPanelScheduleProjects', 'toggleProjectAttentionView']) {
    vm.runInContext(extract(name), scope);
  }
  return scope;
}

test('attention includes overdue and this-week work, but excludes finished, undated and future work', () => {
  const c = context();
  const now = new Date('2026-09-04T12:00:00');
  for (const due of ['2026-08-01', '2026-08-30', '2026-09-05']) {
    assert.equal(c.deliverableNeedsAttention({ due }, now), true, due);
  }
  for (const d of [{ due: '2026-09-06' }, {}, { due: 'invalid' },
    { due: '2026-09-01', statuses: ['Complete'] },
    { hardDue: '2026-09-01', statuses: ['Delivered'] }]) {
    assert.equal(c.deliverableNeedsAttention(d, now), false);
  }
});

test('a hard deadline cannot be hidden by a later internal target', () => {
  const c = context();
  const now = new Date('2026-09-04T12:00:00');
  assert.equal(c.deliverableNeedsAttention({ hardDue: '2026-09-04' }, now), true);
  assert.equal(c.deliverableNeedsAttention({ due: '2026-10-01', hardDue: '2026-09-01' }, now), true);
  assert.equal(c.deliverableNeedsAttention({ due: '2026-09-01', hardDue: '2026-10-01' }, now), true);
  assert.equal(c.deliverableNeedsAttention({ hardDue: '09/04/26' }, now), true);
});

test('attention follows the local week across a year boundary', () => {
  const c = context();
  const now = new Date('2027-01-01T10:00:00');
  assert.equal(c.deliverableNeedsAttention({ due: '2026-12-28' }, now), true);
  assert.equal(c.deliverableNeedsAttention({ due: '2027-01-02' }, now), true);
  assert.equal(c.deliverableNeedsAttention({ due: '2027-01-03' }, now), false);
});

test('leaving attention restores filters, sort and board week', () => {
  let renders = 0;
  const week = new Date('2026-08-02T00:00:00');
  const c = context({ dueFilter: 'future', statusFilter: 'Waiting', deliverablesFilter: 'incomplete',
    currentSort: { key: 'name', dir: 'desc' }, projectCardWeek: week,
    projectAttentionPreviousFilters: null, resetProjectsListPagination() {}, render() { renders++; } });
  c.toggleProjectAttentionView();
  assert.equal(c.dueFilter, 'attention');
  assert.equal(c.statusFilter, 'all');
  c.toggleProjectAttentionView();
  assert.equal(c.dueFilter, 'future');
  assert.equal(c.statusFilter, 'Waiting');
  assert.equal(c.deliverablesFilter, 'incomplete');
  assert.equal(c.currentSort.key, 'name');
  assert.equal(c.currentSort.dir, 'desc');
  assert.equal(c.projectCardWeek.getTime(), week.getTime());
  assert.equal(renders, 2);
});

test('project search supports ID, nickname, full address and combined case-insensitive terms', () => {
  const c = context();
  const options = [
    { id: '260487', label: 'Atherton', project: { name: '207 Atherton Ave, Atherton CA', nick: 'Home' } },
    { id: '260765', label: 'School', project: { name: 'San Mateo HS Pool' } },
  ];
  for (const q of ['260487', 'ATHERTON 207', 'Home']) {
    assert.deepEqual(Array.from(c.filterPanelScheduleProjects(options, q), x => x.id), ['260487']);
  }
  assert.equal(c.filterPanelScheduleProjects(options, 'missing').length, 0);
  assert.equal(c.filterPanelScheduleProjects(options, '  ').length, 2);
  assert.equal(options.length, 2);
});

test('searching project options never changes the current project binding', () => {
  const elements = {
    psmProjectSearch: { value: 'missing' }, psmProjectLabel: {}, psmProjectResultCount: {},
    psmProjectSelect: { children: [], replaceChildren() { this.children = []; }, appendChild(x) { this.children.push(x); } },
  };
  const c = context({
    document: { getElementById: id => elements[id] },
    panelScheduleManagerState: { projectId: '260487', projectPath: 'original-path' },
    getPanelScheduleManagerProjectOptions: () => [{ id: '260487', label: 'Atherton', project: {} }],
    el: (tag, props) => ({ tag, ...props }),
  });
  vm.runInContext(extract('renderPanelScheduleManagerProjects'), c);
  c.renderPanelScheduleManagerProjects();
  assert.equal(c.panelScheduleManagerState.projectId, '260487');
  assert.equal(c.panelScheduleManagerState.projectPath, 'original-path');
  assert.equal(elements.psmProjectSelect.disabled, true);
  assert.equal(elements.psmProjectLabel.textContent, '260487 · Atherton');
  elements.psmProjectSearch.value = '';
  c.renderPanelScheduleManagerProjects();
  assert.equal(elements.psmProjectSelect.value, '260487');
  assert.equal(elements.psmProjectSelect.disabled, false);
});
