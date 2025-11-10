const noop = () => {};

export function createMiniTaskManager({
  elements = {},
  constants = {},
  getCurrentEntryType = () => 'task',
  onDraftsChange = noop
}) {
  const {
    listEl = null,
    addButton = null,
    removeButton = null,
    unitButtons = []
  } = elements;

  const {
    miniTaskOptions = [],
    workingDaysPerWeek = 5,
    defaultSequenceSteps = [],
    defaultSequenceMinimumTotal = defaultSequenceSteps.length || 1,
    defaultSequenceTotalWeight = defaultSequenceSteps.reduce(
      (sum, step) => sum + (Number(step.weight) || 0),
      0
    )
  } = constants;

  let drafts = [];
  let unit = 'days';

  function notifyDraftChange() {
    if (typeof onDraftsChange === 'function') {
      onDraftsChange();
    }
  }

  function getMiniTaskDrafts() {
    return drafts.map((draft) => ({ ...draft }));
  }

  function setMiniTaskDrafts(nextDrafts, options = {}) {
    const { silent = false } = options;
    drafts = (nextDrafts || []).map(sanitizeMiniTaskDraft);
    renderMiniTaskInputs();
    updateMiniTaskToolbar();
    if (!silent) notifyDraftChange();
  }

  function resetMiniTaskState(options = {}) {
    const { silent = false } = options;
    unit = 'days';
    drafts = [];
    renderMiniTaskInputs();
    updateMiniTaskToolbar();
    if (!silent) notifyDraftChange();
  }

  function getMiniTaskUnit() {
    return unit;
  }

  function setMiniTaskUnit(nextUnit, options = {}) {
    const { silent = false } = options;
    if (getCurrentEntryType() !== 'task') return;
    if (nextUnit !== 'days' && nextUnit !== 'weeks') return;
    if (unit === nextUnit) return;
    unit = nextUnit;
    renderMiniTaskInputs();
    updateMiniTaskToolbar();
    if (!silent) notifyDraftChange();
  }

  function addMiniTaskRow(prefill = {}) {
    if (getCurrentEntryType() !== 'task') return;
    const normalized = { ...prefill };
    if (normalized.duration != null) {
      const raw = Number(normalized.duration);
      normalized.duration = Math.max(1, Math.round(Number.isFinite(raw) ? raw : 1));
    } else {
      normalized.duration = unit === 'weeks' ? workingDaysPerWeek : 1;
    }
    drafts = [...drafts, sanitizeMiniTaskDraft({ enabled: true, ...normalized })];
    renderMiniTaskInputs();
    updateMiniTaskToolbar();
    notifyDraftChange();
    focusMiniTask(drafts.length - 1, 'name');
  }

  function removeMiniTaskRow() {
    if (getCurrentEntryType() !== 'task') return;
    if (!drafts.length) return;
    drafts = drafts.slice(0, drafts.length - 1);
    renderMiniTaskInputs();
    updateMiniTaskToolbar();
    notifyDraftChange();
  }

  function handleEntryTypeChange() {
    updateMiniTaskToolbar();
  }

  function renderMiniTaskInputs() {
    if (!listEl) return;
    listEl.innerHTML = '';
    drafts.forEach((mini, index) => {
      const item = document.createElement('div');
      item.className = 'mini-task-item';
      item.dataset.index = String(index);
      if (!mini.enabled) item.classList.add('disabled');
      if (mini.locked) item.classList.add('locked');

      const toggleLabel = document.createElement('label');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = Boolean(mini.enabled);
      checkbox.addEventListener('change', () => {
        setMiniTaskEnabled(index, checkbox.checked);
      });
      const labelText = document.createElement('span');
      labelText.textContent = String(index + 1);
      toggleLabel.appendChild(checkbox);
      toggleLabel.appendChild(labelText);

      const nameSelect = document.createElement('select');
      nameSelect.disabled = !mini.enabled;
      const placeholderOption = document.createElement('option');
      placeholderOption.value = '';
      placeholderOption.textContent = 'Select description';
      nameSelect.appendChild(placeholderOption);
      miniTaskOptions.forEach((option) => {
        const optionEl = document.createElement('option');
        optionEl.value = option;
        optionEl.textContent = option;
        nameSelect.appendChild(optionEl);
      });
      nameSelect.value = miniTaskOptions.includes(mini.name) ? mini.name : '';
      nameSelect.addEventListener('change', () => {
        setMiniTaskName(index, nameSelect.value);
      });

      const isDelay = mini.enabled && mini.name === 'Delays';
      item.classList.toggle('delay-active', isDelay);

      const durationWrapper = document.createElement('div');
      durationWrapper.className = 'mini-task-duration';
      const durationInput = document.createElement('input');
      durationInput.type = 'number';
      durationInput.min = unit === 'weeks' ? '0.2' : '1';
      durationInput.step = unit === 'weeks' ? '0.5' : '1';
      durationInput.value = String(getDraftDurationInCurrentUnit(mini.duration));
      durationInput.disabled = !mini.enabled;
      const handleDurationChange = () => {
        setMiniTaskDuration(index, parseFloat(durationInput.value));
      };
      durationInput.addEventListener('change', handleDurationChange);
      durationInput.addEventListener('blur', handleDurationChange);
      const durationSuffix = document.createElement('span');
      durationSuffix.textContent = unit === 'weeks' ? 'weeks' : 'days';
      durationWrapper.appendChild(durationInput);
      durationWrapper.appendChild(durationSuffix);

      const durationControls = document.createElement('div');
      durationControls.className = 'mini-duration-controls';
      const upButton = document.createElement('button');
      upButton.type = 'button';
      upButton.className = 'mini-duration-btn increase';
      upButton.textContent = '＋';
      upButton.disabled = !mini.enabled;
      upButton.addEventListener('click', () => {
        setMiniTaskDuration(index, getDraftDurationInCurrentUnit(mini.duration) + 1);
      });
      const downButton = document.createElement('button');
      downButton.type = 'button';
      downButton.className = 'mini-duration-btn decrease';
      downButton.textContent = '−';
      downButton.disabled = !mini.enabled;
      downButton.addEventListener('click', () => {
        setMiniTaskDuration(index, Math.max(1, getDraftDurationInCurrentUnit(mini.duration) - 1));
      });
      durationControls.appendChild(upButton);
      durationControls.appendChild(downButton);
      durationWrapper.appendChild(durationControls);

      const moveButtons = document.createElement('div');
      moveButtons.className = 'mini-task-move-buttons';
      const moveUpBtn = document.createElement('button');
      moveUpBtn.type = 'button';
      moveUpBtn.className = 'mini-move-btn move-up';
      moveUpBtn.textContent = '↑';
      moveUpBtn.disabled = index === 0;
      moveUpBtn.addEventListener('click', () => moveMiniTaskDraft(index, -1));
      const moveDownBtn = document.createElement('button');
      moveDownBtn.type = 'button';
      moveDownBtn.className = 'mini-move-btn move-down';
      moveDownBtn.textContent = '↓';
      moveDownBtn.disabled = index === drafts.length - 1;
      moveDownBtn.addEventListener('click', () => moveMiniTaskDraft(index, 1));
      moveButtons.appendChild(moveUpBtn);
      moveButtons.appendChild(moveDownBtn);

      const statusBadge = document.createElement('span');
      statusBadge.className = 'mini-task-status';
      statusBadge.textContent = mini.enabled ? (mini.locked ? 'Locked' : 'Active') : 'Off';

      const delayExtras = document.createElement('div');
      delayExtras.className = 'mini-delay-extras';
      delayExtras.style.display = isDelay ? '' : 'none';
      const delayTextarea = document.createElement('textarea');
      delayTextarea.placeholder = 'Describe delay';
      delayTextarea.value = mini.delayDescription || '';
      delayTextarea.addEventListener('input', () => {
        setMiniTaskDelayDescription(index, delayTextarea.value);
      });
      const chargeWrapper = document.createElement('label');
      chargeWrapper.className = 'mini-delay-charge';
      const chargeCheckbox = document.createElement('input');
      chargeCheckbox.type = 'checkbox';
      chargeCheckbox.checked = Boolean(mini.chargeToClient);
      chargeCheckbox.addEventListener('change', () => {
        setMiniTaskCharge(index, chargeCheckbox.checked);
      });
      const chargeLabel = document.createElement('span');
      chargeLabel.textContent = 'Charge to client';
      chargeWrapper.appendChild(chargeCheckbox);
      chargeWrapper.appendChild(chargeLabel);
      delayExtras.appendChild(delayTextarea);
      delayExtras.appendChild(chargeWrapper);

      item.appendChild(toggleLabel);
      item.appendChild(nameSelect);
      item.appendChild(durationWrapper);
      item.appendChild(moveButtons);
      item.appendChild(statusBadge);
      item.appendChild(delayExtras);
      listEl.appendChild(item);
    });
  }

  function updateMiniTaskToolbar() {
    if (addButton) {
      addButton.disabled = getCurrentEntryType() !== 'task';
    }
    if (removeButton) {
      const disabled = getCurrentEntryType() !== 'task' || drafts.length === 0;
      removeButton.disabled = disabled;
    }
    unitButtons.forEach((button) => {
      const disabled = getCurrentEntryType() !== 'task';
      const isActive = button.dataset.unit === unit;
      button.classList.toggle('active', isActive);
      button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
      button.disabled = disabled;
      button.setAttribute('aria-disabled', disabled ? 'true' : 'false');
      if (disabled) {
        button.setAttribute('tabindex', '-1');
      } else {
        button.removeAttribute('tabindex');
      }
    });
  }

  function setMiniTaskEnabled(index, enabled) {
    if (!drafts[index]) return;
    if (enabled) {
      drafts[index] = sanitizeMiniTaskDraft({ ...drafts[index], enabled: true });
    } else {
      drafts[index] = sanitizeMiniTaskDraft({
        enabled: false,
        name: '',
        duration: 1,
        locked: false,
        delayDescription: '',
        chargeToClient: false
      });
    }
    renderMiniTaskInputs();
    updateMiniTaskToolbar();
    notifyDraftChange();
    if (enabled) focusMiniTask(index, 'name');
  }

  function setMiniTaskName(index, name) {
    if (!drafts[index]) return;
    const isDelay = name === 'Delays';
    drafts[index] = sanitizeMiniTaskDraft({
      ...drafts[index],
      name,
      delayDescription: isDelay ? drafts[index].delayDescription : '',
      chargeToClient: isDelay ? drafts[index].chargeToClient : false
    });
    renderMiniTaskInputs();
    notifyDraftChange();
    focusMiniTask(index, isDelay ? 'delayDescription' : 'name');
  }

  function setMiniTaskDuration(index, value) {
    if (!drafts[index]) return;
    let durationValue = Number(value);
    if (!Number.isFinite(durationValue) || durationValue < 1) durationValue = 1;
    const durationInDays = unit === 'weeks'
      ? durationValue * workingDaysPerWeek
      : durationValue;
    drafts[index] = sanitizeMiniTaskDraft({ ...drafts[index], duration: durationInDays });
    renderMiniTaskInputs();
    notifyDraftChange();
  }

  function moveMiniTaskDraft(index, direction) {
    const targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= drafts.length) return;
    const nextDrafts = [...drafts];
    const [moved] = nextDrafts.splice(index, 1);
    nextDrafts.splice(targetIndex, 0, moved);
    drafts = nextDrafts.map(sanitizeMiniTaskDraft);
    renderMiniTaskInputs();
    notifyDraftChange();
    focusMiniTask(targetIndex, 'name');
  }

  function setMiniTaskDelayDescription(index, value) {
    if (!drafts[index]) return;
    if (drafts[index].name !== 'Delays') {
      drafts[index].delayDescription = '';
      return;
    }
    drafts[index].delayDescription = value;
    notifyDraftChange();
  }

  function setMiniTaskCharge(index, checked) {
    if (!drafts[index]) return;
    if (drafts[index].name !== 'Delays') {
      drafts[index].chargeToClient = false;
      return;
    }
    drafts[index].chargeToClient = Boolean(checked);
    notifyDraftChange();
  }

  function focusMiniTask(index, field = 'name') {
    if (!listEl) return;
    requestAnimationFrame(() => {
      let selector = 'select';
      if (field === 'duration') selector = '.mini-task-duration input';
      if (field === 'delayDescription') selector = '.mini-delay-extras textarea';
      const target = listEl.querySelector(`.mini-task-item[data-index="${index}"] ${selector}`);
      if (target) target.focus();
    });
  }

  function getDraftDurationInCurrentUnit(durationDays) {
    const safeDays = Math.max(1, Math.round(Number.isFinite(durationDays) ? durationDays : 1));
    if (unit === 'weeks') {
      return Number((safeDays / workingDaysPerWeek).toFixed(2));
    }
    return safeDays;
  }

  function collectMiniTasks(taskId) {
    if (getCurrentEntryType() !== 'task') return [];
    return drafts.map((mini, index) => ({
      id: `${taskId || 'new'}-mini-${index + 1}`,
      enabled: Boolean(mini.enabled),
      name: typeof mini.name === 'string' ? mini.name : '',
      duration: Math.max(1, Math.round(mini.duration || 1)),
      locked: Boolean(mini.locked && mini.enabled),
      delayDescription: (mini.delayDescription || '').trim(),
      chargeToClient: Boolean(mini.chargeToClient)
    }));
  }

  function computeDefaultMiniTaskDurations(totalDuration) {
    const normalizedTotal = Math.max(
      defaultSequenceMinimumTotal,
      Math.max(1, Math.round(Number(totalDuration) || 0))
    );
    const stepData = defaultSequenceSteps.map((step, index) => ({
      index,
      weight: Number(step.weight) || 0,
      duration: Math.max(1, Math.round(((Number(step.weight) || 0) / 100) * normalizedTotal))
    }));
    const weightOrder = stepData.slice().sort((a, b) => b.weight - a.weight);
    let diff = normalizedTotal - stepData.reduce((sum, item) => sum + item.duration, 0);
    if (diff !== 0) {
      const maxGuard = 100;
      let guard = 0;
      while (diff !== 0 && guard < maxGuard) {
        let adjusted = false;
        for (let idx = 0; idx < weightOrder.length && diff !== 0; idx += 1) {
          const entry = weightOrder[idx];
          if (diff > 0) {
            entry.duration += 1;
            diff -= 1;
            adjusted = true;
          } else if (entry.duration > 1) {
            entry.duration -= 1;
            diff += 1;
            adjusted = true;
          }
        }
        if (!adjusted) break;
        guard += 1;
      }
    }
    return stepData
      .slice()
      .sort((a, b) => a.index - b.index)
      .map((item) => Math.max(1, Math.round(item.duration)));
  }

  function createDefaultSequencedMiniTasks(taskId, totalDuration) {
    const durations = computeDefaultMiniTaskDurations(totalDuration);
    return defaultSequenceSteps.map((step, index) => ({
      id: `${taskId}-mini-${index + 1}`,
      enabled: true,
      name: step.name,
      duration: durations[index] || 1,
      locked: false
    }));
  }

  function sanitizeMiniTaskDraft(draft) {
    const result = {
      enabled: Boolean(draft?.enabled),
      name: typeof draft?.name === 'string' ? draft.name : '',
      duration: Math.max(1, Math.round(Number.isFinite(draft?.duration) ? draft.duration : 1)),
      locked: Boolean(draft?.locked),
      delayDescription: typeof draft?.delayDescription === 'string' ? draft.delayDescription : '',
      chargeToClient: Boolean(draft?.chargeToClient)
    };
    if (result.name !== 'Delays') {
      result.delayDescription = '';
      result.chargeToClient = false;
    }
    return result;
  }

  function createMiniTaskDraft(overrides = {}) {
    return sanitizeMiniTaskDraft({
      enabled: false,
      name: '',
      duration: 1,
      locked: false,
      delayDescription: '',
      chargeToClient: false,
      ...overrides
    });
  }

  return {
    renderMiniTaskInputs,
    updateMiniTaskToolbar,
    setMiniTaskUnit,
    addMiniTaskRow,
    removeMiniTaskRow,
    collectMiniTasks,
    computeDefaultMiniTaskDurations,
    createDefaultSequencedMiniTasks,
    sanitizeMiniTaskDraft,
    createMiniTaskDraft,
    setMiniTaskDrafts,
    getMiniTaskDrafts,
    getMiniTaskUnit,
    resetMiniTaskState,
    getMiniTaskDraftCount: () => drafts.length,
    handleEntryTypeChange
  };
}
