export function createGanttController({
  elements,
  constants,
  accessors,
  helpers
}) {
  const {
    timelineHeaderEl,
    ganttBodyEl,
    ganttBodyScrollEl,
    plannerViewportEl,
    taskListHeaderEl,
    taskListScrollEl,
    taskListEl,
    topBarEl,
    viewModeButtons = [],
    zoomInBtn,
    zoomOutBtn
  } = elements;

  const {
    DEFAULT_DAY_WIDTH,
    MIN_ZOOM_SCALE,
    MAX_ZOOM_SCALE,
    ZOOM_TOLERANCE,
    WEEKS_HEADER_HEIGHT,
    DAYS_HEADER_HEIGHT
  } = constants;

  const {
    getTasks,
    getSelectedTaskId,
    setSelectedTaskId
  } = accessors;

  const {
    renderAll,
    selectTask,
    computeTimelineRange,
    buildTimelineDays,
    getTodayUtc,
    isSameUtcDay,
    formatDate,
    parseDate,
    compareDates,
    buildMilestoneSegments,
    groupTimelineDaysByYear,
    groupTimelineDaysByMonth,
    groupTimelineDaysByWeek,
    buildMiniSegments,
    buildMiniSegmentsPreview,
    minimumTaskDuration,
    ensureWeekday,
    shiftWeekdays,
    diffInWeekdays,
    sanitiseUndefinedDuration,
    computeMiniDurationSum,
    isStanddownDay,
    formatDisplayDate,
    formatShortDate,
    reconcileMiniDurationsForTask,
    getOrderedTasks
  } = helpers;

  const state = {
    dayWidthScale: 1,
    timelineViewMode: 'days',
    timelineDays: [],
    timelineRangeStart: null,
    timelineRangeEnd: null,
    timelineDayIndexMap: new Map(),
    currentStanddownSegments: [],
    isSyncingScroll: false,
    timelineHorizontalSyncFrame: null
  };

  function getDayWidth() {
    const baseWidth = state.timelineViewMode === 'weeks'
      ? Math.max(1, DEFAULT_DAY_WIDTH / 5)
      : DEFAULT_DAY_WIDTH;
    return Math.max(1, Math.round(baseWidth * state.dayWidthScale));
  }

  function updateHeaderOffset() {
    const root = document.documentElement;
    const headerHeight = state.timelineViewMode === 'weeks'
      ? WEEKS_HEADER_HEIGHT
      : DAYS_HEADER_HEIGHT;
    root.style.setProperty('--timeline-header-height', headerHeight);
    root.style.setProperty('--task-header-height', headerHeight);
    if (taskListHeaderEl) {
      taskListHeaderEl.style.height = headerHeight;
      taskListHeaderEl.style.minHeight = headerHeight;
    }
    if (timelineHeaderEl) {
      timelineHeaderEl.style.height = headerHeight;
      timelineHeaderEl.style.minHeight = headerHeight;
    }
  }

  function updateViewModeButtons() {
    viewModeButtons.forEach((button) => {
      const isActive = button.dataset.viewMode === state.timelineViewMode;
      button.classList.toggle('active', isActive);
      button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
    });
  }

  function updateZoomControls() {
    if (!zoomInBtn || !zoomOutBtn) return;
    zoomOutBtn.disabled = state.dayWidthScale <= MIN_ZOOM_SCALE + ZOOM_TOLERANCE;
    zoomInBtn.disabled = state.dayWidthScale >= MAX_ZOOM_SCALE - ZOOM_TOLERANCE;
  }

  function clampZoom(value) {
    return Math.min(MAX_ZOOM_SCALE, Math.max(MIN_ZOOM_SCALE, value));
  }

  function adjustZoom(delta) {
    const nextScale = clampZoom(state.dayWidthScale + delta);
    if (Math.abs(nextScale - state.dayWidthScale) <= ZOOM_TOLERANCE) return;
    state.dayWidthScale = nextScale;
    renderGantt();
    updateZoomControls();
  }

  function setTimelineViewMode(mode) {
    if (mode !== 'days' && mode !== 'weeks') return;
    if (state.timelineViewMode === mode) {
      updateHeaderOffset();
      updateStickyMetrics();
      return;
    }
    state.timelineViewMode = mode;
    updateHeaderOffset();
    updateViewModeButtons();
    renderGantt();
    updateZoomControls();
    updateStickyMetrics();
  }

  function createMiniSegmentElement() {
    const element = document.createElement('div');
    element.className = 'mini-segment mini-control';
    const handleLeft = document.createElement('div');
    handleLeft.className = 'mini-handle mini-handle-left mini-control';
    handleLeft.dataset.edge = 'left';
    handleLeft.setAttribute('role', 'button');
    handleLeft.setAttribute('tabindex', '-1');
    const handleRight = document.createElement('div');
    handleRight.className = 'mini-handle mini-handle-right mini-control';
    handleRight.dataset.edge = 'right';
    handleRight.setAttribute('role', 'button');
    handleRight.setAttribute('tabindex', '-1');
    element.appendChild(handleLeft);
    element.appendChild(handleRight);
    return element;
  }

  function toggleMiniSegmentLock(task, segmentId) {
    if (!task || !segmentId) return;
    const target = task.miniTasks.find((mini) => mini.id === segmentId);
    if (!target || !target.enabled) return;
    target.locked = !target.locked;
    setSelectedTaskId(task.id);
    renderAll();
  }

  function bindMiniSegmentEvents(element, segment, task, barElement, segments, index) {
    if (element._miniCleanup) {
      element._miniCleanup();
      element._miniCleanup = null;
    }
    if (!task || !segment || segment.isUndefined) {
      element.removeAttribute('aria-label');
      element.removeAttribute('aria-pressed');
      element.removeAttribute('role');
      element.removeAttribute('tabindex');
      element.querySelectorAll('.mini-handle').forEach((handle) => {
        handle.classList.add('disabled');
      });
      return;
    }
    const cleanups = [];
    const accessibleLabel = segment.locked ? 'Unlock mini-task length' : 'Lock mini-task length';
    element.setAttribute('role', 'button');
    element.setAttribute('tabindex', '0');
    element.setAttribute('aria-label', accessibleLabel);
    element.setAttribute('aria-pressed', segment.locked ? 'true' : 'false');
    const onClick = (event) => {
      if (event.target.closest('.mini-handle')) return;
      event.preventDefault();
      event.stopPropagation();
      toggleMiniSegmentLock(task, segment.id);
    };
    element.addEventListener('click', onClick);
    cleanups.push(() => element.removeEventListener('click', onClick));
    const onKeyDown = (event) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      event.stopPropagation();
      toggleMiniSegmentLock(task, segment.id);
    };
    element.addEventListener('keydown', onKeyDown);
    cleanups.push(() => element.removeEventListener('keydown', onKeyDown));
    const previousEnabled = (() => {
      if (index <= 0) return null;
      for (let i = index - 1; i >= 0; i -= 1) {
        const candidate = segments[i];
        if (candidate && !candidate.isUndefined) return candidate;
      }
      return null;
    })();
    element.querySelectorAll('.mini-handle').forEach((handle) => {
      const edge = handle.dataset.edge || 'right';
      const needsNeighbor = edge === 'left';
      if (segment.locked || (needsNeighbor && !previousEnabled)) {
        handle.classList.add('disabled');
        return;
      }
      handle.classList.remove('disabled');
      const onPointerDown = (event) => {
        beginMiniSegmentResize(event, task, segment.id, barElement, edge, previousEnabled?.id || null);
      };
      handle.addEventListener('pointerdown', onPointerDown);
      cleanups.push(() => handle.removeEventListener('pointerdown', onPointerDown));
    });
    element._miniCleanup = () => {
      cleanups.forEach((fn) => fn());
    };
  }

  function updateMiniLayerElements(barElement, segments, options = {}) {
    const miniLayer = barElement.querySelector('.mini-layer');
    if (!miniLayer) return;
    const interactive = options.interactive !== false;
    const task = options.task || null;
    const dayWidth = getDayWidth();
    const startIndex = Number(barElement.dataset.startIndex);
    const endIndex = Number(barElement.dataset.endIndex);
    if (!Number.isFinite(startIndex) || !Number.isFinite(endIndex)) return;
    const workingDayIndices = [];
    for (let index = startIndex; index <= endIndex; index += 1) {
      const day = state.timelineDays[index];
      if (!isStanddownDay(day)) workingDayIndices.push(index);
    }

    while (miniLayer.children.length < segments.length) {
      miniLayer.appendChild(createMiniSegmentElement());
    }

    if (!interactive) {
      miniLayer.querySelectorAll('.mini-segment').forEach((element) => {
        if (element._miniCleanup) {
          element._miniCleanup();
          element._miniCleanup = null;
        }
      });
    }

    const miniElements = Array.from(miniLayer.querySelectorAll('.mini-segment'));
    let workingPointer = 0;
    segments.forEach((segment, index) => {
      const element = miniElements[index];
      if (!element) return;
      element.style.display = '';
      element.dataset.segmentId = segment.id || '';
      element.dataset.boundSegmentId = segment.id || '';
      const requiredSpan = Math.max(1, Math.round(segment.duration || 1));
      const startWorkingIndex = workingDayIndices[workingPointer];
      const segmentEndWorkingIndex = workingDayIndices[Math.min(
        workingPointer + Math.max(0, requiredSpan - 1),
        workingDayIndices.length - 1
      )];
      if (typeof startWorkingIndex !== 'number' || typeof segmentEndWorkingIndex !== 'number') {
        element.style.display = 'none';
        return;
      }
      element.style.left = `${(startWorkingIndex - startIndex) * dayWidth}px`;
      element.style.width = `${(segmentEndWorkingIndex - startWorkingIndex + 1) * dayWidth}px`;
      element.classList.toggle('is-undefined', Boolean(segment.isUndefined));
      element.querySelectorAll('.mini-handle').forEach((handleEl) => {
        const isDisabled = Boolean(segment.locked || segment.isUndefined);
        handleEl.classList.toggle('disabled', isDisabled);
      });
      if (interactive && task && !segment.isUndefined) {
        bindMiniSegmentEvents(element, segment, task, barElement, segments, index);
        element.dataset.boundSegmentId = segment.id || '';
      }
      workingPointer += requiredSpan;
    });

    for (let i = segments.length; i < miniElements.length; i += 1) {
      const element = miniElements[i];
      if (!element) continue;
      if (interactive && element._miniCleanup) {
        element._miniCleanup();
        element._miniCleanup = null;
      }
      element.style.display = 'none';
      element.dataset.segmentId = '';
      element.dataset.boundSegmentId = '';
    }
  }

  function findNearestDonor(enabledSegments, targetIndex) {
    for (let offset = 1; offset < enabledSegments.length; offset += 1) {
      const rightIndex = targetIndex + offset;
      if (rightIndex < enabledSegments.length) {
        const candidate = enabledSegments[rightIndex];
        const current = Math.max(1, Math.round(candidate.duration || 1));
        if (!candidate.locked && current > 1) return candidate;
      }
      const leftIndex = targetIndex - offset;
      if (leftIndex >= 0) {
        const candidate = enabledSegments[leftIndex];
        const current = Math.max(1, Math.round(candidate.duration || 1));
        if (!candidate.locked && current > 1) return candidate;
      }
    }
    return null;
  }

  function findNearestRecipient(enabledSegments, targetIndex) {
    for (let offset = 1; offset < enabledSegments.length; offset += 1) {
      const rightIndex = targetIndex + offset;
      if (rightIndex < enabledSegments.length) {
        const candidate = enabledSegments[rightIndex];
        if (!candidate.locked) return candidate;
      }
      const leftIndex = targetIndex - offset;
      if (leftIndex >= 0) {
        const candidate = enabledSegments[leftIndex];
        if (!candidate.locked) return candidate;
      }
    }
    return null;
  }

  function adjustMiniDurationStep(stateSnapshot, segmentId, direction) {
    const enabled = stateSnapshot.miniTasks.filter((mini) => mini.enabled);
    const targetIndex = enabled.findIndex((mini) => mini.id === segmentId);
    if (targetIndex === -1) return false;
    const target = enabled[targetIndex];
    if (!target || target.locked) return false;
    const targetDuration = Math.max(1, Math.round(target.duration || 1));
    if (direction > 0) {
      const donor = findNearestDonor(enabled, targetIndex);
      if (donor) {
        const donorDuration = Math.max(1, Math.round(donor.duration || 1));
        if (donorDuration <= 1) return false;
        donor.duration = donorDuration - 1;
        target.duration = targetDuration + 1;
        return true;
      }
      target.duration = targetDuration + 1;
      stateSnapshot.undefinedDuration = sanitiseUndefinedDuration(stateSnapshot.undefinedDuration + 1);
      return true;
    }
    if (direction < 0) {
      if (targetDuration <= 1) return false;
      target.duration = targetDuration - 1;
      const recipient = findNearestRecipient(enabled, targetIndex);
      if (recipient) {
        recipient.duration = Math.max(1, Math.round(recipient.duration || 1)) + 1;
        return true;
      }
      if (stateSnapshot.undefinedDuration > 0) {
        stateSnapshot.undefinedDuration = sanitiseUndefinedDuration(stateSnapshot.undefinedDuration - 1);
        return true;
      }
      return false;
    }
    return false;
  }

  function beginMiniSegmentResize(event, task, segmentId, barElement, edge = 'right', neighborId = null) {
    if (!task || !segmentId) return;
    const handleEdge = edge === 'left' ? 'left' : 'right';
    const segmentOrder = task.miniTasks.filter((mini) => mini.enabled);
    const targetIndex = segmentOrder.findIndex((mini) => mini.id === segmentId);
    if (targetIndex === -1) return;
    const segment = segmentOrder[targetIndex];
    if (!segment || segment.locked) return;
    let activeSegmentId = segmentId;
    if (handleEdge === 'left') {
      const previous = neighborId
        ? segmentOrder.find((mini) => mini.id === neighborId)
        : segmentOrder[targetIndex - 1];
      if (!previous || previous.locked) return;
      activeSegmentId = previous.id;
    }
    event.preventDefault();
    event.stopPropagation();
    const handle = event.currentTarget;
    if (!handle || typeof handle.setPointerCapture !== 'function') return;
    const pointerId = event.pointerId;
    handle.setPointerCapture(pointerId);
    const snapshot = {
      miniTasks: task.miniTasks.map((mini) => ({ ...mini })),
      undefinedDuration: sanitiseUndefinedDuration(task.undefinedDuration || 0)
    };
    computeMiniDurationSum(snapshot.miniTasks);
    const dayWidth = getDayWidth();
    const initialPointerX = event.clientX;
    const initialStartDate = parseDate(task.startDate);
    if (!initialStartDate) {
      handle.releasePointerCapture(pointerId);
      return;
    }
    const initialStartIndex = Number(barElement.dataset.startIndex);
    let appliedDelta = 0;

    const applyPreview = () => {
      const totalDuration = Math.max(1, computeMiniDurationSum(snapshot.miniTasks) + sanitiseUndefinedDuration(snapshot.undefinedDuration));
      const previewEndDate = shiftWeekdays(initialStartDate, totalDuration - 1);
      const previewTask = {
        id: task.id,
        miniTasks: snapshot.miniTasks.map((mini) => ({ ...mini })),
        undefinedDuration: sanitiseUndefinedDuration(snapshot.undefinedDuration),
        startDate: formatDate(initialStartDate),
        endDate: formatDate(previewEndDate)
      };
      const previewSegments = buildMiniSegments(previewTask);
      updateMiniLayerElements(barElement, previewSegments, { interactive: false });
      const previewEndIndex = state.timelineDayIndexMap.get(formatDate(previewEndDate));
      if (typeof previewEndIndex === 'number') {
        barElement.dataset.endIndex = String(previewEndIndex);
        const width = Math.max(dayWidth, (previewEndIndex - initialStartIndex + 1) * dayWidth);
        barElement.style.width = `${width}px`;
      } else {
        const fallbackWidth = Math.max(dayWidth, totalDuration * dayWidth);
        barElement.style.width = `${fallbackWidth}px`;
      }
    };

    const applyDelta = (desiredDelta) => {
      let diff = desiredDelta - appliedDelta;
      if (!diff) return;
      while (diff !== 0) {
        const step = diff > 0 ? 1 : -1;
        const success = adjustMiniDurationStep(snapshot, activeSegmentId, step);
        if (!success) break;
        appliedDelta += step;
        diff -= step;
      }
      applyPreview();
    };

    const onPointerMove = (moveEvent) => {
      const deltaPx = moveEvent.clientX - initialPointerX;
      const desiredDelta = Math.round(deltaPx / dayWidth);
      applyDelta(desiredDelta);
    };

    const cleanupPointer = () => {
      handle.removeEventListener('pointermove', onPointerMove);
      handle.removeEventListener('pointerup', onPointerUp);
      handle.removeEventListener('pointercancel', onPointerUp);
      if (handle.hasPointerCapture(pointerId)) {
        handle.releasePointerCapture(pointerId);
      }
    };

    const onPointerUp = () => {
      cleanupPointer();
      if (appliedDelta === 0) {
        setSelectedTaskId(task.id);
        renderAll();
        return;
      }
      computeMiniDurationSum(snapshot.miniTasks);
      task.miniTasks.forEach((mini, index) => {
        if (!snapshot.miniTasks[index]) return;
        mini.duration = snapshot.miniTasks[index].duration;
      });
      task.undefinedDuration = sanitiseUndefinedDuration(snapshot.undefinedDuration);
      const finalTotal = Math.max(1, computeMiniDurationSum(task.miniTasks) + task.undefinedDuration);
      const finalEndDate = shiftWeekdays(initialStartDate, finalTotal - 1);
      task.startDate = formatDate(initialStartDate);
      task.endDate = formatDate(finalEndDate);
      setSelectedTaskId(task.id);
      renderAll();
    };

    handle.addEventListener('pointermove', onPointerMove);
    handle.addEventListener('pointerup', onPointerUp);
    handle.addEventListener('pointercancel', onPointerUp);
    applyPreview();
  }

  function refreshStanddownLayer(barElement) {
    const layer = barElement.querySelector('.standdown-layer');
    if (!layer) return;
    layer.innerHTML = '';
    const startIndex = Number(barElement.dataset.startIndex);
    const endIndex = Number(barElement.dataset.endIndex);
    if (!Number.isFinite(startIndex) || !Number.isFinite(endIndex)) return;
    const dayWidth = getDayWidth();
    state.currentStanddownSegments.forEach((segment) => {
      const overlapStart = Math.max(startIndex, segment.startIndex);
      const overlapEnd = Math.min(endIndex, segment.endIndex);
      if (overlapStart > overlapEnd) return;
      const mask = document.createElement('div');
      mask.className = 'standdown-gap';
      mask.style.left = `${(overlapStart - startIndex) * dayWidth}px`;
      mask.style.width = `${(overlapEnd - overlapStart + 1) * dayWidth}px`;
      if (overlapStart === segment.startIndex) mask.classList.add('standdown-gap-start');
      if (overlapEnd === segment.endIndex) mask.classList.add('standdown-gap-end');
      const rangeStart = state.timelineDays[overlapStart];
      const rangeEnd = state.timelineDays[overlapEnd];
      if (rangeStart) {
        const labelPrefix = segment.name ? `${segment.name} • ` : '';
        const dateLabel = overlapStart === overlapEnd
          ? `${formatDisplayDate(rangeStart)}`
          : `${formatDisplayDate(rangeStart)} → ${formatDisplayDate(rangeEnd)}`;
        mask.title = `${labelPrefix}Stand-down ${dateLabel}`;
      }
      layer.appendChild(mask);
    });
  }

  function setBarLabel(barElement, task, startDate, endDate) {
    const label = barElement.querySelector('.task-label');
    const datesEl = barElement.querySelector('.task-dates') || document.createElement('span');
    const duration = Math.max(1, diffInWeekdays(startDate, endDate));
    barElement.title = `${task.name}\n${formatDisplayDate(startDate)} → ${formatDisplayDate(endDate)} (${duration} days)`;
    datesEl.className = 'task-dates';
    datesEl.textContent = `${formatDisplayDate(startDate)} → ${formatDisplayDate(endDate)}`;
    if (!datesEl.isConnected) {
      const handle = barElement.querySelector('.handle');
      if (handle) {
        barElement.insertBefore(datesEl, handle);
      } else {
        barElement.appendChild(datesEl);
      }
    }
    if (label) label.textContent = task.name;
  }

  function attachResizeHandles(barElement, task) {
    barElement.querySelectorAll('.handle').forEach((handle) => {
      const type = handle.classList.contains('handle-start') ? 'start' : 'end';
      handle.addEventListener('pointerdown', (event) => {
        event.stopPropagation();
        event.preventDefault();
        handle.setPointerCapture(event.pointerId);
        const initialStart = parseDate(task.startDate);
        const initialEnd = parseDate(task.endDate);
        let previewStart = new Date(initialStart.getTime());
        let previewEnd = new Date(initialEnd.getTime());
        const progressEl = barElement.querySelector('.progress');

        const onPointerMove = (moveEvent) => {
          const calendarEl = plannerViewportEl || document.scrollingElement || document.documentElement;
          const headerRect = timelineHeaderEl.getBoundingClientRect();
          const relativeX = moveEvent.clientX - headerRect.left + calendarEl.scrollLeft;
          const dayWidth = getDayWidth();
          let dayIndex = Math.round(relativeX / dayWidth);
          dayIndex = Math.max(0, Math.min(dayIndex, state.timelineDays.length - 1));
          const targetDate = state.timelineDays[dayIndex];
          if (!targetDate) return;

          if (type === 'start') {
            let candidate = ensureWeekday(targetDate, 1);
            if (compareDates(candidate, previewEnd) > 0) {
              candidate = shiftWeekdays(previewEnd, -(minimumTaskDuration(task) - 1));
            }
            previewStart = candidate;
          } else {
            let candidate = ensureWeekday(targetDate, -1);
            if (compareDates(candidate, previewStart) < 0) {
              candidate = shiftWeekdays(previewStart, minimumTaskDuration(task) - 1);
            }
            previewEnd = candidate;
          }

          const duration = Math.max(minimumTaskDuration(task), diffInWeekdays(previewStart, previewEnd));
          const dayWidthForRender = getDayWidth();
          const previewStartIndex = state.timelineDayIndexMap.get(formatDate(previewStart));
          const previewEndIndex = state.timelineDayIndexMap.get(formatDate(previewEnd));
          if (typeof previewStartIndex !== 'number' || typeof previewEndIndex !== 'number') return;
          barElement.style.left = `${previewStartIndex * dayWidthForRender}px`;
          barElement.style.width = `${Math.max(dayWidthForRender, (previewEndIndex - previewStartIndex + 1) * dayWidthForRender)}px`;
          barElement.dataset.startIndex = String(previewStartIndex);
          barElement.dataset.endIndex = String(previewEndIndex);
          const previewSegments = buildMiniSegmentsPreview(task, duration);
          updateMiniLayerElements(barElement, previewSegments, { interactive: false });
          refreshStanddownLayer(barElement);
          if (progressEl) {
            const progressValue = Math.min(100, Math.max(0, task.progress));
            progressEl.style.width = `${progressValue}%`;
            progressEl.classList.toggle('is-full', progressValue >= 99.5);
          }
          setBarLabel(barElement, task, previewStart, previewEnd);
        };

        const onPointerUp = () => {
          handle.removeEventListener('pointermove', onPointerMove);
          handle.removeEventListener('pointerup', onPointerUp);
          handle.removeEventListener('pointercancel', onPointerUp);
          handle.releasePointerCapture(event.pointerId);
          if (type === 'start') {
            applyNewStartDate(task, previewStart, initialEnd);
          } else {
            applyNewEndDate(task, initialStart, previewEnd);
          }
          renderAll();
          selectTask(task.id);
        };

        handle.addEventListener('pointermove', onPointerMove);
        handle.addEventListener('pointerup', onPointerUp);
        handle.addEventListener('pointercancel', onPointerUp);
      });
    });
  }

  function attachProgressHandle(barElement, task) {
    if (!barElement || !task || task.entryType === 'stage') return;
    const progressEl = barElement.querySelector('.progress');
    if (!progressEl) return;
    let handle = barElement.querySelector('.progress-handle');
    if (!handle) {
      handle = document.createElement('div');
      handle.className = 'progress-handle mini-control';
      handle.setAttribute('role', 'slider');
      handle.setAttribute('aria-valuemin', '0');
      handle.setAttribute('aria-valuemax', '100');
      handle.setAttribute('aria-orientation', 'horizontal');
      handle.setAttribute('tabindex', '-1');
      barElement.appendChild(handle);
    }

    const updateAria = (value) => {
      const taskName = ((task.name || '').trim()) || 'task';
      handle.setAttribute('aria-valuenow', String(value));
      handle.setAttribute('aria-valuetext', `${value}%`);
      handle.setAttribute('aria-label', `Adjust progress for ${taskName}`);
    };

    const updateFromClientX = (clientX) => {
      const barRect = barElement.getBoundingClientRect();
      const width = Math.max(1, barRect.width);
      let offset = clientX - barRect.left;
      offset = Math.max(0, Math.min(offset, width));
      const progressValue = Math.round((offset / width) * 100);
      task.progress = progressValue;
      progressEl.style.width = `${progressValue}%`;
      progressEl.classList.toggle('is-full', progressValue >= 99.5);
      handle.style.left = `${progressValue}%`;
      updateAria(progressValue);
      const listRow = taskListEl ? taskListEl.querySelector(`.task-row[data-task-id="${task.id}"]`) : null;
      if (listRow) {
        const input = listRow.querySelector('.task-progress input');
        if (input) {
          input.value = progressValue;
        } else {
          const display = listRow.querySelector('.task-progress');
          if (display) display.textContent = `${progressValue}%`;
        }
      }
    };

    if (!handle._progressHandleReady) {
      handle.addEventListener('pointerdown', (event) => {
        event.stopPropagation();
        event.preventDefault();
        setSelectedTaskId(task.id);
        const pointerId = event.pointerId;
        handle.setPointerCapture(pointerId);
        updateFromClientX(event.clientX);

        const onPointerMove = (moveEvent) => {
          updateFromClientX(moveEvent.clientX);
        };

        const onPointerUp = () => {
          if (handle.hasPointerCapture(pointerId)) handle.releasePointerCapture(pointerId);
          handle.removeEventListener('pointermove', onPointerMove);
          handle.removeEventListener('pointerup', onPointerUp);
          handle.removeEventListener('pointercancel', onPointerUp);
          renderAll();
        };

        handle.addEventListener('pointermove', onPointerMove);
        handle.addEventListener('pointerup', onPointerUp);
        handle.addEventListener('pointercancel', onPointerUp);
      });
      Object.defineProperty(handle, '_progressHandleReady', {
        value: true,
        enumerable: false,
        configurable: true,
        writable: false
      });
    }
    const currentProgress = Math.min(100, Math.max(0, Number(task.progress) || 0));
    updateAria(currentProgress);
    handle.style.left = `${currentProgress}%`;
  }

  function enableBarDrag(barElement, task) {
    barElement.addEventListener('pointerdown', (event) => {
      if (event.button && event.button !== 0) return;
      if (event.target.closest('.handle') || event.target.closest('.mini-handle')) return;
      if (event.target.closest('button, input, select, textarea, label')) return;
      event.stopPropagation();
      const initialStart = parseDate(task.startDate);
      const initialEnd = parseDate(task.endDate);
      if (!initialStart || !initialEnd) return;
      const calendarEl = plannerViewportEl || document.scrollingElement || document.documentElement;
      if (!calendarEl) return;
      const initialPointer = event.clientX + calendarEl.scrollLeft;
      const initialLeft = barElement.style.left;
      let latestShift = 0;
      let dragActive = false;

      const activateDrag = () => {
        if (dragActive) return;
        dragActive = true;
        barElement.setPointerCapture(event.pointerId);
      };

      const cleanup = () => {
        barElement.removeEventListener('pointermove', onPointerMove);
        barElement.removeEventListener('pointerup', onPointerUp);
        barElement.removeEventListener('pointercancel', onPointerUp);
        if (dragActive && barElement.hasPointerCapture(event.pointerId)) {
          barElement.releasePointerCapture(event.pointerId);
        }
      };

      const onPointerMove = (moveEvent) => {
        const currentPointer = moveEvent.clientX + calendarEl.scrollLeft;
        const deltaPx = currentPointer - initialPointer;
        const dayWidth = getDayWidth();
        if (!dragActive && Math.abs(deltaPx) >= Math.max(6, dayWidth * 0.35)) {
          activateDrag();
        }
        if (!dragActive) return;
        moveEvent.preventDefault();
        const shift = Math.round(deltaPx / dayWidth);
        if (shift === latestShift) return;
        latestShift = shift;
        const previewStart = shiftWeekdays(initialStart, shift);
        const previewEnd = shiftWeekdays(initialEnd, shift);
        const previewStartIndex = state.timelineDayIndexMap.get(formatDate(previewStart));
        const previewEndIndex = state.timelineDayIndexMap.get(formatDate(previewEnd));
        if (typeof previewStartIndex !== 'number' || typeof previewEndIndex !== 'number') return;
        barElement.style.left = `${previewStartIndex * dayWidth}px`;
        barElement.dataset.startIndex = String(previewStartIndex);
        barElement.dataset.endIndex = String(previewEndIndex);
        refreshStanddownLayer(barElement);
        setBarLabel(barElement, task, previewStart, previewEnd);
      };

      const onPointerUp = () => {
        if (dragActive) {
          if (latestShift !== 0) {
            const newStart = shiftWeekdays(initialStart, latestShift);
            const newEnd = shiftWeekdays(initialEnd, latestShift);
            task.startDate = formatDate(newStart);
            task.endDate = formatDate(newEnd);
            renderAll();
            selectTask(task.id);
          } else {
            barElement.style.left = initialLeft;
            const originalStartIndex = state.timelineDayIndexMap.get(task.startDate);
            const originalEndIndex = state.timelineDayIndexMap.get(task.endDate);
            if (typeof originalStartIndex === 'number') {
              barElement.dataset.startIndex = String(originalStartIndex);
            } else {
              delete barElement.dataset.startIndex;
            }
            if (typeof originalEndIndex === 'number') {
              barElement.dataset.endIndex = String(originalEndIndex);
            } else {
              delete barElement.dataset.endIndex;
            }
            refreshStanddownLayer(barElement);
            setBarLabel(barElement, task, initialStart, initialEnd);
          }
        }
        cleanup();
      };

      barElement.addEventListener('pointermove', onPointerMove);
      barElement.addEventListener('pointerup', onPointerUp);
      barElement.addEventListener('pointercancel', onPointerUp);
    });
  }

  function renderGantt() {
    if (!ganttBodyEl || !timelineHeaderEl) return;
    const previousScrollTop = ganttBodyScrollEl ? ganttBodyScrollEl.scrollTop : 0;
    ganttBodyEl.innerHTML = '';
    timelineHeaderEl.innerHTML = '';
    timelineHeaderEl.classList.remove('week-view', 'day-view');
    const { start, end } = computeTimelineRange();
    state.timelineRangeStart = start;
    state.timelineRangeEnd = end;
    state.timelineDays = buildTimelineDays(start, end);
    const todayUtc = getTodayUtc();
    const todayIndex = todayUtc ? state.timelineDays.findIndex((day) => isSameUtcDay(day, todayUtc)) : -1;
    state.timelineDayIndexMap = new Map();
    state.timelineDays.forEach((day, index) => {
      state.timelineDayIndexMap.set(formatDate(day), index);
    });
    const getTimelineIndex = (value) => {
      if (typeof value === 'number' && Number.isFinite(value)) return value;
      const key = typeof value === 'string' ? value : formatDate(value instanceof Date ? value : parseDate(value));
      if (typeof key === 'string' && state.timelineDayIndexMap.has(key)) {
        const stored = state.timelineDayIndexMap.get(key);
        if (typeof stored === 'number') return stored;
      }
      const date = typeof value === 'string' ? parseDate(value) : (value instanceof Date ? value : null);
      if (!(date instanceof Date)) return -1;
      for (let i = 0; i < state.timelineDays.length; i += 1) {
        if (compareDates(state.timelineDays[i], date) === 0) return i;
      }
      return -1;
    };
    const dayHighlightMap = state.timelineDays.map(() => new Set());
    const highlightSegments = [];
    if (todayIndex !== -1) {
      highlightSegments.push({
        id: 'today',
        type: 'today',
        startIndex: todayIndex,
        endIndex: todayIndex
      });
    }
    const milestoneSegments = buildMilestoneSegments(state.timelineDays);
    milestoneSegments.forEach((segment) => {
      highlightSegments.push(segment);
    });
    const normalizedSegments = highlightSegments
      .map((segment) => {
        const startIndexSafe = Math.max(0, Math.min(state.timelineDays.length - 1, segment.startIndex));
        const endIndexSafe = Math.max(startIndexSafe, Math.min(state.timelineDays.length - 1, segment.endIndex));
        return { ...segment, startIndex: startIndexSafe, endIndex: endIndexSafe };
      })
      .filter((segment) => segment.startIndex <= segment.endIndex);
    state.currentStanddownSegments = normalizedSegments.filter((segment) => segment.type === 'standdown');
    normalizedSegments.forEach((segment) => {
      for (let index = segment.startIndex; index <= segment.endIndex; index += 1) {
        dayHighlightMap[index]?.add(segment.type);
      }
      if (segment.type === 'standdown') {
        dayHighlightMap[segment.startIndex]?.add('standdown-start');
        dayHighlightMap[segment.endIndex]?.add('standdown-end');
      }
    });
    const applyHeaderHighlightClasses = (element, highlightSet) => {
      if (!element || !(highlightSet instanceof Set)) return;
      if (highlightSet.has('today')) element.classList.add('today');
      if (highlightSet.has('deadline')) element.classList.add('milestone-deadline');
      if (highlightSet.has('standdown')) element.classList.add('milestone-standdown');
      if (highlightSet.has('standdown-start')) element.classList.add('milestone-standdown-start');
      if (highlightSet.has('standdown-end')) element.classList.add('milestone-standdown-end');
    };
    const combineHighlightSet = (startIndex, endIndex) => {
      const combined = new Set();
      for (let idx = startIndex; idx <= endIndex; idx += 1) {
        const highlightSet = dayHighlightMap[idx];
        if (highlightSet instanceof Set) {
          highlightSet.forEach((type) => combined.add(type));
        }
      }
      return combined;
    };
    const getWeekStartLabelDate = (date) => {
      if (!(date instanceof Date)) return null;
      const base = new Date(date.getTime());
      const day = base.getUTCDay();
      const offset = day === 0 ? -6 : (1 - day);
      if (offset !== 0) {
        base.setUTCDate(base.getUTCDate() + offset);
      }
      return base;
    };
    if (!state.timelineDays.length) {
      timelineHeaderEl.classList.remove('week-view', 'day-view');
      const emptyHeader = document.createElement('div');
      emptyHeader.className = 'empty-state';
      emptyHeader.textContent = 'No working days to display.';
      timelineHeaderEl.appendChild(emptyHeader);
      updateZoomControls();
      if (ganttBodyScrollEl) ganttBodyScrollEl.scrollTop = 0;
      return;
    }
    const dayWidth = getDayWidth();
    const getLabelOffsetWithinRange = (targetDate, startIndex, endIndex) => {
      if (!Number.isFinite(startIndex) || !Number.isFinite(endIndex) || endIndex < startIndex) return null;
      const totalWidth = (endIndex - startIndex + 1) * dayWidth;
      if (totalWidth <= 0) return null;
      let labelIndex = -1;
      if (targetDate instanceof Date) {
        for (let idx = startIndex; idx <= endIndex; idx += 1) {
          const day = state.timelineDays[idx];
          if (compareDates(day, targetDate) >= 0) {
            labelIndex = idx;
            break;
          }
        }
      }
      if (!Number.isFinite(labelIndex) || labelIndex < startIndex || labelIndex > endIndex) {
        labelIndex = Math.floor((startIndex + endIndex) / 2);
      }
      const relativeIndex = Math.max(0, labelIndex - startIndex);
      const offsetPx = (relativeIndex + 0.5) * dayWidth;
      return Math.max(dayWidth * 0.5, Math.min(offsetPx, totalWidth - (dayWidth * 0.5)));
    };
    const setHeaderCellLabel = (cell, text, offsetPx) => {
      if (!cell) return;
      cell.textContent = '';
      const label = document.createElement('span');
      label.className = 'timeline-header-label';
      label.textContent = text;
      if (Number.isFinite(offsetPx)) {
        label.style.left = `${offsetPx}px`;
      } else {
        label.style.left = '50%';
      }
      cell.appendChild(label);
    };
    timelineHeaderEl.style.gridTemplateColumns = `repeat(${state.timelineDays.length}, ${dayWidth}px)`;
    const yearGroups = groupTimelineDaysByYear(state.timelineDays);
    const monthGroups = groupTimelineDaysByMonth(state.timelineDays);
    if (state.timelineViewMode === 'weeks') {
      timelineHeaderEl.classList.add('week-view');
      timelineHeaderEl.style.gridTemplateRows = 'auto auto auto';
    } else {
      timelineHeaderEl.classList.add('day-view');
      timelineHeaderEl.style.gridTemplateRows = 'auto auto auto auto';
    }
    adjustGanttOffset();

    normalizedSegments.forEach((segment) => {
      const highlightEl = document.createElement('div');
      highlightEl.className = `timeline-highlight timeline-highlight--${segment.type}`;
      if (segment.type === 'standdown') {
        highlightEl.classList.add('timeline-highlight-standdown');
      }
      highlightEl.style.gridColumn = `${segment.startIndex + 1} / ${segment.endIndex + 2}`;
      highlightEl.style.gridRow = '1 / -1';
      if (segment.type === 'today') {
        highlightEl.title = 'Today';
      } else if (segment.name) {
        const label = segment.type === 'deadline' ? 'Deadline' : 'Stand-down';
        highlightEl.title = `${segment.name} • ${label}`;
      }
      timelineHeaderEl.appendChild(highlightEl);
    });

    yearGroups.forEach((group) => {
      const cell = document.createElement('div');
      cell.className = 'timeline-year';
      cell.style.gridColumn = `${group.startIndex + 1} / ${group.endIndex + 2}`;
      cell.style.gridRow = '1';
      const yearMidDate = new Date(Date.UTC(group.year, 5, 30));
      const offsetPx = getLabelOffsetWithinRange(yearMidDate, group.startIndex, group.endIndex);
      setHeaderCellLabel(cell, String(group.year), offsetPx);
      timelineHeaderEl.appendChild(cell);
    });

    monthGroups.forEach((group) => {
      const cell = document.createElement('div');
      cell.className = 'timeline-month';
      cell.style.gridColumn = `${group.startIndex + 1} / ${group.endIndex + 2}`;
      cell.style.gridRow = '2';
      const monthMidDate = new Date(Date.UTC(group.year, group.month, 15));
      const offsetPx = getLabelOffsetWithinRange(monthMidDate, group.startIndex, group.endIndex);
      setHeaderCellLabel(cell, group.name || '', offsetPx);
      timelineHeaderEl.appendChild(cell);
    });

    if (state.timelineViewMode === 'weeks') {
      const weekGroups = groupTimelineDaysByWeek(state.timelineDays);
      weekGroups.forEach((group) => {
        const highlightSet = combineHighlightSet(group.startIndex, group.endIndex);
        const weekDate = document.createElement('div');
        weekDate.className = 'timeline-week-date timeline-day-number is-monday';
        const mondayDate = getWeekStartLabelDate(group.start);
        const labelDate = mondayDate || group.start;
        weekDate.textContent = formatShortDate(labelDate).split('/')[0];
        weekDate.style.gridRow = '3';
        weekDate.style.gridColumn = `${group.startIndex + 1} / ${group.endIndex + 2}`;
        applyHeaderHighlightClasses(weekDate, highlightSet);
        timelineHeaderEl.appendChild(weekDate);
      });
    } else {
      state.timelineDays.forEach((day, index) => {
        const highlightSet = dayHighlightMap[index] || new Set();
        const dayName = document.createElement('div');
        dayName.className = 'timeline-day-name';
        dayName.textContent = day.toLocaleDateString('en-GB', { weekday: 'short' });
        dayName.style.gridRow = '3';
        dayName.style.gridColumn = `${index + 1}`;
        if (day.getUTCDay() === 1) dayName.classList.add('is-monday');
        applyHeaderHighlightClasses(dayName, highlightSet);
        timelineHeaderEl.appendChild(dayName);

        const dayNumber = document.createElement('div');
        dayNumber.className = 'timeline-day-number';
        dayNumber.textContent = formatShortDate(day).split('/')[0];
        dayNumber.style.gridRow = '4';
        dayNumber.style.gridColumn = `${index + 1}`;
        if (day.getUTCDay() === 1) {
          dayNumber.classList.add('is-monday');
        }
        applyHeaderHighlightClasses(dayNumber, highlightSet);
        timelineHeaderEl.appendChild(dayNumber);
      });
    }

    const tasks = getTasks();
    if (!tasks.length) {
      const empty = document.createElement('div');
      empty.className = 'empty-state';
      empty.textContent = 'Create a task to visualise the schedule.';
      ganttBodyEl.appendChild(empty);
      updateZoomControls();
      if (ganttBodyScrollEl) ganttBodyScrollEl.scrollTop = 0;
      return;
    }

    const orderedEntries = getOrderedTasks();
    const selectedTaskId = getSelectedTaskId();
    orderedEntries.forEach(({ task, depth }) => {
      const isStage = task.entryType === 'stage';
      const isStageChild = depth > 0 && !isStage;
      const row = document.createElement('div');
      row.className = 'gantt-row';
      row.style.width = `${state.timelineDays.length * dayWidth}px`;
      row.dataset.taskId = task.id;
      if (isStage) row.classList.add('stage-row');
      if (isStageChild) row.classList.add('stage-child');

      const grid = document.createElement('div');
      grid.className = 'gantt-grid';
      grid.style.gridTemplateColumns = `repeat(${state.timelineDays.length}, ${dayWidth}px)`;
      state.timelineDays.forEach((_, index) => {
        const cell = document.createElement('div');
        cell.className = 'day-cell';
        if (index % 5 === 0) cell.classList.add('week-start');
        const highlightSet = dayHighlightMap[index] || new Set();
        if (highlightSet.has('today')) cell.classList.add('today');
        if (highlightSet.has('deadline')) cell.classList.add('milestone-deadline');
        if (highlightSet.has('standdown')) cell.classList.add('milestone-standdown');
        if (highlightSet.has('standdown-start')) cell.classList.add('milestone-standdown-start');
        if (highlightSet.has('standdown-end')) cell.classList.add('milestone-standdown-end');
        grid.appendChild(cell);
      });

      row.appendChild(grid);

      const startDate = parseDate(task.startDate);
      const startIndex = getTimelineIndex(task.startDate);
      if (isStage && (!startDate || typeof startIndex !== 'number')) {
        row.addEventListener('click', () => selectTask(task.id));
        ganttBodyEl.appendChild(row);
        return;
      }
      if (!isStage && (!startDate || typeof startIndex !== 'number')) {
        return;
      }
      const isSingleEvent = Boolean(task.singleEvent);
      if (isSingleEvent) {
        const marker = document.createElement('div');
        marker.className = 'single-event-marker';
        if (task.id === selectedTaskId) marker.classList.add('selected');
        const markerLeft = (startIndex * dayWidth) + (dayWidth / 2);
        marker.style.left = `${markerLeft}px`;
        marker.dataset.startIndex = String(startIndex);
        marker.dataset.endIndex = String(startIndex);
        marker.title = `${task.name}\n${formatDisplayDate(startDate)} (Single event)`;
        marker.addEventListener('click', (event) => {
          event.stopPropagation();
          selectTask(task.id);
        });
        row.appendChild(marker);
        ganttBodyEl.appendChild(row);
        return;
      }
      const endDate = parseDate(task.endDate);
      const endIndex = getTimelineIndex(task.endDate);
      if (typeof endIndex !== 'number' || startIndex > endIndex) {
        return;
      }
      const barLeft = startIndex * dayWidth;
      const barWidth = Math.max(dayWidth, (endIndex - startIndex + 1) * dayWidth);
      const bar = document.createElement('div');
      bar.className = 'task-bar';
      if (isStage) bar.classList.add('stage-bar');
      if (task.id === selectedTaskId) bar.classList.add('selected');
      bar.style.left = `${barLeft}px`;
      bar.style.width = `${barWidth}px`;
      bar.dataset.startIndex = String(startIndex);
      bar.dataset.endIndex = String(endIndex);

      const label = document.createElement('div');
      label.className = 'task-label';
      label.textContent = task.name || '';
      const progress = document.createElement('div');
      progress.className = 'progress';
      const progressValue = Math.min(100, Math.max(0, task.progress));
      progress.style.width = `${progressValue}%`;
      progress.classList.toggle('is-full', progressValue >= 99.5);

      let miniLayer = null;
      let standdownLayer = null;
      if (!isStage) {
        miniLayer = document.createElement('div');
        miniLayer.className = 'mini-layer';
        const segments = buildMiniSegments(task);
        segments.forEach(() => {
          miniLayer.appendChild(createMiniSegmentElement());
        });
        standdownLayer = document.createElement('div');
        standdownLayer.className = 'standdown-layer';
        bar.appendChild(miniLayer);
        bar.appendChild(progress);
        bar.appendChild(standdownLayer);
        refreshStanddownLayer(bar);
        bar.appendChild(label);
        const handleStart = document.createElement('div');
        handleStart.className = 'handle handle-start';
        const handleEnd = document.createElement('div');
        handleEnd.className = 'handle handle-end';
        bar.appendChild(handleStart);
        bar.appendChild(handleEnd);
        setBarLabel(bar, task, startDate, endDate);
        updateMiniLayerElements(bar, segments, { task });
        attachResizeHandles(bar, task);
        attachProgressHandle(bar, task);
        enableBarDrag(bar, task);
      } else {
        standdownLayer = document.createElement('div');
        standdownLayer.className = 'standdown-layer';
        bar.appendChild(progress);
        bar.appendChild(standdownLayer);
        refreshStanddownLayer(bar);
        bar.appendChild(label);
        setBarLabel(bar, task, startDate, endDate);
      }

      bar.addEventListener('click', (event) => {
        event.stopPropagation();
        selectTask(task.id);
      });

      row.appendChild(bar);
      row.addEventListener('click', () => selectTask(task.id));
      ganttBodyEl.appendChild(row);
    });
    adjustGanttOffset();
    updateZoomControls();
    if (ganttBodyScrollEl) ganttBodyScrollEl.scrollTop = previousScrollTop;
    scheduleTimelineHorizontalSync();
  }

  function updateTaskStartDate(task, newStartDate) {
    if (task?.singleEvent) {
      const start = ensureWeekday(newStartDate, 1);
      const normalized = formatDate(start);
      task.startDate = normalized;
      task.endDate = normalized;
      return;
    }
    const currentEnd = parseDate(task.endDate) || newStartDate;
    let duration = diffInWeekdays(newStartDate, currentEnd);
    const minDuration = minimumTaskDuration(task);
    if (duration < minDuration) {
      const adjustedEnd = shiftWeekdays(newStartDate, minDuration - 1);
      task.endDate = formatDate(adjustedEnd);
      duration = diffInWeekdays(newStartDate, adjustedEnd);
    }
    reconcileMiniDurationsForTask(task, Math.max(1, duration), {
      anchor: 'start',
      referenceStart: newStartDate
    });
  }

  function updateTaskEndDate(task, newEndDate) {
    if (task?.singleEvent) {
      const normalizedStart = formatDate(ensureWeekday(parseDate(task.startDate) || newEndDate, 1));
      task.startDate = normalizedStart;
      task.endDate = normalizedStart;
      return;
    }
    const currentStart = parseDate(task.startDate) || newEndDate;
    let duration = diffInWeekdays(currentStart, newEndDate);
    const minDuration = minimumTaskDuration(task);
    if (duration < minDuration) {
      newEndDate = shiftWeekdays(currentStart, minDuration - 1);
      duration = diffInWeekdays(currentStart, newEndDate);
    }
    reconcileMiniDurationsForTask(task, Math.max(1, duration), {
      anchor: 'end',
      referenceEnd: newEndDate
    });
  }

  function applyNewStartDate(task, newStartDate, fixedEndDate) {
    if (task?.singleEvent) {
      const normalized = formatDate(ensureWeekday(newStartDate, 1));
      task.startDate = normalized;
      task.endDate = normalized;
      return;
    }
    let start = ensureWeekday(newStartDate, 1);
    let end = ensureWeekday(fixedEndDate, -1);
    if (compareDates(end, start) < 0) {
      end = shiftWeekdays(start, minimumTaskDuration(task) - 1);
    }
    const duration = Math.max(minimumTaskDuration(task), diffInWeekdays(start, end));
    end = shiftWeekdays(start, duration - 1);
    reconcileMiniDurationsForTask(task, duration, {
      anchor: 'start',
      referenceStart: start
    });
  }

  function applyNewEndDate(task, fixedStartDate, newEndDate) {
    if (task?.singleEvent) {
      const normalized = formatDate(ensureWeekday(parseDate(task.startDate) || fixedStartDate, 1));
      task.startDate = normalized;
      task.endDate = normalized;
      return;
    }
    let start = ensureWeekday(fixedStartDate, 1);
    let end = ensureWeekday(newEndDate, -1);
    if (compareDates(end, start) < 0) {
      end = shiftWeekdays(start, minimumTaskDuration(task) - 1);
    }
    const duration = Math.max(minimumTaskDuration(task), diffInWeekdays(start, end));
    end = shiftWeekdays(start, duration - 1);
    reconcileMiniDurationsForTask(task, duration, {
      anchor: 'end',
      referenceEnd: end
    });
  }

  function adjustGanttOffset() {
    if (!ganttBodyEl) return;
    ganttBodyEl.style.marginTop = '0px';
  }

  function updateStickyMetrics() {
    const topOffset = topBarEl ? topBarEl.getBoundingClientRect().height : 0;
    document.documentElement.style.setProperty('--top-offset', `${topOffset}px`);
    updateHeaderOffset();

    const sidebarHeaderEl = document.querySelector('.sidebar-header');
    const sidebarHeaderHeight = sidebarHeaderEl ? sidebarHeaderEl.getBoundingClientRect().height : 0;
    document.documentElement.style.setProperty('--sidebar-header-height', `${sidebarHeaderHeight}px`);

    const ganttHeaderEl = document.querySelector('.gantt-header');
    const ganttHeaderHeight = ganttHeaderEl ? ganttHeaderEl.getBoundingClientRect().height : 0;
    document.documentElement.style.setProperty('--gantt-header-height', `${ganttHeaderHeight}px`);

    const rootStyles = getComputedStyle(document.documentElement);
    const rootFontSize = parseFloat(rootStyles.fontSize) || 16;
    const toPixels = (raw) => {
      if (!raw) return 0;
      const trimmed = raw.trim();
      if (trimmed.endsWith('rem')) return parseFloat(trimmed) * rootFontSize;
      if (trimmed.endsWith('px')) return parseFloat(trimmed);
      return parseFloat(trimmed) || 0;
    };

    const taskHeaderHeightPx = toPixels(rootStyles.getPropertyValue('--task-header-height'));
    const timelineHeaderHeightPx = toPixels(rootStyles.getPropertyValue('--timeline-header-height'));
    const contentOffsetPx = toPixels(rootStyles.getPropertyValue('--content-offset'));
    const verticalHeaderOffset = toPixels(rootStyles.getPropertyValue('--vertical-header-offset'));
    const bodyOffset = toPixels(rootStyles.getPropertyValue('--body-offset'));

    const taskHeaderTopPx = Math.max(0, topOffset + contentOffsetPx - verticalHeaderOffset);
    const timelineHeaderTopPx = Math.max(0, topOffset + contentOffsetPx - verticalHeaderOffset);
    const rowHeight = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--row-height')) || 54;

    document.documentElement.style.setProperty('--task-header-top', `${taskHeaderTopPx}px`);
    document.documentElement.style.setProperty('--timeline-header-top', `${timelineHeaderTopPx}px`);

    const ganttHeaderElHeight = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--gantt-header-height')) || 0;
    const ganttScrollStartPx = timelineHeaderTopPx + timelineHeaderHeightPx + bodyOffset;
    const availableTaskHeight = Math.max(120, window.innerHeight - (ganttScrollStartPx + ganttHeaderElHeight));
    const availableGanttHeight = Math.max(120, window.innerHeight - (ganttScrollStartPx + ganttHeaderElHeight));
    const unifiedScrollHeight = Math.floor(Math.max(availableTaskHeight, availableGanttHeight) - rowHeight - rowHeight - rowHeight);

    document.documentElement.style.setProperty('--task-scroll-height', `${unifiedScrollHeight}px`);
    document.documentElement.style.setProperty('--gantt-scroll-height', `${unifiedScrollHeight}px`);
  }

  function syncScrollPositions(source, target) {
    if (!source || !target) return;
    if (state.isSyncingScroll) return;
    state.isSyncingScroll = true;
    target.scrollTop = source.scrollTop;
    requestAnimationFrame(() => {
      state.isSyncingScroll = false;
    });
  }

  function updateTimelineHorizontalOffset() {
    if (!timelineHeaderEl) return;
    const offset = ganttBodyScrollEl ? Number(ganttBodyScrollEl.scrollLeft) || 0 : 0;
    if (Math.abs(offset) < 0.5) {
      timelineHeaderEl.style.transform = '';
    } else {
      timelineHeaderEl.style.transform = `translateX(${-offset}px)`;
    }
  }

  function scheduleTimelineHorizontalSync() {
    if (state.timelineHorizontalSyncFrame !== null) return;
    state.timelineHorizontalSyncFrame = requestAnimationFrame(() => {
      state.timelineHorizontalSyncFrame = null;
      updateTimelineHorizontalOffset();
    });
  }

  return {
    renderGantt,
    adjustZoom,
    setTimelineViewMode,
    updateViewModeButtons,
    updateStickyMetrics,
    scheduleTimelineHorizontalSync,
    syncScrollPositions,
    updateTaskStartDate,
    updateTaskEndDate,
    getTimelineDays: () => state.timelineDays,
    getTimelineViewMode: () => state.timelineViewMode,
    adjustGanttOffset,
    updateZoomControls
  };
}
