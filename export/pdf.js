let pdfLogoPromise = null;
let jsPdfConstructorCache = null;
let jsPdfLoaderPromise = null;

async function loadPdfLogo() {
  if (!pdfLogoPromise) {
    pdfLogoPromise = (async () => {
      try {
        const response = await fetch('assets/bb-logo.png', { cache: 'no-store' });
        if (!response || !response.ok) throw new Error(`Logo request failed: ${response?.status}`);
        const blob = await response.blob();
        const arrayBuffer = await blob.arrayBuffer();
        const bytes = new Uint8Array(arrayBuffer);
        let binary = '';
        for (let i = 0; i < bytes.length; i += 1) {
          binary += String.fromCharCode(bytes[i]);
        }
        const base64 = btoa(binary);
        const dataUrl = `data:image/png;base64,${base64}`;
        const dimensions = await new Promise((resolve, reject) => {
          const img = new Image();
          img.onload = () => {
            resolve({
              width: img.naturalWidth || img.width || 1,
              height: img.naturalHeight || img.height || 1
            });
          };
          img.onerror = (error) => reject(error);
          img.src = dataUrl;
        });
        return { dataUrl, ...dimensions };
      } catch (error) {
        console.warn('Unable to load PDF logo asset.', error);
        return null;
      }
    })();
  }
  return pdfLogoPromise;
}

function setGlobalJsPdf(ctor) {
  if (!ctor) return;
  if (!window.jspdf) window.jspdf = { jsPDF: ctor };
  else if (typeof window.jspdf.jsPDF !== 'function') window.jspdf.jsPDF = ctor;
}

function loadJsPdfFromScript(options) {
  if (window.jspdf && typeof window.jspdf.jsPDF === 'function') {
    return Promise.resolve(window.jspdf.jsPDF);
  }
  const { src, integrity } = options;
  return new Promise((resolve) => {
    const existing = document.querySelector(`script[data-jspdf-source="${src}"]`);
    const finalize = () => {
      const ctor = window.jspdf && typeof window.jspdf.jsPDF === 'function' ? window.jspdf.jsPDF : null;
      if (ctor) setGlobalJsPdf(ctor);
      resolve(ctor);
    };
    const handleError = () => resolve(null);
    if (existing) {
      existing.addEventListener('load', finalize, { once: true });
      existing.addEventListener('error', handleError, { once: true });
      return;
    }
    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.dataset.jspdfSource = src;
    script.referrerPolicy = 'no-referrer';
    if (integrity) {
      script.integrity = integrity;
      script.crossOrigin = 'anonymous';
    }
    script.addEventListener('load', finalize, { once: true });
    script.addEventListener('error', handleError, { once: true });
    document.head.appendChild(script);
  });
}

async function loadJsPdfFallbacks() {
  const sources = [
    {
      src: 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
      integrity: 'sha512-2JvIUMI5KOZ6+Jo7oBMnp1vOSfQ1+avG4v8M45bNaMdWOh9fvydMBaTlM3Dd3yXTVzTuvq8rqBiTWLY4dTqHOg=='
    },
    {
      src: 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js'
    }
  ];
  for (let index = 0; index < sources.length; index += 1) {
    const ctor = await loadJsPdfFromScript(sources[index]);
    if (typeof ctor === 'function') return ctor;
  }
  return null;
}

async function ensureJsPdfLoaded() {
  if (jsPdfConstructorCache) return jsPdfConstructorCache;
  if (window.jspdf && typeof window.jspdf.jsPDF === 'function') {
    jsPdfConstructorCache = window.jspdf.jsPDF;
    return jsPdfConstructorCache;
  }
  if (jsPdfLoaderPromise) return jsPdfLoaderPromise;
  jsPdfLoaderPromise = import('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.es.min.js')
    .then((module) => {
      const ctor = module?.jsPDF || module?.default || (window.jspdf && window.jspdf.jsPDF);
      if (typeof ctor === 'function') {
        jsPdfConstructorCache = ctor;
        setGlobalJsPdf(ctor);
      }
      return jsPdfConstructorCache;
    })
    .catch(async (error) => {
      console.error('Failed to load jsPDF ESM bundle', error);
      const fallbackCtor = await loadJsPdfFallbacks();
      if (typeof fallbackCtor === 'function') {
        jsPdfConstructorCache = fallbackCtor;
        setGlobalJsPdf(fallbackCtor);
      }
      return jsPdfConstructorCache;
    })
    .finally(() => {
      jsPdfLoaderPromise = null;
    });
  return jsPdfLoaderPromise;
}

function ensureFunction(fn, name) {
  if (typeof fn !== 'function') {
    throw new Error(`createPdfExporter requires a "${name}" function`);
  }
}

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function alertUser(message) {
  if (typeof window !== 'undefined' && typeof window.alert === 'function') {
    window.alert(message);
  }
}

export function createPdfExporter(config = {}) {
  const {
    getTasks,
    getTimelineDays,
    getTimelineViewMode = () => 'days',
    renderGantt,
    updateAllStageSummaries,
    getOrderedTasks,
    buildTaskIdentifiers,
    getTodayUtc,
    compareDates,
    formatDate,
    parseDate,
    buildMilestoneSegments,
    buildMiniSegments,
    getMiniColor,
    getTaskDuration,
    isStanddownDay,
    groupTimelineDaysByYear,
    parseColorToRgb,
    taskColorMap = {},
    defaultSegmentColor = '#3056d3'
  } = config;

  [
    ['getTasks', getTasks],
    ['getTimelineDays', getTimelineDays],
    ['updateAllStageSummaries', updateAllStageSummaries],
    ['getOrderedTasks', getOrderedTasks],
    ['buildTaskIdentifiers', buildTaskIdentifiers],
    ['getTodayUtc', getTodayUtc],
    ['compareDates', compareDates],
    ['formatDate', formatDate],
    ['parseDate', parseDate],
    ['buildMilestoneSegments', buildMilestoneSegments],
    ['buildMiniSegments', buildMiniSegments],
    ['getMiniColor', getMiniColor],
    ['getTaskDuration', getTaskDuration],
    ['isStanddownDay', isStanddownDay],
    ['groupTimelineDaysByYear', groupTimelineDaysByYear],
    ['parseColorToRgb', parseColorToRgb]
  ].forEach(([name, fn]) => ensureFunction(fn, name));

  const getViewMode = typeof getTimelineViewMode === 'function' ? getTimelineViewMode : () => 'days';

  return async function handleSavePdf() {
    const jsPdfCtor = await ensureJsPdfLoaded();
    if (typeof jsPdfCtor !== 'function') {
      console.error('jsPDF library failed to load.');
      alertUser('Unable to save the PDF because the export component failed to load.');
      return;
    }
    const tasks = asArray(getTasks());
    if (!tasks.length) {
      alertUser('Add at least one task before saving the PDF.');
      return;
    }
    let timelineDays = asArray(getTimelineDays());
    if (!timelineDays.length && typeof renderGantt === 'function') {
      renderGantt();
      timelineDays = asArray(getTimelineDays());
    }
    if (!timelineDays.length) {
      alertUser('Generate the schedule before exporting it as a PDF.');
      return;
    }

    if (typeof updateAllStageSummaries === 'function') {
      updateAllStageSummaries();
    }

    const timelineViewMode = getViewMode() || 'days';
    const doc = new jsPdfCtor({ orientation: 'landscape', unit: 'mm', format: 'a3' });
    const margin = 8;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const contentWidth = pageWidth - (margin * 2);
    const contentHeight = pageHeight - (margin * 2);
    const leftColumnWidth = 95;
    const rightColumnWidth = contentWidth - leftColumnWidth;
    const columnLabels = ['ID', 'Task', 'Start', 'End', 'Duration', 'Progress'];
    const columnWeights = [45, 115, 85, 85, 60, 70];
    const totalWeight = columnWeights.reduce((sum, weight) => sum + weight, 0);
    const columnWidths = columnWeights.map((weight) => (weight / totalWeight) * leftColumnWidth);
    const orderedEntries = getOrderedTasks();
    const identifierMap = buildTaskIdentifiers(orderedEntries);
    const tasksSnapshot = orderedEntries.map(({ task, depth }) => ({
      ...task,
      depth: depth || 0,
      identifier: identifierMap.get(task.id) || '—',
      miniTasks: Array.isArray(task.miniTasks) ? task.miniTasks.map((mini) => ({ ...mini })) : []
    }));
    const dayCount = timelineDays.length;
    const dayWidth = dayCount ? rightColumnWidth / dayCount : rightColumnWidth;
    const timelineRowHeight = 6;
    const timelineRowCount = 2;
    const timelineHeaderHeight = timelineRowHeight * timelineRowCount;
    const textPadding = 1.5;
    const timelineStartX = margin + leftColumnWidth;
    const pageBottom = margin + contentHeight;
    const defaultTextColor = { r: 31, g: 41, b: 55 };
    const defaultDrawColor = { r: 214, g: 219, b: 231 };
    const headerFill = { r: 234, g: 238, b: 248 };
    const zebraFill = { r: 246, g: 248, b: 252 };
    const stageFill = { r: 250, g: 250, b: 210 };
    const legendItems = Object.entries(taskColorMap || {});
    const supportsOpacity = typeof doc.GState === 'function';
    const progressOverlayState = supportsOpacity ? doc.GState({ opacity: 0.35 }) : null;
    const resetOpacityState = supportsOpacity ? doc.GState({ opacity: 1 }) : null;
    const todayUtcPdf = getTodayUtc();
    const todayIndexPdf = todayUtcPdf ? timelineDays.findIndex((day) => compareDates(day, todayUtcPdf) === 0) : -1;
    const todayMarker = todayIndexPdf !== -1 ? {
      index: todayIndexPdf,
      x: timelineStartX + (todayIndexPdf * dayWidth)
    } : null;
    const pdfTimelineIndexMap = new Map();
    timelineDays.forEach((day, index) => {
      pdfTimelineIndexMap.set(formatDate(day), index);
    });
    const getPdfDayIndex = (date) => {
      if (!(date instanceof Date)) return -1;
      const key = formatDate(date);
      if (pdfTimelineIndexMap.has(key)) return pdfTimelineIndexMap.get(key);
      for (let i = 0; i < timelineDays.length; i += 1) {
        if (compareDates(timelineDays[i], date) === 0) return i;
      }
      return -1;
    };
    const highlightPriority = { standdown: 0, deadline: 1, today: 2 };
    const highlightIndexOffset = 0;
    const standdownFillColor = { r: 220, g: 223, b: 230 };
    const standdownBorderColor = { r: 0, g: 0, b: 0 };
    const standdownBoundaryLineWidth = 0.8;
    const highlightPalette = {
      today: { fill: { r: 255, g: 228, b: 232 }, border: { r: 220, g: 38, b: 38 } },
      deadline: { fill: { r: 255, g: 243, b: 207 }, border: { r: 217, g: 119, b: 6 } },
      standdown: { fill: standdownFillColor, border: standdownBorderColor }
    };
    const timelineDateFontSize = 4;
    const mondayLineWidth = 0.35;
    const mondayLineColor = { r: 0, g: 0, b: 0 };
    const mondayLabelOffset = 1.2; // keeps header date tucked against the Monday divider
    const formatDayMonth = (date) => {
      if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
      const day = String(date.getUTCDate()).padStart(2, '0');
      const month = String(date.getUTCMonth() + 1).padStart(2, '0');
      return `${day}/${month}`;
    };
    const getWeekMonday = (date) => {
      if (!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
      const monday = new Date(date.getTime());
      const dayOfWeek = monday.getUTCDay();
      const offset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      monday.setUTCDate(monday.getUTCDate() + offset);
      return monday;
    };
    const mondayBoundaryEntries = [];
    timelineDays.forEach((day, index) => {
      if (day instanceof Date && day.getUTCDay() === 1) {
        mondayBoundaryEntries.push({ day, index });
      }
    });
    if (!mondayBoundaryEntries.length && timelineDays.length) {
      mondayBoundaryEntries.push({ day: timelineDays[0], index: 0 });
    }
    const mondayBoundarySet = new Set(mondayBoundaryEntries.map((entry) => entry.index));
    const mondayLabelSpans = mondayBoundaryEntries.map((entry, idx) => {
      const nextIndex = mondayBoundaryEntries[idx + 1]?.index ?? dayCount;
      const length = Math.max(1, nextIndex - entry.index);
      return { ...entry, length };
    });

    const logoImage = await loadPdfLogo();
    const highlightSegmentsPdf = [];
    if (todayMarker) {
      highlightSegmentsPdf.push({
        type: 'today',
        startIndex: todayMarker.index,
        endIndex: todayMarker.index
      });
    }
    const milestoneSegmentsPdf = buildMilestoneSegments(timelineDays);
    milestoneSegmentsPdf.forEach((segment) => {
      highlightSegmentsPdf.push(segment);
    });
    const normalizedHighlightSegmentsRaw = highlightSegmentsPdf
      .map((segment) => {
        let startIndex = Number.isFinite(segment.startIndex) ? segment.startIndex : NaN;
        let endIndex = Number.isFinite(segment.endIndex) ? segment.endIndex : NaN;
        if (!Number.isFinite(startIndex) || !Number.isFinite(endIndex)) return null;
        startIndex = Math.max(0, Math.min(dayCount - 1, startIndex));
        endIndex = Math.max(startIndex, Math.min(dayCount - 1, endIndex));
        if (startIndex > endIndex) return null;
        return { ...segment, startIndex, endIndex };
      })
      .filter(Boolean);
    const normalizedHighlightSegments = normalizedHighlightSegmentsRaw
      .map((segment) => {
        let startIndex = Number.isFinite(segment.startIndex) ? segment.startIndex + highlightIndexOffset : NaN;
        let endIndex = Number.isFinite(segment.endIndex) ? segment.endIndex + highlightIndexOffset : NaN;
        if (!Number.isFinite(startIndex) || !Number.isFinite(endIndex)) return null;
        startIndex = Math.max(0, Math.min(dayCount - 1, startIndex));
        endIndex = Math.max(startIndex, Math.min(dayCount - 1, endIndex));
        if (startIndex > endIndex) return null;
        return { ...segment, startIndex, endIndex };
      })
      .filter(Boolean);
    const segmentsForDrawing = normalizedHighlightSegments
      .slice()
      .sort((a, b) => {
        const priorityDifference = (highlightPriority[a.type] ?? 0) - (highlightPriority[b.type] ?? 0);
        if (priorityDifference !== 0) return priorityDifference;
        return a.startIndex - b.startIndex;
      });
    const pdfDayHighlights = timelineDays.map(() => new Set());
    normalizedHighlightSegmentsRaw.forEach((segment) => {
      for (let index = segment.startIndex; index <= segment.endIndex; index += 1) {
        pdfDayHighlights[index]?.add(segment.type);
      }
    });
    const pdfDayHighlightsShifted = timelineDays.map(() => new Set());
    normalizedHighlightSegments.forEach((segment) => {
      for (let index = segment.startIndex; index <= segment.endIndex; index += 1) {
        pdfDayHighlightsShifted[index]?.add(segment.type);
      }
    });
    const standdownSegmentsPdf = normalizedHighlightSegments.filter((segment) => segment.type === 'standdown');
    const boundaryTypes = {};
    const applyBoundaryType = (index, type) => {
      if (index < 0 || index > dayCount) return;
      const existing = boundaryTypes[index];
      if (!existing || (highlightPriority[type] ?? 0) >= (highlightPriority[existing] ?? 0)) {
        boundaryTypes[index] = type;
      }
    };
    segmentsForDrawing.forEach((segment) => {
      applyBoundaryType(segment.startIndex, segment.type);
      applyBoundaryType(segment.endIndex + 1, segment.type);
    });

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.setLineHeightFactor(1.2);
    doc.setTextColor(defaultTextColor.r, defaultTextColor.g, defaultTextColor.b);
    doc.setLineWidth(0.2);

    const truncateText = (value, maxWidth) => {
      if (!value) return '';
      let text = String(value);
      if (doc.getTextWidth(text) <= maxWidth) return text;
      const ellipsis = '...';
      while (text.length > 0 && doc.getTextWidth(`${text}${ellipsis}`) > maxWidth) {
        text = text.slice(0, -1);
      }
      return text.length ? `${text}${ellipsis}` : ellipsis;
    };

    const lineHeight = (doc.getFontSize() * doc.getLineHeightFactor()) / doc.internal.scaleFactor;
    const minRowHeight = 12;
    const verticalPadding = 2;
    const nameColumnWidth = Math.max(10, columnWidths[1] - (textPadding * 2));

    const formatDateCell = (value) => {
      const date = parseDate(value);
      if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
        return ['--/--', '----'];
      }
      const day = String(date.getUTCDate()).padStart(2, '0');
      const month = String(date.getUTCMonth() + 1).padStart(2, '0');
      const year = String(date.getUTCFullYear());
      return [`${day}/${month}`, year];
    };

    const buildNameLines = (rowTask) => {
      const depth = Math.max(0, rowTask?.depth || 0);
      const indent = depth ? ''.padStart(depth * 2, ' ') : '';
      const baseName = (rowTask?.name || '').trim() || (rowTask?.entryType === 'stage' ? 'Stage' : 'Untitled task');
      const label = `${indent}${baseName}`;
      let lines = doc.splitTextToSize(label, nameColumnWidth);
      if (!Array.isArray(lines) || !lines.length) lines = ['Untitled task'];
      const maxLines = 3;
      if (lines.length > maxLines) {
        const limited = lines.slice(0, maxLines);
        const ellipsis = '...';
        let last = limited[maxLines - 1].trimEnd();
        while (last.length && doc.getTextWidth(`${last}${ellipsis}`) > nameColumnWidth) {
          last = last.slice(0, -1).trimEnd();
        }
        limited[maxLines - 1] = last ? `${last}${ellipsis}` : ellipsis;
        return limited;
      }
      return lines;
    };

    const computeRowLayout = (task) => {
      const duration = getTaskDuration(task);
      const progressValue = Math.min(100, Math.max(0, Number.isFinite(task.progress) ? task.progress : 0));
      const columnLines = [
        [task.identifier || task.id || '—'],
        buildNameLines(task),
        formatDateCell(task.startDate),
        formatDateCell(task.endDate),
        [String(duration), 'Days'],
        [`${Math.round(progressValue)}%`]
      ].map((cellLines) => cellLines.map((line) => (line === undefined || line === null ? '' : String(line))));
      const lineCounts = columnLines.map((lines) => Math.max(1, lines.length));
      const maxLineCount = Math.max(...lineCounts);
      const rowHeight = Math.max(minRowHeight, (maxLineCount * lineHeight) + (verticalPadding * 2));
      return { rowHeight, columnLines, lineCounts, duration, progressValue };
    };

    const applyHighlightTextStyle = (types) => {
      if (types.has('today')) {
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(220, 38, 38);
        return true;
      }
      if (types.has('deadline')) {
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(180, 83, 9);
        return true;
      }
      if (types.has('standdown')) {
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(51, 65, 85);
        return true;
      }
      return false;
    };

    const resetHighlightTextStyle = () => {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(timelineDateFontSize);
      doc.setTextColor(defaultTextColor.r, defaultTextColor.g, defaultTextColor.b);
    };

    const drawVerticalDateLabel = (text, centreX, centreY, highlights) => {
      const highlightSet = highlights instanceof Set ? highlights : new Set();
      const styled = highlightSet.size > 0 ? applyHighlightTextStyle(highlightSet) : false;
      const supportsRotation = typeof doc.saveGraphicsState === 'function' && typeof doc.rotate === 'function';
      if (supportsRotation) {
        doc.saveGraphicsState();
        doc.rotate(90, { origin: [centreX, centreY] });
        doc.text(text, centreX, centreY, { align: 'center', baseline: 'middle' });
        doc.restoreGraphicsState();
      } else {
        doc.text(text, centreX, centreY, { align: 'center', angle: 90, baseline: 'middle' });
      }
      if (styled) resetHighlightTextStyle();
    };

    const drawLegend = (topY, startX = margin) => {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.setTextColor(defaultTextColor.r, defaultTextColor.g, defaultTextColor.b);
      if (!legendItems.length) {
        doc.text('Legend: not configured', margin, topY + 4);
        return topY + 8;
      }
      const swatchSize = 4;
      const legendLineHeight = swatchSize + 2.5;
      let x = startX;
      let y = topY;
      const maxX = margin + contentWidth;
      legendItems.forEach(([label, color]) => {
        const textWidth = doc.getTextWidth(label);
        const requiredWidth = swatchSize + 2 + textWidth + 6;
        if (x + requiredWidth > maxX) {
          x = startX;
          y += legendLineHeight;
        }
        const rgb = parseColorToRgb(color);
        doc.setFillColor(rgb.r, rgb.g, rgb.b);
        doc.rect(x, y, swatchSize, swatchSize, 'F');
        doc.setDrawColor(160, 167, 187);
        doc.rect(x, y, swatchSize, swatchSize);
        doc.setTextColor(defaultTextColor.r, defaultTextColor.g, defaultTextColor.b);
        doc.text(label, x + swatchSize + 2, y + swatchSize - 0.6);
        x += requiredWidth;
      });
      doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      return y + legendLineHeight;
    };

    const drawDocumentHeader = (topY) => {
      let textStartX = margin;
      let legendStartY = topY;
      if (logoImage && logoImage.dataUrl) {
        const aspect = logoImage.width && logoImage.height ? logoImage.width / logoImage.height : 1;
        const logoHeight = 16;
        const logoWidth = logoHeight * (Number.isFinite(aspect) && aspect > 0 ? aspect : 1);
        try {
          doc.addImage(logoImage.dataUrl, 'PNG', margin, topY, logoWidth, logoHeight);
          textStartX = margin + logoWidth + 4;
          legendStartY = topY + logoHeight + 2;
        } catch (error) {
          console.warn('Unable to add logo to PDF header.', error);
        }
      }
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(18);
      doc.setTextColor(defaultTextColor.r, defaultTextColor.g, defaultTextColor.b);
      const titleBaseline = topY + 7;
      doc.text('BB: Project Programme', textStartX, titleBaseline);
      const legendTop = Math.max(titleBaseline + 2.5, legendStartY);
      const legendBottom = drawLegend(legendTop, margin);
      return legendBottom + 2;
    };

    const drawTimelineGridLines = (topY, height, options = {}) => {
      if (!dayCount) return;
      const { highlightsOnly = false } = options;
      const originalLineWidth = doc.getLineWidth();
      for (let index = 0; index <= dayCount; index += 1) {
        const x = timelineStartX + (index * dayWidth);
        const boundaryType = boundaryTypes[index];
        if (highlightsOnly) {
          if (!boundaryType || boundaryType === 'standdown' || mondayBoundarySet.has(index)) continue;
          const palette = highlightPalette[boundaryType];
          if (!palette) continue;
          doc.setDrawColor(palette.border.r, palette.border.g, palette.border.b);
          doc.setLineWidth(boundaryType === 'today' ? 0.45 : 0.3);
          doc.line(x, topY, x, topY + height);
          continue;
        }
        if (!mondayBoundarySet.has(index)) continue;
        doc.setDrawColor(mondayLineColor.r, mondayLineColor.g, mondayLineColor.b);
        doc.setLineWidth(mondayLineWidth);
        doc.line(x, topY, x, topY + height);
      }
      doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      doc.setLineWidth(originalLineWidth);
    };

    const drawStanddownHatch = (startX, topY, width, height) => {
      if (width <= 0 || height <= 0) return;
      const originalLineWidth = doc.getLineWidth();
      doc.setFillColor(standdownFillColor.r, standdownFillColor.g, standdownFillColor.b);
      doc.rect(startX, topY, width, height, 'F');
      doc.setDrawColor(standdownBorderColor.r, standdownBorderColor.g, standdownBorderColor.b);
      doc.setLineWidth(standdownBoundaryLineWidth);
      doc.line(startX, topY, startX, topY + height);
      doc.line(startX + width, topY, startX + width, topY + height);
      doc.setLineWidth(originalLineWidth);
      doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      doc.setFillColor(255, 255, 255);
    };

    const drawStanddownBands = (topY, height) => {
      if (!standdownSegmentsPdf.length) return;
      standdownSegmentsPdf.forEach((segment) => {
        const startX = timelineStartX + (segment.startIndex * dayWidth);
        const width = (segment.endIndex - segment.startIndex + 1) * dayWidth;
        drawStanddownHatch(startX, topY, width, height);
      });
    };

    const drawTimelineHeader = (topY) => {
      doc.setFillColor(244, 246, 252);
      doc.rect(timelineStartX, topY, rightColumnWidth, timelineHeaderHeight, 'F');
      doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      doc.rect(timelineStartX, topY, rightColumnWidth, timelineHeaderHeight);

      segmentsForDrawing.forEach((segment) => {
        const palette = highlightPalette[segment.type];
        const startX = timelineStartX + (segment.startIndex * dayWidth);
        const width = (segment.endIndex - segment.startIndex + 1) * dayWidth;
        if (segment.type === 'standdown') {
          drawStanddownHatch(startX, topY, width, timelineHeaderHeight);
          return;
        }
        if (!palette?.fill) return;
        doc.setFillColor(palette.fill.r, palette.fill.g, palette.fill.b);
        doc.rect(startX, topY, width, timelineHeaderHeight, 'F');
      });

      doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      doc.rect(timelineStartX, topY, rightColumnWidth, timelineHeaderHeight);

      const yearBaseline = topY + timelineRowHeight - 1;
      const dateCenterY = topY + timelineRowHeight + (timelineRowHeight / 2);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8);
      const yearGroups = groupTimelineDaysByYear(timelineDays);
      yearGroups.forEach((group) => {
        const startX = timelineStartX + (group.startIndex * dayWidth);
        const centreX = startX + ((group.length * dayWidth) / 2);
        doc.text(String(group.year), centreX, yearBaseline, { align: 'center' });
      });

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(timelineDateFontSize);
      mondayLabelSpans.forEach(({ day, index, length }) => {
        const spanLength = Math.max(1, Math.min(dayCount, length));
        const mondayLineX = timelineStartX + (index * dayWidth);
        const labelAnchorX = mondayLineX + mondayLabelOffset;
        const combinedHighlights = new Set();
        const spanStart = index;
        const spanEnd = Math.min(dayCount - 1, index + spanLength - 1);
        for (let dayIndex = spanStart; dayIndex <= spanEnd; dayIndex += 1) {
          pdfDayHighlights[dayIndex]?.forEach((type) => combinedHighlights.add(type));
          pdfDayHighlightsShifted[dayIndex]?.forEach((type) => combinedHighlights.add(type));
        }
        const mondayDate = getWeekMonday(day);
        const mondayLabel = formatDayMonth(mondayDate || day);
        drawVerticalDateLabel(mondayLabel, labelAnchorX, dateCenterY, combinedHighlights);
      });

      drawTimelineGridLines(topY, timelineHeaderHeight);
      drawTimelineGridLines(topY, timelineHeaderHeight, { highlightsOnly: true });
    };

    const drawTaskRow = (task, topY, rowIndex, layout) => {
      const {
        rowHeight,
        columnLines,
        lineCounts,
        progressValue
      } = layout;
      const isStageRow = task.entryType === 'stage';
      if (isStageRow) {
        doc.setFillColor(stageFill.r, stageFill.g, stageFill.b);
        doc.rect(margin, topY, contentWidth, rowHeight, 'F');
      } else if (rowIndex % 2 === 0) {
        doc.setFillColor(zebraFill.r, zebraFill.g, zebraFill.b);
        doc.rect(margin, topY, contentWidth, rowHeight, 'F');
      }

      if (segmentsForDrawing.length && dayWidth > 0) {
        segmentsForDrawing.forEach((segment) => {
          const palette = highlightPalette[segment.type];
          if (!palette) return;
          const segmentX = timelineStartX + (segment.startIndex * dayWidth);
          const width = (segment.endIndex - segment.startIndex + 1) * dayWidth;
          doc.setFillColor(palette.fill.r, palette.fill.g, palette.fill.b);
          doc.rect(segmentX, topY, width, rowHeight, 'F');
        });
      }

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      let cellX = margin;
      const originalLineWidth = doc.getLineWidth();
      doc.setLineWidth(0.1);

      columnLines.forEach((lines, index) => {
        const width = columnWidths[index];
        const effectiveLines = Math.max(1, lineCounts[index]);
        const contentHeight = effectiveLines * lineHeight;
        const textTop = topY + Math.max(verticalPadding, (rowHeight - contentHeight) / 2);
        const cleanedLines = lines.length ? lines.map((line) => (line === '' ? ' ' : line)) : [' '];
        if (index === 0) {
          const originalFontSize = doc.getFontSize();
          doc.setFont('helvetica', 'normal');
          doc.setFontSize(7);
          doc.text(cleanedLines, cellX + textPadding, textTop, { baseline: 'top' });
          doc.setFontSize(originalFontSize);
        } else if (index === 1) {
          doc.text(cleanedLines, cellX + textPadding, textTop, { baseline: 'top' });
        } else {
          const centreX = cellX + (width / 2);
          doc.text(cleanedLines, centreX, textTop, { baseline: 'top', align: 'center' });
        }
        cellX += width;
        if (index < columnLines.length - 1) {
          doc.line(cellX, topY, cellX, topY + rowHeight);
        }
      });

      doc.setLineWidth(originalLineWidth);

      drawTimelineGridLines(topY, rowHeight);

      const startDate = parseDate(task.startDate);
      const endDate = parseDate(task.endDate);
      if (!startDate || !endDate || !dayCount) return;

      let startIndex = getPdfDayIndex(startDate);
      let endIndex = getPdfDayIndex(endDate);
      if (startIndex === -1 || endIndex === -1) return;
      startIndex = Math.max(0, Math.min(dayCount - 1, startIndex));
      endIndex = Math.max(startIndex, Math.min(dayCount - 1, endIndex));
      if (typeof startIndex !== 'number' || typeof endIndex !== 'number' || startIndex > endIndex) return;

      const rawBarX = timelineStartX + (startIndex * dayWidth);
      const rawBarWidth = (endIndex - startIndex + 1) * dayWidth;
      const barX = Math.round(rawBarX * 1000) / 1000;
      const barWidth = Math.max(dayWidth, Math.round(rawBarWidth * 1000) / 1000);
      const barHeight = rowHeight * 0.6;
      const barY = topY + ((rowHeight - barHeight) / 2);

      if (barWidth <= 0) return;

      const workingIndices = [];
      for (let index = startIndex; index <= endIndex; index += 1) {
        const day = timelineDays[index];
        if (!isStanddownDay(day)) workingIndices.push(index);
      }

      const segments = buildMiniSegments(task);
      const paintSegments = segments.length ? segments : [{ name: null, duration: workingIndices.length }];
      const paintedRects = [];
      let workingPointer = 0;
      paintSegments.forEach((segment) => {
        const duration = Math.max(0, segment.duration || 0);
        if (duration <= 0) return;
        const startWorkingIndex = workingIndices[workingPointer];
        if (startWorkingIndex == null) return;
        const endWorkingIndex = workingIndices[Math.min(
          workingPointer + duration - 1,
          workingIndices.length - 1
        )];
        if (endWorkingIndex == null || endWorkingIndex < startWorkingIndex) return;
        const left = (startWorkingIndex - startIndex) * dayWidth;
        const width = (endWorkingIndex - startWorkingIndex + 1) * dayWidth;
        const color = segment.name
          ? parseColorToRgb(getMiniColor(segment.name))
          : parseColorToRgb(defaultSegmentColor);
        doc.setFillColor(color.r, color.g, color.b);
        doc.rect(barX + left, barY, width, barHeight, 'F');
        paintedRects.push({ left, width });
        workingPointer += duration;
      });

      if (!paintedRects.length && workingIndices.length && !isStageRow) {
        const left = 0;
        const width = (workingIndices[workingIndices.length - 1] - startIndex + 1) * dayWidth;
        const fallback = parseColorToRgb(defaultSegmentColor);
        doc.setFillColor(fallback.r, fallback.g, fallback.b);
        doc.rect(barX + left, barY, width, barHeight, 'F');
        paintedRects.push({ left, width });
      }

      if (!isStageRow && progressValue > 0 && workingIndices.length) {
        let remaining = (progressValue / 100) * workingIndices.length;
        let prevIndex = null;
        let currentSegment = null;
        const overlaySegments = [];
        for (let i = 0; i < workingIndices.length && remaining > 1e-6; i += 1) {
          const idx = workingIndices[i];
          const left = (idx - startIndex) * dayWidth;
          const portion = Math.min(1, remaining);
          const contiguous = currentSegment && prevIndex != null && idx === prevIndex + 1 && !currentSegment.partial;
          if (!contiguous) {
            currentSegment = { left, width: portion * dayWidth, partial: portion < 1 };
            overlaySegments.push(currentSegment);
          } else {
            currentSegment.width += portion * dayWidth;
            currentSegment.partial = portion < 1;
          }
          remaining -= portion;
          prevIndex = idx;
          if (portion < 1) break;
        }

        overlaySegments.forEach((segment) => {
          const overlayX = barX + segment.left;
          const overlayWidth = Math.min(segment.width, barWidth);
          if (overlayWidth <= 0) return;
          if (progressOverlayState && resetOpacityState) {
            doc.setGState(progressOverlayState);
            doc.setFillColor(255, 255, 255);
            doc.rect(overlayX, barY, overlayWidth, barHeight, 'F');
            doc.setGState(resetOpacityState);
          }
          doc.setFillColor(255, 255, 255);
          const stripeWidth = 0.45;
          const stripeSpacing = 2;
          const stripeRight = overlayX + overlayWidth;
          for (let stripeX = overlayX; stripeX < stripeRight; stripeX += stripeSpacing) {
            const width = Math.min(stripeWidth, stripeRight - stripeX);
            if (width <= 0) continue;
            doc.rect(stripeX, barY, width, barHeight, 'F');
          }
        });
      }

      doc.setDrawColor(80, 94, 120);
      if (isStageRow) {
        doc.setFillColor(stageFill.r, stageFill.g, stageFill.b);
        doc.rect(barX, barY, barWidth, barHeight, 'F');
        doc.setDrawColor(212, 163, 18);
        doc.setLineWidth(0.3);
        doc.rect(barX, barY, barWidth, barHeight);
        doc.setLineWidth(originalLineWidth);
        doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      } else {
        doc.setDrawColor(80, 94, 120);
        paintedRects.forEach((rect) => {
          doc.rect(barX + rect.left, barY, rect.width, barHeight);
        });
        doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      }

      const labelMaxWidth = Math.max(barWidth - 4, 0);
      if (labelMaxWidth > 0) {
        const label = truncateText(task.name, labelMaxWidth);
        if (label) {
          doc.setFont('helvetica', 'bold');
          doc.setFontSize(7);
          doc.setTextColor(33, 45, 67);
          doc.text(label, barX + 2, barY + (barHeight / 2) + 0.2, { baseline: 'middle' });
          doc.setFont('helvetica', 'normal');
          doc.setFontSize(9);
          doc.setTextColor(defaultTextColor.r, defaultTextColor.g, defaultTextColor.b);
        }
      }

      drawStanddownBands(topY, rowHeight);

      doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      doc.setLineWidth(0.1);
      doc.line(margin, topY, margin + contentWidth, topY);
      doc.line(margin, topY + rowHeight, margin + contentWidth, topY + rowHeight);
      doc.line(timelineStartX, topY, timelineStartX, topY + rowHeight);
      doc.setLineWidth(originalLineWidth);

      drawTimelineGridLines(topY, rowHeight, { highlightsOnly: true });
    };

    const startNewPage = (isFirstPage) => {
      if (!isFirstPage) {
        doc.addPage();
      }
      doc.setTextColor(defaultTextColor.r, defaultTextColor.g, defaultTextColor.b);
      doc.setDrawColor(defaultDrawColor.r, defaultDrawColor.g, defaultDrawColor.b);
      doc.setLineWidth(0.2);
      let cursorY = drawDocumentHeader(margin);

      doc.setFillColor(headerFill.r, headerFill.g, headerFill.b);
      doc.rect(margin, cursorY, leftColumnWidth, timelineHeaderHeight, 'F');
      doc.rect(margin, cursorY, leftColumnWidth, timelineHeaderHeight);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      let headerX = margin;
      const headerBaseline = cursorY + 4;
      columnLabels.forEach((label, index) => {
        doc.text(label, headerX + textPadding, headerBaseline);
        headerX += columnWidths[index];
      });
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      drawTimelineHeader(cursorY);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      return cursorY + timelineHeaderHeight;
    };

    const ensurePageCapacity = (requiredHeight) => {
      if (currentY + requiredHeight > pageBottom) {
        currentY = startNewPage(false);
      }
    };

    let currentY = startNewPage(true);
    tasksSnapshot.forEach((task, index) => {
      const layout = computeRowLayout(task);
      ensurePageCapacity(layout.rowHeight);
      drawTaskRow(task, currentY, index, layout);
      currentY += layout.rowHeight;
    });

    doc.save('project-programme.pdf');
  };
}
