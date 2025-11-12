const TASK_COLOR_MAP = {
  'Bond Bryan Design': 'rgba(220, 230, 247, 1)',
  'Bond Bryan review': 'rgba(181, 182, 205, 1)',
  'CDP': 'rgba(218, 235, 226, 1)',
  'Client Comments': 'rgba(242, 225, 206, 1)',
  'Delays': 'rgba(215, 215, 215, 1)',
  'Undefined': 'rgba(249, 249, 249, 1)'
};

const DEFAULT_SEGMENT_COLOR = '#d0cac3ff';
const STAGE_COLOR = 'rgba(250, 250, 210, 1)';
const SINGLE_EVENT_COLOR = '#2fb879';
const SINGLE_EVENT_SELECTED_COLOR = '#228a5f';

const colorConstants = Object.freeze({
  TASK_COLOR_MAP: Object.freeze({ ...TASK_COLOR_MAP }),
  DEFAULT_SEGMENT_COLOR,
  STAGE_COLOR,
  SINGLE_EVENT_COLOR,
  SINGLE_EVENT_SELECTED_COLOR
});

export default colorConstants;
export {
  TASK_COLOR_MAP,
  DEFAULT_SEGMENT_COLOR,
  STAGE_COLOR,
  SINGLE_EVENT_COLOR,
  SINGLE_EVENT_SELECTED_COLOR
};
