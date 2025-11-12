const TASK_COLOR_MAP = {
  'Bond Bryan Design': 'rgba(220, 230, 247, 1)',
  'Bond Bryan review': 'rgba(181, 182, 205, 1)',
  'CDP': 'rgba(218, 235, 226, 1)',
  'Client Comments': 'rgba(242, 225, 206, 1)',
  'Delays': 'rgba(215, 215, 215, 1)',
  'Undefined': 'rgba(249, 249, 249, 1)'
};

const DEFAULT_SEGMENT_COLOR = '#3056d3';

const colorConstants = Object.freeze({
  TASK_COLOR_MAP: Object.freeze({ ...TASK_COLOR_MAP }),
  DEFAULT_SEGMENT_COLOR
});

export default colorConstants;
export { TASK_COLOR_MAP, DEFAULT_SEGMENT_COLOR };
