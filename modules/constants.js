const TASK_COLOR_MAP = {
  'Bond Bryan Design': 'rgb(176, 207, 255)',
  'Bond Bryan review': 'rgb(255, 192, 203)',
  'CDP': 'rgb(196, 235, 214)',
  'Client Comments': 'rgb(255, 223, 186)',
  'Delays': 'rgb(220, 220, 230)',
  'Undefined': 'rgb(245, 245, 245)'
};

const DEFAULT_SEGMENT_COLOR = '#3056d3';

const colorConstants = Object.freeze({
  TASK_COLOR_MAP: Object.freeze({ ...TASK_COLOR_MAP }),
  DEFAULT_SEGMENT_COLOR
});

export default colorConstants;
export { TASK_COLOR_MAP, DEFAULT_SEGMENT_COLOR };
