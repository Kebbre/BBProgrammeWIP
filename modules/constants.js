const TASK_COLOR_MAP = {
  'Bond Bryan Design': 'rgba(134, 176, 240, 1)',
  'Bond Bryan review': 'rgba(199, 201, 238, 1)',
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
