import { populateColumnSelects, renderTemplateOptions } from './ui.js';
import { attachEventHandlers } from './events.js';

const initApp = () => {
  populateColumnSelects();
  renderTemplateOptions();
  attachEventHandlers();
};

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}
