// On macOS Chrome reports a trackpad pinch as a ctrl+wheel event. Outlook
// handles that event itself and changes only its message content. Stop the
// event before Outlook receives it, but do not call preventDefault(): Chrome
// can then perform its normal, smooth browser-page pinch zoom.
const DEFAULTS = {
  enabled: true,
  nativePinch: true,
  archiveButton: true,
  emailSizeControls: true
};
let settings = { ...DEFAULTS };
let archiveLabelUpdateQueued = false;
let emailSizeControlsQueued = false;
let emailScale = 1;
let scaledEmailContent;
let emailSizeControls;
const ARCHIVE_LABEL_CLASS = 'enhanced-outlook-archive-label';
const EMAIL_SCALE_MIN = 0.8;
const EMAIL_SCALE_MAX = 1.6;
const EMAIL_SCALE_STEP = 0.1;

const ARCHIVE_CONTROL_SELECTOR = [
  '[data-automation-type="RibbonButton"][label="Archive"]',
  '[dataautomationtype="RibbonButton"][label="Archive"]',
  'button[aria-label^="Archive"]',
  '[role="button"][aria-label^="Archive"]',
  'button[title^="Archive"]',
  '[role="button"][title^="Archive"]',
  'button:has([aria-label^="Archive"])',
  '[role="button"]:has([aria-label^="Archive"])',
  'button:has([data-icon-name="Archive"])',
  '[role="button"]:has([data-icon-name="Archive"])'
].join(',');

chrome.storage.sync.get(DEFAULTS, (stored) => {
  settings = stored;
  applyVisualSettings();
});

chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== 'sync') return;
  for (const key of Object.keys(DEFAULTS)) {
    if (changes[key]) settings[key] = changes[key].newValue;
  }
  applyVisualSettings();
});

// Outlook replaces toolbar controls as its single-page UI changes, and can set
// their identifying attributes after inserting them. Add a real text node only
// to icon-only Archive controls. A CSS pseudo-element can be clipped by
// Outlook's toolbar internals, while a real node participates in the control's
// normal layout.
const outlookDomObserver = new MutationObserver((mutations) => {
  let archiveMayHaveChanged = false;
  let emailContentMayHaveChanged = false;

  for (const mutation of mutations) {
    if (mutation.type === 'childList') {
      archiveMayHaveChanged = true;
      emailContentMayHaveChanged = true;
      continue;
    }

    const target = mutation.target;
    if (isArchiveControlOrChild(target)) archiveMayHaveChanged = true;
    if (
      mutation.attributeName === 'role' &&
      target instanceof Element &&
      target.matches('[role="document"]')
    ) {
      emailContentMayHaveChanged = true;
    }
  }

  if (archiveMayHaveChanged) queueArchiveLabelUpdate();
  if (emailContentMayHaveChanged) queueEmailSizeControlsUpdate();
});

function startOutlookDomObserver() {
  // At document_start Chrome can inject into an Outlook frame before its root
  // element exists. Do not let that timing prevent the wheel listener below
  // from being registered.
  if (!document.documentElement) {
    document.addEventListener('DOMContentLoaded', startOutlookDomObserver, { once: true });
    return;
  }

  outlookDomObserver.observe(document.documentElement, {
    attributes: true,
    attributeFilter: [
      'aria-label',
      'data-automation-type',
      'dataautomationtype',
      'data-icon-name',
      'label',
      'role',
      'title'
    ],
    childList: true,
    subtree: true
  });
}

function isArchiveControlOrChild(target) {
  return target instanceof Element && Boolean(
    target.matches(ARCHIVE_CONTROL_SELECTOR) ||
    target.closest(ARCHIVE_CONTROL_SELECTOR)
  );
}

document.addEventListener('wheel', (event) => {
  if (!settings.enabled) return;

  if (!settings.nativePinch || !event.ctrlKey || event.metaKey) return;

  // This prevents Outlook's event handler from turning the pinch into its own
  // text/content zoom. It intentionally does *not* cancel the event's default
  // action, so Chrome performs the same zoom it uses on ordinary webpages.
  event.stopImmediatePropagation();
}, { capture: true, passive: false });

document.addEventListener('keydown', (event) => {
  if (
    !settings.enabled ||
    !settings.emailSizeControls ||
    !scaledEmailContent ||
    event.ctrlKey ||
    event.metaKey ||
    event.altKey ||
    isEditableTarget(event.target)
  ) return;

  const direction = getEmailScaleDirection(event);
  if (!direction) return;

  event.preventDefault();
  changeEmailScale(direction);
}, { capture: true });

startOutlookDomObserver();

function applyVisualSettings() {
  if (!document.documentElement) {
    document.addEventListener('DOMContentLoaded', applyVisualSettings, { once: true });
    return;
  }

  document.documentElement.dataset.enhancedOutlookArchiveButton = String(
    settings.enabled && settings.archiveButton
  );
  queueArchiveLabelUpdate();
  queueEmailSizeControlsUpdate();
}

function queueArchiveLabelUpdate() {
  if (archiveLabelUpdateQueued) return;
  archiveLabelUpdateQueued = true;
  requestAnimationFrame(() => {
    archiveLabelUpdateQueued = false;
    const showArchiveLabel = settings.enabled && settings.archiveButton;
    for (const control of document.querySelectorAll(ARCHIVE_CONTROL_SELECTOR)) {
      const injectedLabel = control.querySelector(`.${ARCHIVE_LABEL_CLASS}`);
      const visibleLabel = control.querySelector(
        ':scope > :not(.fui-Button__icon):not(.acui-hidden-content):not([aria-hidden="true"]):not(.enhanced-outlook-archive-label)'
      );
      const nativeText = visibleLabel?.textContent.trim() || '';

      if (showArchiveLabel && !nativeText && !injectedLabel) {
        const label = document.createElement('span');
        label.className = ARCHIVE_LABEL_CLASS;
        label.textContent = 'Archive';
        label.setAttribute('aria-hidden', 'true');
        control.append(label);
      } else if ((!showArchiveLabel || nativeText) && injectedLabel) {
        injectedLabel.remove();
      }
    }
  });
}

function queueEmailSizeControlsUpdate() {
  if (emailSizeControlsQueued) return;
  emailSizeControlsQueued = true;
  requestAnimationFrame(() => {
    emailSizeControlsQueued = false;
    const featureEnabled = settings.enabled && settings.emailSizeControls;
    const emailContent = findEmailContent();

    if (scaledEmailContent && scaledEmailContent !== emailContent) {
      clearEmailScale(scaledEmailContent);
      emailScale = 1;
    }

    scaledEmailContent = emailContent;
    if (featureEnabled && emailContent) {
      applyEmailScale(emailContent);
      if (window.top === window) ensureEmailSizeControls();
    } else {
      clearEmailScale(emailContent);
      removeEmailSizeControls();
    }
  });
}

function findEmailContent() {
  if (window.top !== window) return document.body;
  return document.querySelector('[role="document"]');
}

function applyEmailScale(element) {
  element.style.setProperty('--enhanced-outlook-email-scale', emailScale);
  element.dataset.enhancedOutlookEmailScaled = 'true';
  updateEmailSizeControls();
}

function clearEmailScale(element) {
  if (!element) return;
  element.style.removeProperty('--enhanced-outlook-email-scale');
  delete element.dataset.enhancedOutlookEmailScaled;
}

function ensureEmailSizeControls() {
  if (emailSizeControls) return;
  emailSizeControls = document.createElement('div');
  emailSizeControls.className = 'enhanced-outlook-email-size-controls';
  emailSizeControls.setAttribute('aria-label', 'Email size controls');
  emailSizeControls.innerHTML = [
    '<button type="button" data-email-scale="down" aria-label="Decrease email size">−</button>',
    '<span aria-live="polite"></span>',
    '<button type="button" data-email-scale="up" aria-label="Increase email size">+</button>'
  ].join('');
  emailSizeControls.addEventListener('click', (event) => {
    const direction = event.target.closest('button')?.dataset.emailScale;
    if (!direction) return;
    changeEmailScale(direction);
  });
  document.body.append(emailSizeControls);
  updateEmailSizeControls();
}

function updateEmailSizeControls() {
  if (!emailSizeControls) return;
  emailSizeControls.querySelector('[data-email-scale="down"]').disabled = emailScale <= EMAIL_SCALE_MIN;
  emailSizeControls.querySelector('[data-email-scale="up"]').disabled = emailScale >= EMAIL_SCALE_MAX;
  emailSizeControls.querySelector('span').textContent = `${Math.round(emailScale * 100)}%`;
}

function removeEmailSizeControls() {
  emailSizeControls?.remove();
  emailSizeControls = undefined;
}

function getEmailScaleDirection(event) {
  if (event.key === '=' || event.code === 'NumpadAdd') return 'up';
  if (event.key === '-' || event.code === 'NumpadSubtract') return 'down';
  return undefined;
}

function isEditableTarget(target) {
  return target instanceof Element && Boolean(target.closest(
    'input, textarea, select, [contenteditable="true"], [contenteditable="plaintext-only"]'
  ));
}

function changeEmailScale(direction) {
  emailScale = Math.min(
    EMAIL_SCALE_MAX,
    Math.max(EMAIL_SCALE_MIN, +(emailScale + (direction === 'up' ? EMAIL_SCALE_STEP : -EMAIL_SCALE_STEP)).toFixed(1))
  );
  if (scaledEmailContent) applyEmailScale(scaledEmailContent);
}
