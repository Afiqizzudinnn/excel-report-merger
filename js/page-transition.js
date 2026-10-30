// Picks the slide direction when moving between the tools with the menu:
// going right in the menu slides forward, going left slides back.
// Loaded as a classic script in <head> so it is ready before the new page is first shown.
// Shared by every tool in the Office Toolkit, so keep this file identical across the repos.
(() => {
  const MENU_ORDER = ['excel-data-cleaner', 'excel-report-merger', 'automation-roi-calculator'];

  const menuIndex = (url) => MENU_ORDER.findIndex((name) => new URL(url).pathname.includes(`/${name}/`));

  window.addEventListener('pagereveal', (event) => {
    // A skipped transition (fast clicks, a hidden tab) rejects `ready`; that is expected, not an error.
    event.viewTransition?.ready.catch(() => {});
    const from = window.navigation?.activation?.from;
    if (!event.viewTransition?.types || !from) return;
    const direction = menuIndex(from.url) > menuIndex(window.location.href) ? 'back' : 'forward';
    event.viewTransition.types.add(direction);
  });
})();
