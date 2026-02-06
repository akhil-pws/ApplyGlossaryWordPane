export function initModeDropdownSingleDouble() {
  let clickCount = 0;
  let clickTimer: number | null = null;

  // ✅ handles single/double click reliably
  document.addEventListener("click", (e) => {
    const target = e.target as HTMLElement;

    // Only handle modeDropdown clicks
    const icon = target.closest("#modeDropdown") as HTMLElement | null;
    if (!icon) return;

    e.preventDefault();
    e.stopPropagation();

    clickCount++;

    if (clickCount === 1) {
      clickTimer = window.setTimeout(() => {
        clickCount = 0;

        // ✅ SINGLE CLICK ACTION (your existing thing)
        console.log("SINGLE CLICK");
        // yourExistingIconClick();

      }, 250);
    } else {
      if (clickTimer) clearTimeout(clickTimer);
      clickCount = 0;

      console.log("DOUBLE CLICK - OPEN DROPDOWN");

      // ✅ SHOW dropdown
      // @ts-ignore
      const dd = bootstrap.Dropdown.getOrCreateInstance(icon, { autoClose: true });
      dd.show();
    }
  });

  // ✅ Close dropdown when clicking outside
  document.addEventListener("click", (e) => {
    const icon = document.getElementById("modeDropdown") as HTMLElement | null;
    if (!icon) return;

    // If click is outside icon & outside menu -> hide
    const menu = icon.parentElement?.querySelector(".dropdown-menu") as HTMLElement | null;
    const t = e.target as HTMLElement;

    if (menu && !icon.contains(t) && !menu.contains(t)) {
      // @ts-ignore
      bootstrap.Dropdown.getOrCreateInstance(icon).hide();
    }
  });

  // ✅ Close dropdown after selecting menu item
  document.addEventListener("click", (e) => {
    const item = (e.target as HTMLElement).closest(".dropdown-menu .dropdown-item") as HTMLElement | null;
    if (!item) return;

    const icon = document.getElementById("modeDropdown") as HTMLElement | null;
    if (!icon) return;

    // @ts-ignore
    bootstrap.Dropdown.getOrCreateInstance(icon).hide();
  });
}