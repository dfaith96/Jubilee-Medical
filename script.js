const launchDate = new Date("2026-08-01T09:00:00+01:00");
const countdownEl = document.querySelector("#countdown");
const yearEl = document.querySelector("#year");
const navLinks = document.querySelector("[data-nav-links]");
const menuToggle = document.querySelector(".menu-toggle");
const modal = document.querySelector("[data-modal]");
const toast = document.querySelector("[data-toast]");

function plural(value, label) {
  return `${value} ${label}${value === 1 ? "" : "s"}`;
}

function updateCountdown() {
  if (!countdownEl) return;

  const distance = launchDate.getTime() - Date.now();
  if (distance <= 0) {
    countdownEl.textContent = "Beta access is open";
    return;
  }

  const days = Math.floor(distance / 86400000);
  const hours = Math.floor((distance % 86400000) / 3600000);
  const minutes = Math.floor((distance % 3600000) / 60000);
  const seconds = Math.floor((distance % 60000) / 1000);

  countdownEl.textContent = `${plural(days, "day")} ${String(hours).padStart(2, "0")}h ${String(
    minutes,
  ).padStart(2, "0")}m ${String(seconds).padStart(2, "0")}s`;
}

function showToast(message) {
  if (!toast) return;

  toast.textContent = message;
  toast.hidden = false;
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => {
    toast.hidden = true;
  }, 3600);
}

function openModal() {
  if (!modal) return;

  modal.hidden = false;
  document.body.style.overflow = "hidden";
  const firstInput = modal.querySelector("input, select, button");
  firstInput?.focus();
}

function closeModal() {
  if (!modal) return;

  modal.hidden = true;
  document.body.style.overflow = "";
}

function toggleMenu() {
  if (!navLinks || !menuToggle) return;

  const isOpen = navLinks.classList.toggle("is-open");
  menuToggle.setAttribute("aria-expanded", String(isOpen));
}

function closeMenu() {
  if (!navLinks || !menuToggle) return;

  navLinks.classList.remove("is-open");
  menuToggle.setAttribute("aria-expanded", "false");
}

function initJourneyTabs() {
  const tabs = document.querySelectorAll("[data-role-tab]");
  const panels = document.querySelectorAll("[data-journey]");

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      const role = tab.dataset.roleTab;

      tabs.forEach((item) => {
        const isActive = item === tab;
        item.classList.toggle("active", isActive);
        item.setAttribute("aria-selected", String(isActive));
      });

      panels.forEach((panel) => {
        panel.classList.toggle("hidden", panel.dataset.journey !== role);
      });
    });
  });
}

function initCareSearch() {
  const input = document.querySelector("#careSearch");
  const items = document.querySelectorAll("[data-tags]");

  if (!input || !items.length) return;

  input.addEventListener("input", () => {
    const query = input.value.trim().toLowerCase();

    items.forEach((item) => {
      const haystack = `${item.dataset.tags} ${item.textContent}`.toLowerCase();
      item.classList.toggle("is-hidden", query !== "" && !haystack.includes(query));
    });
  });
}

function initForms() {
  document.querySelectorAll("form[data-form]").forEach((form) => {
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const type = form.dataset.form;
      const messages = {
        waitlist: "You are on the Jubilee Med beta waitlist.",
        contact: "Your Jubilee Med contact request has been prepared.",
        demo: "Your Jubilee Med demo request has been prepared.",
      };

      showToast(messages[type] || "Request received.");
      form.reset();

      if (type === "demo") {
        closeModal();
      }
    });
  });
}

function initModal() {
  document.querySelectorAll("[data-open-demo]").forEach((button) => {
    button.addEventListener("click", openModal);
  });

  document.querySelectorAll("[data-close-modal]").forEach((button) => {
    button.addEventListener("click", closeModal);
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeModal();
  });
}

menuToggle?.addEventListener("click", toggleMenu);

document.querySelectorAll(".nav-links a").forEach((link) => {
  link.addEventListener("click", closeMenu);
});

yearEl.textContent = String(new Date().getFullYear());
updateCountdown();
window.setInterval(updateCountdown, 1000);
initJourneyTabs();
initCareSearch();
initForms();
initModal();
