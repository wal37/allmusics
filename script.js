const header = document.querySelector(".site-header");
const toggle = document.querySelector(".menu-toggle");

if (header && toggle) {
  toggle.addEventListener("click", () => {
    const isOpen = header.classList.toggle("nav-open");
    toggle.setAttribute("aria-expanded", String(isOpen));
  });
}

document.querySelectorAll(".site-nav a").forEach((link) => {
  link.addEventListener("click", () => {
    if (header && header.classList.contains("nav-open")) {
      header.classList.remove("nav-open");
      if (toggle) {
        toggle.setAttribute("aria-expanded", "false");
      }
    }
  });
});
