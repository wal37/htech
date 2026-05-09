document.querySelectorAll('form[action="#"]').forEach((form) => {
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const button = form.querySelector('button[type="submit"]');
    if (!button) return;

    const original = button.textContent;
    button.textContent = "Received";
    button.disabled = true;

    window.setTimeout(() => {
      button.textContent = original;
      button.disabled = false;
      form.reset();
    }, 1800);
  });
});
