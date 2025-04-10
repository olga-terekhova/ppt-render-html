window.addEventListener("DOMContentLoaded", () => {
  const canvas = document.getElementById("slideCanvas");
  const ctx = canvas.getContext("2d");

  const dpi = window.devicePixelRatio || 1;

  // Set canvas dimensions
  canvas.width = slideMetadata.width * dpi;
  canvas.height = slideMetadata.height * dpi;
  canvas.style.width = slideMetadata.width + "px";
  canvas.style.height = slideMetadata.height + "px";

  ctx.scale(dpi, dpi);

  // Load and draw background image
  const bgImage = new Image();
  bgImage.src = "slide_1.png";
  bgImage.onload = () => {
    ctx.drawImage(bgImage, 0, 0, slideMetadata.width, slideMetadata.height);
  };

  const tooltip = document.getElementById("tooltip");

  canvas.addEventListener("mousemove", (e) => {
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (slideMetadata.width / canvas.clientWidth);
    const y = (e.clientY - rect.top) * (slideMetadata.height / canvas.clientHeight);

    let found = false;

    for (let i = slideMetadata.links.length - 1; i >= 0; i--) {
      const link = slideMetadata.links[i];
      if (x >= link.x && x <= link.x + link.w && y >= link.y && y <= link.y + link.h) {
        tooltip.textContent = link.url;
        tooltip.style.left = `${e.pageX + 10}px`;
        tooltip.style.top = `${e.pageY + 10}px`;
        tooltip.style.display = "block";
        canvas.style.cursor = "pointer";
        found = true;
        break;
      }
    }

    if (!found) {
      tooltip.style.display = "none";
      canvas.style.cursor = "default";
    }
  });

  canvas.addEventListener("mouseleave", () => {
    tooltip.style.display = "none";
    canvas.style.cursor = "default";
  });

  canvas.addEventListener("click", (e) => {
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (slideMetadata.width / canvas.clientWidth);
    const y = (e.clientY - rect.top) * (slideMetadata.height / canvas.clientHeight);

    for (let i = slideMetadata.links.length - 1; i >= 0; i--) {
      const link = slideMetadata.links[i];
      if (x >= link.x && x <= link.x + link.w && y >= link.y && y <= link.y + link.h) {
        window.open(link.url, "_blank");
        break;
      }
    }
  });
});
