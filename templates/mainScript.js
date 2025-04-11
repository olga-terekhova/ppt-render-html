window.addEventListener("DOMContentLoaded", () => {
  const canvas = document.getElementById("slideCanvas");
  const ctx = canvas.getContext("2d");
  const tooltip = document.getElementById("tooltip");

  const dpi = window.devicePixelRatio || 1;

  const originalWidth = slideMetadata.width;
  const originalHeight = slideMetadata.height;
  const originalRatio = originalWidth / originalHeight;

  // Function to calculate canvas size based on window dimensions
  function updateCanvasSize() {
    const windowWidth = window.innerWidth;
    const windowHeight = window.innerHeight;
    const windowRatio = windowWidth / windowHeight;

    let renderedWidth, renderedHeight;

    if (originalRatio >= windowRatio) {
      // Fit to width
      renderedWidth = windowWidth;
      renderedHeight = windowWidth / originalRatio;
    } else {
      // Fit to height
      renderedHeight = windowHeight;
      renderedWidth = windowHeight * originalRatio;
    }

    // DPI-scaled internal canvas resolution
    canvas.width = renderedWidth * dpi;
    canvas.height = renderedHeight * dpi;

    // CSS dimensions (logical)
    canvas.style.width = `${renderedWidth}px`;
    canvas.style.height = `${renderedHeight}px`;

    ctx.scale(dpi, dpi);

    // Redraw the background image
    const bgImage = new Image();
    bgImage.src = "slide_1.png";
    bgImage.onload = () => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(bgImage, 0, 0, renderedWidth, renderedHeight);
    };
  }

  // Initial canvas size update
  updateCanvasSize();

  // Resize event listener to update canvas when window size changes
  window.addEventListener("resize", updateCanvasSize);

  // Helper for scaling mouse coordinates
  function getRelativeCoords(e) {
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (originalWidth / canvas.style.width.replace("px", ""));
    const y = (e.clientY - rect.top) * (originalHeight / canvas.style.height.replace("px", ""));
    return { x, y };
  }

  // Mousemove listener to display tooltip and handle link clicks
  canvas.addEventListener("mousemove", (e) => {
    const { x, y } = getRelativeCoords(e);
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

  // Mouseleave listener to hide tooltip
  canvas.addEventListener("mouseleave", () => {
    tooltip.style.display = "none";
    canvas.style.cursor = "default";
  });

  // Click listener to handle opening links
  canvas.addEventListener("click", (e) => {
    const { x, y } = getRelativeCoords(e);
    for (let i = slideMetadata.links.length - 1; i >= 0; i--) {
      const link = slideMetadata.links[i];
      if (x >= link.x && x <= link.x + link.w && y >= link.y && y <= link.y + link.h) {
        window.open(link.url, "_blank");
        break;
      }
    }
  });
});
