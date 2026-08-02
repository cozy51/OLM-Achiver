function drawIcon(size) {
  const canvas = new OffscreenCanvas(size, size);
  const context = canvas.getContext("2d");
  const scale = size / 128;
  const unit = value => value * scale;

  context.clearRect(0, 0, size, size);
  context.fillStyle = "#0078d4";
  context.beginPath();
  context.roundRect(unit(8), unit(12), unit(108), unit(104), unit(23));
  context.fill();

  context.fillStyle = "#1060b0";
  context.beginPath();
  context.moveTo(unit(13), unit(51));
  context.lineTo(unit(62), unit(83));
  context.lineTo(unit(111), unit(51));
  context.lineTo(unit(111), unit(104));
  context.lineTo(unit(13), unit(104));
  context.closePath();
  context.fill();

  context.strokeStyle = "white";
  context.lineWidth = unit(12);
  context.beginPath();
  context.ellipse(unit(42), unit(55), unit(18), unit(19), 0, 0, Math.PI * 2);
  context.stroke();

  context.fillStyle = "#ebf8ff";
  context.beginPath();
  context.roundRect(unit(64), unit(25), unit(45), unit(66), unit(5));
  context.fill();
  context.fillStyle = "#50beeb";
  context.beginPath();
  context.moveTo(unit(92), unit(25));
  context.lineTo(unit(109), unit(42));
  context.lineTo(unit(92), unit(42));
  context.closePath();
  context.fill();

  context.fillStyle = "#004987";
  context.beginPath();
  context.roundRect(unit(55), unit(76), unit(61), unit(29), unit(7));
  context.fill();
  if (size >= 32) {
    context.fillStyle = "white";
    context.font = `700 ${unit(16)}px Arial, sans-serif`;
    context.textAlign = "center";
    context.textBaseline = "middle";
    context.fillText("PDF", unit(85.5), unit(91));
  }
  return context.getImageData(0, 0, size, size);
}

export async function installActionIcon() {
  try {
    await chrome.action.setIcon({
      imageData: {
        16: drawIcon(16),
        32: drawIcon(32)
      }
    });
  } catch (error) {
    console.warn("Could not install the generated action icon.", error);
  }
}
