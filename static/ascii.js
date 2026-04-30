class AsciiEffect {
  constructor(video, canvas) {
    this.video = video;
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d');

    this.offCanvas = document.createElement('canvas');
    this.offCtx = this.offCanvas.getContext('2d', { willReadFrequently: true });

    // More spaces for sparsity, and fewer dense characters
    this.chars = ". : - = + * # % @ ";

    this.fontSize = 16;
    this.mouse = { x: -1000, y: -1000 };
    this.radius = 300; // Reveal radius

    window.addEventListener('resize', this.resize.bind(this));
    this.resize();

    document.addEventListener('mousemove', (e) => {
      this.mouse.x = e.clientX;
      this.mouse.y = e.clientY;
    });

    this.loop = this.loop.bind(this);
    requestAnimationFrame(this.loop);
  }

  resize() {
    const rect = this.video.parentElement.getBoundingClientRect();
    // Setup high-DPI canvas
    const dpr = window.devicePixelRatio || 1;
    this.canvas.width = rect.width * dpr;
    this.canvas.height = rect.height * dpr;
    this.canvas.style.width = `${rect.width}px`;
    this.canvas.style.height = `${rect.height}px`;

    this.ctx.scale(dpr, dpr);

    this.cols = Math.ceil(rect.width / this.fontSize);
    this.rows = Math.ceil(rect.height / this.fontSize);

    this.offCanvas.width = this.cols;
    this.offCanvas.height = this.rows;

    this.ctx.font = `400 ${this.fontSize}px "Geist Mono", monospace`;
    this.ctx.textBaseline = 'top';
  }

  loop() {
    if (this.video.readyState >= 2 && !this.video.paused) {
      this.offCtx.drawImage(this.video, 0, 0, this.cols, this.rows);
      const imgData = this.offCtx.getImageData(0, 0, this.cols, this.rows).data;

      this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
      const canvasRect = this.canvas.getBoundingClientRect();

      for (let y = 0; y < this.rows; y++) {
        for (let x = 0; x < this.cols; x++) {
          const i = (y * this.cols + x) * 4;
          const r = imgData[i];
          const g = imgData[i + 1];
          const b = imgData[i + 2];

          const brightness = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
          const charIndex = Math.floor(brightness * (this.chars.length - 1));
          const char = this.chars[charIndex];

          if (char !== ' ') {
            const charX = x * this.fontSize;
            const charY = y * this.fontSize;

            const screenX = canvasRect.left + charX;
            const screenY = canvasRect.top + charY;

            const dx = this.mouse.x - screenX;
            const dy = this.mouse.y - screenY;
            const dist = Math.sqrt(dx * dx + dy * dy);

            // Codex effect: very subtle white text
            let opacity = 0.05; // Base opacity very low
            if (dist < this.radius) {
              const intensity = 1 - (dist / this.radius);
              opacity += intensity * 0.4; // Max opacity near mouse
            }

            // Draw character using white (text-white-80) as seen in codex snippet
            this.ctx.fillStyle = `rgba(255, 255, 255, ${opacity})`;
            this.ctx.fillText(char, charX, charY);
          }
        }
      }
    }
    requestAnimationFrame(this.loop);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  const videos = document.querySelectorAll('.bg-video');
  videos.forEach(video => {
    const canvas = document.createElement('canvas');
    canvas.style.position = 'absolute';
    canvas.style.top = '0';
    canvas.style.left = '0';
    canvas.style.pointerEvents = 'none';
    canvas.setAttribute('aria-hidden', 'true');
    video.parentElement.appendChild(canvas);

    // Ensure video is playing
    video.play().catch(e => console.log('Video autoplay blocked'));

    // Initialize effect
    new AsciiEffect(video, canvas);
  });
});
