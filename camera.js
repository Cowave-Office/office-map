/* One camera owns translation and scale; layout never retains an unscaled scroll area. */
window.OfficeMapCamera = class {
  constructor(viewport, content, onZoom) {
    this.viewport = viewport;
    this.content = content;
    this.onZoom = onZoom;
    this.zoom = 1;
    this.x = this.y = 0;
    this.width = this.height = 1;
    this.fitted = true;
    this.pointers = new Map();
    this.frame = 0;
    this.motion = null;
    this.rendered = null;
    this.reducedMotion = window.matchMedia?.('(prefers-reduced-motion: reduce)');
    this.ready = false;
    this.observer = new ResizeObserver(() => this.resize());
    this.observer.observe(viewport);
    viewport.addEventListener('pointerdown', event => {
      if (event.button !== 0 || event.target.closest('button,input,select,a')) return;
      this.stopMotion();
      viewport.setPointerCapture(event.pointerId);
      this.pointers.set(event.pointerId, { x: event.clientX, y: event.clientY });
      viewport.classList.add('grabbing-mode');
    });
    viewport.addEventListener('pointermove', event => {
      if (!this.pointers.has(event.pointerId)) return;
      const before = [...this.pointers.values()];
      const previous = this.pointers.get(event.pointerId);
      this.pointers.set(event.pointerId, { x: event.clientX, y: event.clientY });
      if (this.pointers.size === 1) {
        this.x += event.clientX - previous.x;
        this.y += event.clientY - previous.y;
        this.fitted = false;
        this.clamp();
        this.paint();
      } else if (this.pointers.size === 2) {
        const after = [...this.pointers.values()];
        const distance = points => Math.hypot(points[0].x - points[1].x, points[0].y - points[1].y);
        const center = points => ({ x: (points[0].x + points[1].x) / 2, y: (points[0].y + points[1].y) / 2 });
        const a = center(before), b = center(after);
        this.setZoom(this.zoom * distance(after) / Math.max(1, distance(before)), a.x, a.y, false);
        this.x += b.x - a.x;
        this.y += b.y - a.y;
        this.clamp();
        this.paint();
      }
    });
    const release = event => {
      this.pointers.delete(event.pointerId);
      if (!this.pointers.size) viewport.classList.remove('grabbing-mode');
    };
    ['pointerup', 'pointercancel', 'lostpointercapture'].forEach(type => viewport.addEventListener(type, release));
    viewport.addEventListener('dblclick', event => {
      if (event.target.closest('button,input,select,a')) return;
      event.preventDefault();
      this.setZoom(this.zoom + 0.4, event.clientX, event.clientY);
    });
    viewport.addEventListener('wheel', event => {
      event.preventDefault();
      this.stopMotion();
      if (event.ctrlKey || event.metaKey) {
        this.setZoom(this.zoom + (event.deltaY < 0 ? 0.08 : -0.08), event.clientX, event.clientY, false);
      } else {
        const unit = event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? viewport.clientHeight : 1;
        this.x -= (event.shiftKey ? event.deltaY : event.deltaX) * unit;
        this.y -= (event.shiftKey ? 0 : event.deltaY) * unit;
        this.fitted = false;
        this.clamp();
        this.paint();
      }
    }, { passive: false });
  }

  get minZoom() {
    return Math.min((this.viewport.clientWidth - 16) / this.width, (this.viewport.clientHeight - 16) / this.height, 2);
  }

  measure() {
    this.width = this.content.offsetWidth;
    this.height = this.content.offsetHeight;
    this.ready = this.width > 0 && this.height > 0;
    this.viewWidth = this.viewport.clientWidth;
    this.viewHeight = this.viewport.clientHeight;
  }

  fit() {
    this.stopMotion();
    this.measure();
    if (!this.ready) return;
    this.fitted = true;
    this.zoom = this.minZoom;
    this.x = (this.viewport.clientWidth - this.width * this.zoom) / 2;
    this.y = (this.viewport.clientHeight - this.height * this.zoom) / 2;
    this.paint();
  }

  resize() {
    if (!this.ready) return;
    this.stopMotion();
    if (this.fitted) return this.fit();
    const centerX = (this.viewWidth / 2 - this.x) / this.zoom;
    const centerY = (this.viewHeight / 2 - this.y) / this.zoom;
    this.measure();
    this.zoom = Math.max(this.minZoom, this.zoom);
    this.x = this.viewWidth / 2 - centerX * this.zoom;
    this.y = this.viewHeight / 2 - centerY * this.zoom;
    this.clamp();
    this.paint();
  }

  setZoom(value, clientX, clientY, animate = true) {
    if (!this.ready) return;
    this.stopMotion();
    const rect = this.viewport.getBoundingClientRect();
    const px = clientX === undefined ? rect.width / 2 : clientX - rect.left;
    const py = clientY === undefined ? rect.height / 2 : clientY - rect.top;
    const anchorX = (px - this.x) / this.zoom;
    const anchorY = (py - this.y) / this.zoom;
    this.zoom = Math.min(2, Math.max(this.minZoom, value));
    this.fitted = Math.abs(this.zoom - this.minZoom) < 0.00001;
    this.x = px - anchorX * this.zoom;
    this.y = py - anchorY * this.zoom;
    this.clamp();
    this.paint(animate);
  }

  clamp() {
    const w = this.width * this.zoom, h = this.height * this.zoom;
    this.x = w <= this.viewport.clientWidth ? (this.viewport.clientWidth - w) / 2 : Math.min(0, Math.max(this.viewport.clientWidth - w, this.x));
    this.y = h <= this.viewport.clientHeight ? (this.viewport.clientHeight - h) / 2 : Math.min(0, Math.max(this.viewport.clientHeight - h, this.y));
  }

  focus(element) {
    this.stopMotion();
    const target = element.getBoundingClientRect(), map = this.content.getBoundingClientRect();
    // A tab change can schedule fit before its transform reaches the next frame.
    const renderedScale = map.width / this.content.offsetWidth || this.zoom;
    const centerX = (target.left + target.width / 2 - map.left) / renderedScale;
    const centerY = (target.top + target.height / 2 - map.top) / renderedScale;
    this.zoom = Math.max(this.zoom, Math.min(1, Math.max(this.minZoom, 0.8)));
    this.fitted = false;
    this.x = this.viewport.clientWidth / 2 - centerX * this.zoom;
    this.y = this.viewport.clientHeight / 2 - centerY * this.zoom;
    this.clamp();
    this.paint();
  }

  stopMotion() {
    // Continue direct gestures from what is visible, never from a pending destination.
    if (this.motion && this.rendered) {
      Object.assign(this, this.rendered);
      this.fitted = Math.abs(this.zoom - this.minZoom) < 0.00001;
    }
    this.motion = null;
  }

  paint(animate = false) {
    this.motion = animate && this.rendered && !this.reducedMotion?.matches
      ? { from: { ...this.rendered }, to: { x: this.x, y: this.y, zoom: this.zoom }, start: performance.now() }
      : null;
    this.schedulePaint();
  }

  schedulePaint() {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      let view = { x: this.x, y: this.y, zoom: this.zoom };
      if (this.motion) {
        const { from, to, start } = this.motion;
        const progress = Math.min(1, Math.max(0, (performance.now() - start) / 260));
        const eased = 1 - (1 - progress) ** 3;
        // Interpolate translation and scale together to keep the zoom anchor stationary.
        view = Object.fromEntries(['x', 'y', 'zoom'].map(key => [key, from[key] + (to[key] - from[key]) * eased]));
        if (progress === 1) this.motion = null;
      }
      this.rendered = view;
      this.content.style.transform = `translate(${view.x}px, ${view.y}px) scale(${view.zoom})`;
      this.content.style.setProperty('--border-compensation', 1 / Math.min(1, view.zoom));
      this.onZoom(view.zoom);
      if (this.motion) this.schedulePaint();
    });
  }
};
