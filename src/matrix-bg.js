import * as THREE from 'three';

const container = document.getElementById('matrix-bg');

if (container) {
  initAmbientBackground(container);
}

function initAmbientBackground(target) {
  const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const compactViewport = window.matchMedia('(max-width: 900px)').matches;

  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(48, window.innerWidth / window.innerHeight, 0.1, 100);
  camera.position.z = 24;

  const renderer = new THREE.WebGLRenderer({ alpha: true, antialias: true, powerPreference: 'high-performance' });
  const maxPixelRatio = 1.25;
  renderer.setSize(window.innerWidth, window.innerHeight);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, maxPixelRatio));
  renderer.setClearColor(0x000000, 0);
  target.appendChild(renderer.domElement);

  function createGlowTexture() {
    const canvas = document.createElement('canvas');
    canvas.width = 256;
    canvas.height = 256;
    const ctx = canvas.getContext('2d');
    const gradient = ctx.createRadialGradient(128, 128, 10, 128, 128, 128);

    gradient.addColorStop(0, 'rgba(255,255,255,0.95)');
    gradient.addColorStop(0.4, 'rgba(255,255,255,0.16)');
    gradient.addColorStop(1, 'rgba(255,255,255,0)');

    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, 256, 256);

    return new THREE.CanvasTexture(canvas);
  }

  function createBeamTexture() {
    const canvas = document.createElement('canvas');
    canvas.width = 128;
    canvas.height = 512;
    const ctx = canvas.getContext('2d');
    const gradient = ctx.createLinearGradient(0, 0, canvas.width, 0);

    gradient.addColorStop(0, 'rgba(255,255,255,0)');
    gradient.addColorStop(0.5, 'rgba(255,255,255,1)');
    gradient.addColorStop(1, 'rgba(255,255,255,0)');

    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    const fade = ctx.createLinearGradient(0, 0, 0, canvas.height);
    fade.addColorStop(0, 'rgba(255,255,255,0)');
    fade.addColorStop(0.18, 'rgba(255,255,255,0.55)');
    fade.addColorStop(0.82, 'rgba(255,255,255,0.55)');
    fade.addColorStop(1, 'rgba(255,255,255,0)');

    ctx.globalCompositeOperation = 'destination-in';
    ctx.fillStyle = fade;
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    return new THREE.CanvasTexture(canvas);
  }

  const glowTexture = createGlowTexture();
  const beamTexture = createBeamTexture();
  const glows = [];
  const beams = [];

  const glowCount = compactViewport ? 5 : 7;
  for (let i = 0; i < glowCount; i++) {
    const material = new THREE.SpriteMaterial({
      map: glowTexture,
      color: i % 3 === 0 ? 0x7abfff : 0x78f0d3,
      transparent: true,
      opacity: i % 3 === 0 ? 0.055 : 0.07,
      depthWrite: false,
      blending: THREE.AdditiveBlending
    });

    const sprite = new THREE.Sprite(material);
    const scale = compactViewport ? 7 + Math.random() * 4 : 10 + Math.random() * 8;
    sprite.scale.set(scale, scale, 1);
    scene.add(sprite);

    glows.push({
      sprite,
      baseX: (Math.random() - 0.5) * (compactViewport ? 18 : 26),
      baseY: (Math.random() - 0.5) * (compactViewport ? 12 : 16),
      baseZ: -8 - Math.random() * 16,
      drift: Math.random() * Math.PI * 2,
      ampX: 0.4 + Math.random() * 0.8,
      ampY: 0.4 + Math.random() * 0.7,
      speed: 0.05 + Math.random() * 0.05
    });
  }

  const beamCount = compactViewport ? 4 : 6;
  for (let i = 0; i < beamCount; i++) {
    const material = new THREE.SpriteMaterial({
      map: beamTexture,
      color: i === beamCount - 1 ? 0xffc58b : 0x8cdcff,
      transparent: true,
      opacity: i === beamCount - 1 ? 0.05 : 0.07,
      depthWrite: false,
      blending: THREE.AdditiveBlending
    });

    const sprite = new THREE.Sprite(material);
    sprite.scale.set(compactViewport ? 4 + Math.random() * 1.5 : 4.5 + Math.random() * 2, compactViewport ? 20 + Math.random() * 6 : 24 + Math.random() * 10, 1);
    scene.add(sprite);

    beams.push({
      sprite,
      baseX: (Math.random() - 0.5) * (compactViewport ? 16 : 24),
      baseY: (Math.random() - 0.5) * 4,
      baseZ: -6 - Math.random() * 18,
      drift: Math.random() * Math.PI * 2,
      speed: 0.03 + Math.random() * 0.03,
      ampX: 0.35 + Math.random() * 0.45
    });
  }

  const mouse = new THREE.Vector2(0, 0);
  const guide = new THREE.Vector2(0, 0);
  const clock = new THREE.Clock();

  window.addEventListener('mousemove', (event) => {
    mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
    mouse.y = -(event.clientY / window.innerHeight) * 2 + 1;
  });

  function animate() {
    requestAnimationFrame(animate);

    const time = clock.getElapsedTime();
    guide.x += (mouse.x - guide.x) * 0.02;
    guide.y += (mouse.y - guide.y) * 0.02;

    for (const glow of glows) {
      const speed = prefersReducedMotion ? glow.speed * 0.15 : glow.speed;
      glow.sprite.position.set(
        glow.baseX + Math.sin(time * speed + glow.drift) * glow.ampX,
        glow.baseY + Math.cos(time * speed * 0.8 + glow.drift) * glow.ampY,
        glow.baseZ
      );
    }

    for (const beam of beams) {
      const speed = prefersReducedMotion ? beam.speed * 0.15 : beam.speed;
      beam.sprite.position.set(
        beam.baseX + Math.sin(time * speed + beam.drift) * beam.ampX,
        beam.baseY,
        beam.baseZ
      );
    }

    camera.position.x += (guide.x * 0.8 - camera.position.x) * 0.025;
    camera.position.y += (guide.y * 0.55 - camera.position.y) * 0.025;

    renderer.render(scene, camera);
  }

  window.addEventListener('resize', () => {
    camera.aspect = window.innerWidth / window.innerHeight;
    camera.updateProjectionMatrix();
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, maxPixelRatio));
    renderer.setSize(window.innerWidth, window.innerHeight);
  });

  animate();
}
