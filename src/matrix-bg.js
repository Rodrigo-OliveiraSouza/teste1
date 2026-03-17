import * as THREE from 'three';

const container = document.getElementById('matrix-bg');

if (container) {
  initAmbientBackground(container);
}

function initAmbientBackground(target) {
  const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const compactViewport = window.matchMedia('(max-width: 900px)').matches;

  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(50, window.innerWidth / window.innerHeight, 0.1, 100);
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
    const gradient = ctx.createRadialGradient(128, 128, 8, 128, 128, 128);

    gradient.addColorStop(0, 'rgba(255,255,255,0.95)');
    gradient.addColorStop(0.35, 'rgba(255,255,255,0.26)');
    gradient.addColorStop(1, 'rgba(255,255,255,0)');

    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, 256, 256);

    return new THREE.CanvasTexture(canvas);
  }

  const glowTexture = createGlowTexture();
  const glowCount = compactViewport ? 8 : 12;
  const glows = [];

  for (let i = 0; i < glowCount; i++) {
    const material = new THREE.SpriteMaterial({
      map: glowTexture,
      color: i % 4 === 0 ? 0xffb876 : 0x76f0d0,
      transparent: true,
      opacity: i % 4 === 0 ? 0.11 : 0.14,
      depthWrite: false,
      blending: THREE.AdditiveBlending
    });

    const sprite = new THREE.Sprite(material);
    const scale = compactViewport ? 5 + Math.random() * 5 : 7 + Math.random() * 8;
    sprite.scale.set(scale, scale, 1);
    scene.add(sprite);

    glows.push({
      sprite,
      baseX: (Math.random() - 0.5) * (compactViewport ? 20 : 28),
      baseY: (Math.random() - 0.5) * (compactViewport ? 14 : 18),
      baseZ: -8 - Math.random() * 18,
      drift: Math.random() * Math.PI * 2,
      ampX: 0.8 + Math.random() * 1.2,
      ampY: 0.6 + Math.random() * 1,
      speed: 0.08 + Math.random() * 0.08
    });
  }

  const mouse = new THREE.Vector2(0, 0);
  const guide = new THREE.Vector2(0, 0);

  window.addEventListener('mousemove', (event) => {
    mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
    mouse.y = -(event.clientY / window.innerHeight) * 2 + 1;
  });

  const clock = new THREE.Clock();

  function animate() {
    requestAnimationFrame(animate);

    const time = clock.getElapsedTime();
    guide.x += (mouse.x - guide.x) * 0.025;
    guide.y += (mouse.y - guide.y) * 0.025;

    for (const glow of glows) {
      const speed = prefersReducedMotion ? glow.speed * 0.15 : glow.speed;
      glow.sprite.position.set(
        glow.baseX + Math.sin(time * speed + glow.drift) * glow.ampX,
        glow.baseY + Math.cos(time * speed * 0.8 + glow.drift) * glow.ampY,
        glow.baseZ
      );
    }

    camera.position.x += (guide.x * 1.1 - camera.position.x) * 0.03;
    camera.position.y += (guide.y * 0.9 - camera.position.y) * 0.03;

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
