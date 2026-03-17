import * as THREE from 'three';
import { EffectComposer } from 'three/examples/jsm/postprocessing/EffectComposer.js';
import { RenderPass } from 'three/examples/jsm/postprocessing/RenderPass.js';
import { UnrealBloomPass } from 'three/examples/jsm/postprocessing/UnrealBloomPass.js';

const container = document.getElementById('matrix-bg');

if (container) {
  initMatrixBackground(container);
}

function initMatrixBackground(target) {
  const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const compactViewport = window.matchMedia('(max-width: 900px)').matches;

  const scene = new THREE.Scene();
  scene.fog = new THREE.FogExp2(0x05080d, compactViewport ? 0.06 : 0.045);

  const camera = new THREE.PerspectiveCamera(58, window.innerWidth / window.innerHeight, 0.1, 120);
  const renderer = new THREE.WebGLRenderer({ antialias: false, powerPreference: 'high-performance', alpha: true });
  const maxPixelRatio = 1.25;

  renderer.setSize(window.innerWidth, window.innerHeight);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, maxPixelRatio));
  renderer.toneMapping = THREE.ACESFilmicToneMapping;
  renderer.toneMappingExposure = 0.72;
  renderer.setClearColor(0x000000, 0);
  target.appendChild(renderer.domElement);

  const ambientLight = new THREE.AmbientLight(0x9ab8ac, 0.28);
  const keyLight = new THREE.DirectionalLight(0x7ed8b5, 0.42);
  keyLight.position.set(8, 10, 12);
  scene.add(ambientLight, keyLight);

  function createMatrixTexture() {
    const canvas = document.createElement('canvas');
    canvas.width = 256;
    canvas.height = 256;
    const ctx = canvas.getContext('2d');

    ctx.fillStyle = '#04080b';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const fontSize = 16;
    const columns = canvas.width / fontSize;

    for (let i = 0; i < columns; i++) {
      const x = i * fontSize;
      const drops = Math.floor(Math.random() * 12) + 8;

      for (let j = 0; j < drops; j++) {
        const text = chars.charAt(Math.floor(Math.random() * chars.length));
        const y = j * fontSize * 1.25;

        if (j > drops - 2) ctx.fillStyle = '#97c9b2';
        else if (j > drops - 5) ctx.fillStyle = '#2c6d57';
        else ctx.fillStyle = '#0b1812';

        ctx.font = `600 ${fontSize}px monospace`;
        ctx.fillText(text, x, y);
      }
    }

    const texture = new THREE.CanvasTexture(canvas);
    texture.wrapS = THREE.RepeatWrapping;
    texture.wrapT = THREE.RepeatWrapping;
    return texture;
  }

  const matrixTexture = createMatrixTexture();
  const panelGeometry = new THREE.PlaneGeometry(3.6, 3.6);
  const panelMaterial = new THREE.MeshStandardMaterial({
    color: 0x05080d,
    roughness: 0.7,
    metalness: 0.18,
    emissive: 0x295f4b,
    emissiveMap: matrixTexture,
    emissiveIntensity: 0.22,
    map: matrixTexture,
    transparent: true,
    opacity: 0.62,
    side: THREE.DoubleSide
  });

  const panelCount = compactViewport ? 48 : 96;
  const depthRange = compactViewport ? 44 : 56;
  const loopSpan = depthRange * 2;
  const halfSpan = loopSpan / 2;
  const panelMesh = new THREE.InstancedMesh(panelGeometry, panelMaterial, panelCount);
  panelMesh.frustumCulled = false;
  scene.add(panelMesh);

  const panelDummy = new THREE.Object3D();
  const panels = [];

  for (let i = 0; i < panelCount; i++) {
    panels.push({
      x: (Math.random() - 0.5) * (compactViewport ? 26 : 38),
      y: (Math.random() - 0.5) * (compactViewport ? 18 : 24),
      z: -Math.random() * loopSpan,
      rotX: (Math.random() - 0.5) * 0.6,
      rotY: (Math.random() - 0.5) * 0.45,
      rotZ: (Math.random() - 0.5) * 0.15,
      scale: 0.8 + Math.random() * 1.15,
      drift: Math.random() * Math.PI * 2,
      speed: 0.018 + Math.random() * 0.028
    });
  }

  const composer = new EffectComposer(renderer);
  composer.addPass(new RenderPass(scene, camera));

  const bloomPass = new UnrealBloomPass(new THREE.Vector2(window.innerWidth, window.innerHeight), 0.22, 0.18, 0.96);
  bloomPass.threshold = 0.42;
  bloomPass.strength = 0.14;
  bloomPass.radius = 0.1;
  composer.addPass(bloomPass);

  const mouse = new THREE.Vector2(0, 0);
  const guide = new THREE.Vector2(0, 0);
  const guideStrength = compactViewport ? 1.2 : 2.1;
  const yawStrength = 0.045;
  const pitchStrength = 0.03;
  const baseSpeed = prefersReducedMotion ? 0.006 : 0.02;
  let cameraZ = 0;

  window.addEventListener('mousemove', (event) => {
    mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
    mouse.y = -(event.clientY / window.innerHeight) * 2 + 1;
  });

  const clock = new THREE.Clock();

  function wrapZ(z) {
    return ((z + halfSpan) % loopSpan + loopSpan) % loopSpan - halfSpan;
  }

  function animate() {
    requestAnimationFrame(animate);

    const time = clock.getElapsedTime();
    guide.x += (mouse.x - guide.x) * 0.03;
    guide.y += (mouse.y - guide.y) * 0.03;

    matrixTexture.offset.y = time * 0.025;
    cameraZ -= baseSpeed;
    if (cameraZ < -loopSpan) {
      cameraZ += loopSpan;
    }

    const zShift = -cameraZ;

    for (let i = 0; i < panelCount; i++) {
      const panel = panels[i];
      const swayX = Math.sin(time * 0.18 + panel.drift) * 0.35;
      const swayY = Math.cos(time * 0.15 + panel.drift) * 0.28;
      const z = wrapZ(panel.z + zShift * (0.65 + panel.speed));

      panelDummy.position.set(panel.x + swayX, panel.y + swayY, z);
      panelDummy.rotation.set(
        panel.rotX + Math.sin(time * 0.12 + panel.drift) * 0.04,
        panel.rotY + Math.cos(time * 0.11 + panel.drift) * 0.05,
        panel.rotZ
      );
      panelDummy.scale.set(panel.scale, panel.scale * 1.18, 1);
      panelDummy.updateMatrix();
      panelMesh.setMatrixAt(i, panelDummy.matrix);
    }

    panelMesh.instanceMatrix.needsUpdate = true;

    camera.position.set(guide.x * guideStrength, guide.y * guideStrength, 0);
    camera.rotation.x += (guide.y * pitchStrength - camera.rotation.x) * 0.05;
    camera.rotation.y += (-guide.x * yawStrength - camera.rotation.y) * 0.05;
    camera.rotation.z = Math.sin(time * 0.08) * 0.012;

    composer.render();
  }

  window.addEventListener('resize', () => {
    camera.aspect = window.innerWidth / window.innerHeight;
    camera.updateProjectionMatrix();
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, maxPixelRatio));
    renderer.setSize(window.innerWidth, window.innerHeight);
    composer.setSize(window.innerWidth, window.innerHeight);
  });

  animate();
}
