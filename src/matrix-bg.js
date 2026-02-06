import * as THREE from 'three';
import { EffectComposer } from 'three/examples/jsm/postprocessing/EffectComposer.js';
import { RenderPass } from 'three/examples/jsm/postprocessing/RenderPass.js';
import { UnrealBloomPass } from 'three/examples/jsm/postprocessing/UnrealBloomPass.js';

const container = document.getElementById('matrix-bg');

if (container) {
  initMatrixBackground(container);
}

function initMatrixBackground(target) {
  const scene = new THREE.Scene();
  scene.fog = new THREE.FogExp2(0x001100, 0.025);

  const camera = new THREE.PerspectiveCamera(70, window.innerWidth / window.innerHeight, 0.1, 120);

  const renderer = new THREE.WebGLRenderer({ antialias: false, powerPreference: 'high-performance', alpha: true });
  const maxPixelRatio = 1.25;
  renderer.setSize(window.innerWidth, window.innerHeight);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, maxPixelRatio));
  renderer.toneMapping = THREE.ACESFilmicToneMapping;
  renderer.toneMappingExposure = 1.0;
  renderer.setClearColor(0x000000, 0);
  target.appendChild(renderer.domElement);

  const range = 40;
  const loopSpan = range * 4;
  const halfSpan = loopSpan / 2;

  const rainStart = 30;
  const rainRamp = 240;
  const rainCountRamp = 360;
  const baseRainCount = 1200;
  const maxRainCount = 3200;
  const rainCount = maxRainCount;
  const rainArea = { x: loopSpan, y: 50, z: loopSpan };
  const rainCeil = rainArea.y / 2;
  const rainSpeedMin = 3;
  const rainSpeedMax = 10;
  const rainSpanMin = 0.2;
  const rainSpanMax = 0.8;

  function createMatrixTexture() {
    const canvas = document.createElement('canvas');
    canvas.width = 256;
    canvas.height = 256;
    const ctx = canvas.getContext('2d');

    ctx.fillStyle = '#000000';
    ctx.fillRect(0, 0, 256, 256);

    const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const fontSize = 16;
    const columns = canvas.width / fontSize;

    for (let i = 0; i < columns; i++) {
      const x = i * fontSize;
      const drops = Math.random() * 30 + 10;

      for (let j = 0; j < drops; j++) {
        const text = chars.charAt(Math.floor(Math.random() * chars.length));
        const y = j * fontSize * 1.2;

        if (j > drops - 2) ctx.fillStyle = '#ccffcc';
        else if (j > drops - 6) ctx.fillStyle = '#00ff41';
        else ctx.fillStyle = '#003311';

        ctx.font = `bold ${fontSize}px monospace`;
        ctx.fillText(text, x, y);
      }
    }

    const texture = new THREE.CanvasTexture(canvas);
    texture.wrapS = THREE.RepeatWrapping;
    texture.wrapT = THREE.RepeatWrapping;
    return texture;
  }

  const matrixTexture = createMatrixTexture();

  const rainTextureCanvas = document.createElement('canvas');
  rainTextureCanvas.width = 48;
  rainTextureCanvas.height = 192;
  const rainTextureCtx = rainTextureCanvas.getContext('2d');
  const rainTexture = new THREE.CanvasTexture(rainTextureCanvas);
  rainTexture.minFilter = THREE.LinearFilter;
  rainTexture.magFilter = THREE.LinearFilter;
  rainTexture.generateMipmaps = false;

  function updateRainTexture() {
    rainTextureCtx.clearRect(0, 0, rainTextureCanvas.width, rainTextureCanvas.height);
    const fontSize = 16;
    rainTextureCtx.font = `bold ${fontSize}px monospace`;
    rainTextureCtx.textAlign = 'center';
    rainTextureCtx.textBaseline = 'top';
    for (let y = 0; y < rainTextureCanvas.height; y += fontSize) {
      const digit = Math.random() > 0.5 ? '1' : '0';
      rainTextureCtx.fillStyle = Math.random() > 0.9 ? '#ccffcc' : '#00ff41';
      rainTextureCtx.fillText(digit, rainTextureCanvas.width / 2, y);
    }
    rainTexture.needsUpdate = true;
  }

  updateRainTexture();

  const rainGeometry = new THREE.PlaneGeometry(0.35, 2.4);
  const rainMaterial = new THREE.MeshBasicMaterial({
    map: rainTexture,
    transparent: true,
    opacity: 0.9,
    depthWrite: false,
    depthTest: false,
    blending: THREE.AdditiveBlending,
    side: THREE.DoubleSide
  });
  const rainMesh = new THREE.InstancedMesh(rainGeometry, rainMaterial, rainCount);
  rainMesh.frustumCulled = false;
  rainMesh.renderOrder = 2;
  scene.add(rainMesh);

  const rainDummy = new THREE.Object3D();
  const rainDrops = [];
  for (let i = 0; i < rainCount; i++) {
    const x = (Math.random() - 0.5) * rainArea.x;
    const z = (Math.random() - 0.5) * rainArea.z;
    const speed = rainSpeedMin + Math.random() * (rainSpeedMax - rainSpeedMin);
    const span = rainArea.y * (rainSpanMin + Math.random() * (rainSpanMax - rainSpanMin));
    const phase = Math.random() * span;
    rainDrops.push({ x, z, speed, span, phase });
  }

  const castleGeometry = new THREE.BoxGeometry(4, 4, 0.2);
  const castleMaterial = new THREE.MeshStandardMaterial({
    color: 0x000000,
    roughness: 0.2,
    metalness: 0.8,
    emissive: 0x00ff41,
    emissiveMap: matrixTexture,
    emissiveIntensity: 1.5,
    map: matrixTexture,
    side: THREE.DoubleSide
  });

  const castleCount = 1800;
  const castleMesh = new THREE.InstancedMesh(castleGeometry, castleMaterial, castleCount);
  scene.add(castleMesh);

  const castleDummy = new THREE.Object3D();
  const castlePositions = [];
  const castleMatrices = [];

  for (let i = 0; i < castleCount; i++) {
    const x = (Math.floor(Math.random() * range) - range / 2) * 4;
    const y = (Math.floor(Math.random() * range) - range / 2) * 4;
    const z = (Math.floor(Math.random() * range) - range / 2) * 4;

    castleDummy.position.set(x, y, z);

    const rotType = Math.random();
    if (rotType < 0.33) castleDummy.rotation.x = Math.PI / 2;
    else if (rotType < 0.66) castleDummy.rotation.y = Math.PI / 2;
    else castleDummy.rotation.z = Math.PI / 2;

    const scale = 0.9 + Math.random() * 0.2;
    castleDummy.scale.set(scale, scale, 1);

    castleDummy.updateMatrix();
    castleMesh.setMatrixAt(i, castleDummy.matrix);

    const color = new THREE.Color();
    if (Math.random() > 0.1) color.setHex(0xffffff);
    else color.setHex(0x111111);
    castleMesh.setColorAt(i, color);

    castlePositions.push({ x, y, z });
    castleMatrices.push(castleDummy.matrix.clone());
  }
  castleMesh.instanceMatrix.needsUpdate = true;
  if (castleMesh.instanceColor) castleMesh.instanceColor.needsUpdate = true;

  const composer = new EffectComposer(renderer);
  composer.addPass(new RenderPass(scene, camera));

  const bloomPass = new UnrealBloomPass(new THREE.Vector2(window.innerWidth, window.innerHeight), 1.0, 0.35, 0.85);
  bloomPass.threshold = 0.1;
  bloomPass.strength = 0.7;
  bloomPass.radius = 0.3;
  composer.addPass(bloomPass);

  const mouse = new THREE.Vector2(0, 0);
  const guide = new THREE.Vector2(0, 0);
  const guideStrength = 6;
  const yawStrength = 0.35;
  const pitchStrength = 0.25;
  const baseSpeed = 0.15;
  const fastMultiplier = 2.5;
  let currentSpeed = baseSpeed;
  let targetSpeed = baseSpeed;

  window.addEventListener('mousedown', () => {
    targetSpeed = baseSpeed * fastMultiplier;
  });
  window.addEventListener('mouseup', () => {
    targetSpeed = baseSpeed;
  });
  window.addEventListener('mouseleave', () => {
    targetSpeed = baseSpeed;
  });
  window.addEventListener('mousemove', (event) => {
    mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
    mouse.y = -(event.clientY / window.innerHeight) * 2 + 1;
  });

  const clock = new THREE.Clock();
  let lastTime = 0;
  let cameraZ = 0;
  let rainTextureTimer = 0;

  function wrapZ(z) {
    return ((z + halfSpan) % loopSpan + loopSpan) % loopSpan - halfSpan;
  }

  function animate() {
    requestAnimationFrame(animate);
    const time = clock.getElapsedTime();
    const dt = time - lastTime;
    lastTime = time;

    matrixTexture.offset.y = time * 0.3;
    if (Math.random() > 0.95) matrixTexture.offset.x = Math.random() * 0.1;
    else matrixTexture.offset.x = 0;

    currentSpeed += (targetSpeed - currentSpeed) * 0.1;
    cameraZ -= currentSpeed;
    if (cameraZ < -loopSpan) cameraZ += loopSpan;

    guide.x += (mouse.x - guide.x) * 0.06;
    guide.y += (mouse.y - guide.y) * 0.06;
    const cameraX = guide.x * guideStrength;
    const cameraY = guide.y * guideStrength;

    const zShift = -cameraZ;

    for (let i = 0; i < castleCount; i++) {
      const p = castlePositions[i];
      const z = wrapZ(p.z + zShift);
      castleDummy.matrix.copy(castleMatrices[i]);
      castleDummy.matrix.setPosition(p.x, p.y, z);
      castleMesh.setMatrixAt(i, castleDummy.matrix);
    }
    castleMesh.instanceMatrix.needsUpdate = true;

    camera.position.set(cameraX, cameraY, 0);
    const targetYaw = -guide.x * yawStrength;
    const targetPitch = guide.y * pitchStrength;
    camera.rotation.x += (targetPitch - camera.rotation.x) * 0.1;
    camera.rotation.y += (targetYaw - camera.rotation.y) * 0.1;
    camera.rotation.z = Math.sin(time * 0.2) * 0.2;

    const rainLevel = Math.min(1, Math.max(0, (time - rainStart) / rainRamp));
    const rainCountLevel = Math.min(1, Math.max(0, (time - rainStart) / rainCountRamp));
    const smoothLevel = rainLevel * rainLevel;
    const smoothCountLevel = rainCountLevel * rainCountLevel;
    const activeRain = Math.floor(baseRainCount + (maxRainCount - baseRainCount) * smoothCountLevel);
    rainMesh.count = time < rainStart ? 0 : activeRain;
    const rainSpeedScale = 0.6 + smoothLevel * 1.4;
    rainMaterial.opacity = 0.2 + smoothLevel * 0.7;

    rainTextureTimer -= dt;
    if (rainTextureTimer <= 0) {
      updateRainTexture();
      rainTextureTimer = 0.6 + Math.random() * 0.4;
    }

    rainDummy.quaternion.copy(camera.quaternion);
    for (let i = 0; i < rainCount; i++) {
      const drop = rainDrops[i];
      const fall = (time * drop.speed * rainSpeedScale + drop.phase) % drop.span;
      const y = rainCeil - (fall / drop.span) * rainArea.y;
      const z = wrapZ(drop.z + zShift);
      rainDummy.position.set(drop.x, y, z);
      rainDummy.updateMatrix();
      rainMesh.setMatrixAt(i, rainDummy.matrix);
    }
    rainMesh.instanceMatrix.needsUpdate = true;

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
