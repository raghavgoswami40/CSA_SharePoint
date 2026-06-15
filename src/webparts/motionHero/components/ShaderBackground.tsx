/**
 * ShaderBackground.tsx
 *
 * Three.js WebGL shader animation adapted for SPFx.
 * Removed: "use client", Tailwind classes.
 * Added: inline styles, SPFx-compatible lifecycle via useEffect/useRef.
 */

import * as React from 'react';
import * as THREE from 'three';

const vertexShader = `
  void main() {
    gl_Position = vec4( position, 1.0 );
  }
`;

const fragmentShader = `
  #define TWO_PI 6.2831853072
  #define PI 3.14159265359

  precision highp float;
  uniform vec2 resolution;
  uniform float time;

  void main(void) {
    vec2 uv = (gl_FragCoord.xy * 2.0 - resolution.xy) / min(resolution.x, resolution.y);
    float t = time * 0.05;
    float lineWidth = 0.002;

    vec3 color = vec3(0.0);
    for(int j = 0; j < 3; j++){
      for(int i = 0; i < 5; i++){
        color[j] += lineWidth * float(i * i) / abs(
          fract(t - 0.01 * float(j) + float(i) * 0.01) * 5.0
          - length(uv)
          + mod(uv.x + uv.y, 0.2)
        );
      }
    }

    gl_FragColor = vec4(color[0], color[1], color[2], 1.0);
  }
`;

interface ISceneRef {
  renderer: THREE.WebGLRenderer;
  uniforms: { time: { value: number }; resolution: { value: THREE.Vector2 } };
  geometry: THREE.PlaneGeometry;
  material: THREE.ShaderMaterial;
  animationId: number;
}

const ShaderBackground: React.FC = () => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const sceneRef     = React.useRef<ISceneRef | null>(null);

  React.useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    // Scene setup
    const camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
    const scene   = new THREE.Scene();
    const geometry = new THREE.PlaneGeometry(2, 2);

    const uniforms = {
      time:       { value: 1.0 },
      resolution: { value: new THREE.Vector2() },
    };

    const material = new THREE.ShaderMaterial({ uniforms, vertexShader, fragmentShader });
    scene.add(new THREE.Mesh(geometry, material));

    const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
    renderer.setPixelRatio(window.devicePixelRatio);
    container.appendChild(renderer.domElement);

    // Sizing
    const resize = (): void => {
      const w = container.clientWidth;
      const h = container.clientHeight;
      renderer.setSize(w, h);
      uniforms.resolution.value.set(renderer.domElement.width, renderer.domElement.height);
    };
    resize();
    window.addEventListener('resize', resize, false);

    // Animation loop
    const tick = (): void => {
      uniforms.time.value += 0.05;
      renderer.render(scene, camera);
      if (sceneRef.current) {
        sceneRef.current.animationId = requestAnimationFrame(tick);
      }
    };

    sceneRef.current = { renderer, uniforms, geometry, material, animationId: 0 };
    sceneRef.current.animationId = requestAnimationFrame(tick);

    return () => {
      window.removeEventListener('resize', resize);
      if (sceneRef.current) {
        cancelAnimationFrame(sceneRef.current.animationId);
        if (renderer.domElement.parentNode === container) {
          container.removeChild(renderer.domElement);
        }
        renderer.dispose();
        geometry.dispose();
        material.dispose();
      }
    };
  }, []);

  return (
    <div
      ref={containerRef}
      style={{
        position: 'absolute',
        inset: 0,
        width: '100%',
        height: '100%',
        overflow: 'hidden',
        background: '#000',
      }}
    />
  );
};

export default ShaderBackground;
