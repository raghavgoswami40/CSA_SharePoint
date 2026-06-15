/**
 * Minimal ambient type declarations for the Three.js subset used in ShaderBackground.
 * Avoids @types/three which requires TypeScript ≥5.4 (SPFx uses 4.7).
 */
declare module 'three' {
  export class Vector2 {
    constructor(x?: number, y?: number);
    set(x: number, y: number): this;
  }

  export class OrthographicCamera {
    constructor(left: number, right: number, top: number, bottom: number, near: number, far: number);
  }

  export class Scene {
    add(object: object): this;
  }

  export class PlaneGeometry {
    constructor(width?: number, height?: number);
    dispose(): void;
  }

  export class ShaderMaterial {
    constructor(parameters: {
      uniforms: Record<string, { value: unknown }>;
      vertexShader: string;
      fragmentShader: string;
    });
    dispose(): void;
  }

  export class Mesh {
    constructor(geometry: object, material: object);
  }

  export class WebGLRenderer {
    domElement: HTMLCanvasElement;
    constructor(parameters?: { antialias?: boolean; alpha?: boolean });
    setPixelRatio(value: number): void;
    setSize(width: number, height: number): void;
    render(scene: object, camera: object): void;
    dispose(): void;
  }
}
