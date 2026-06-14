/**
 * Lightweight re-implementation of Remotion's core helpers.
 * Follows the same API: interpolate(), Easing, so the animation
 * code reads identically to a Remotion composition.
 */

// ── Cubic Bézier solver ───────────────────────────────────────────────────────
// Newton-Raphson method — same algorithm browsers use for CSS cubic-bezier().

function cubicBezierSample(t: number, p1: number, p2: number): number {
  // B(t) = 3*(1-t)^2*t*p1 + 3*(1-t)*t^2*p2 + t^3  (endpoints 0 and 1)
  return 3 * (1 - t) * (1 - t) * t * p1
       + 3 * (1 - t) * t * t * p2
       + t * t * t;
}

function cubicBezierDerivative(t: number, p1: number, p2: number): number {
  return 3 * (1 - t) * (1 - t) * p1
       + 6 * (1 - t) * t * (p2 - p1)
       + 3 * t * t * (1 - p2);
}

function solveCubicBezierX(x: number, x1: number, x2: number): number {
  // Binary search fallback for edge cases, Newton-Raphson for speed
  let t = x;
  for (let i = 0; i < 8; i++) {
    const currentX = cubicBezierSample(t, x1, x2) - x;
    if (Math.abs(currentX) < 1e-6) break;
    const d = cubicBezierDerivative(t, x1, x2);
    if (Math.abs(d) < 1e-6) break;
    t -= currentX / d;
  }
  return t;
}

// ── Easing ────────────────────────────────────────────────────────────────────

export const Easing = {
  /**
   * Custom cubic-bezier — identical API to Remotion / CSS animations.
   * bezier(0.16, 1, 0.3, 1) is the "crisp ease-out" used throughout XLYOUR.
   */
  bezier: (x1: number, y1: number, x2: number, y2: number) =>
    (t: number): number => {
      if (t === 0 || t === 1) return t;
      const solvedT = solveCubicBezierX(t, x1, x2);
      return cubicBezierSample(solvedT, y1, y2);
    },

  linear: (t: number): number => t,
  sin:    (t: number): number => 1 - Math.cos((t * Math.PI) / 2),
  quad:   (t: number): number => t * t,
  cubic:  (t: number): number => t * t * t,

  in:    (easing: (t: number) => number) => (t: number): number => easing(t),
  out:   (easing: (t: number) => number) => (t: number): number => 1 - easing(1 - t),
  inOut: (easing: (t: number) => number) => (t: number): number =>
    t < 0.5 ? easing(t * 2) / 2 : 1 - easing((1 - t) * 2) / 2,
};

// ── interpolate() ─────────────────────────────────────────────────────────────

export type ExtrapolateType = 'clamp' | 'extend';

export interface InterpolateOptions {
  easing?: (t: number) => number;
  extrapolateLeft?: ExtrapolateType;
  extrapolateRight?: ExtrapolateType;
}

export function interpolate(
  value: number,
  inputRange: [number, number],
  outputRange: [number, number],
  options: InterpolateOptions = {},
): number {
  const [inMin, inMax] = inputRange;
  const [outMin, outMax] = outputRange;

  let t = (value - inMin) / (inMax - inMin);

  if ((options.extrapolateLeft ?? 'extend') === 'clamp') t = Math.max(0, t);
  if ((options.extrapolateRight ?? 'extend') === 'clamp') t = Math.min(1, t);

  if (options.easing) t = options.easing(Math.max(0, Math.min(1, t)));

  return outMin + t * (outMax - outMin);
}
