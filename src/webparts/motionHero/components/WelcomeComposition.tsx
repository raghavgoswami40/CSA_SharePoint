/**
 * WelcomeComposition.tsx
 *
 * Timeline (30 fps):
 *   0  – 36f   Logo fades + scales in, centred
 *   36 – 66f   Logo holds
 *   66 – 81f   Flip OUT — logo scaleX 1 → 0
 *   81 – 96f   Flip IN  — text scaleX 0 → 1
 *   96f+        Text holds; loops at 8 s
 */

import * as React from 'react';
import { interpolate, Easing } from './remotionHelpers';

// ── Frame hook ────────────────────────────────────────────────────────────────

const FPS      = 30;
const DURATION = FPS * 8;

function useCurrentFrame(): number {
  const [frame, setFrame] = React.useState(0);
  const rafRef   = React.useRef<number>(0);
  const startRef = React.useRef<number | null>(null);

  React.useEffect(() => {
    const tick = (now: number): void => {
      if (startRef.current === null) startRef.current = now;
      setFrame(Math.floor(((now - startRef.current) / 1000) * FPS) % DURATION);
      rafRef.current = requestAnimationFrame(tick);
    };
    rafRef.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(rafRef.current);
  }, []);

  return frame;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

const easeOut     = Easing.bezier(0.16, 1, 0.3, 1);
const easeInOut   = Easing.bezier(0.45, 0, 0.55, 1);
const easeOvershoot = Easing.bezier(0.34, 1.18, 0.64, 1);

function fadeIn(frame: number, start: number, dur: number): number {
  return interpolate(frame, [start, start + dur], [0, 1], {
    extrapolateLeft: 'clamp', extrapolateRight: 'clamp', easing: easeOut,
  });
}

// ── Timeline ──────────────────────────────────────────────────────────────────

const LOGO_IN        = 0;
const LOGO_DUR       = 36;
const LOGO_HOLD      = 30;
const FLIP_OUT_START = LOGO_IN + LOGO_DUR + LOGO_HOLD;  // 66
const FLIP_OUT_END   = FLIP_OUT_START + 15;              // 81
const FLIP_IN_END    = FLIP_OUT_END + 15;                // 96

// ── Logo asset ────────────────────────────────────────────────────────────────

// eslint-disable-next-line @typescript-eslint/no-var-requires
const logoSrc: string = require('../assets/logo-white.png');

// ── Composition ───────────────────────────────────────────────────────────────

export const WelcomeComposition: React.FC = () => {
  const frame = useCurrentFrame();

  const bgOpacity = interpolate(frame, [0, LOGO_DUR], [0, 1], {
    extrapolateLeft: 'clamp', extrapolateRight: 'clamp', easing: easeOut,
  });

  const logoOpacity = fadeIn(frame, LOGO_IN, LOGO_DUR);
  const logoScale   = interpolate(frame, [LOGO_IN, LOGO_IN + LOGO_DUR], [0.82, 1], {
    extrapolateLeft: 'clamp', extrapolateRight: 'clamp', easing: easeOvershoot,
  });

  const logoScaleX = interpolate(frame, [FLIP_OUT_START, FLIP_OUT_END], [1, 0], {
    extrapolateLeft: 'clamp', extrapolateRight: 'clamp', easing: easeInOut,
  });

  const textScaleX = interpolate(frame, [FLIP_OUT_END, FLIP_IN_END], [0, 1], {
    extrapolateLeft: 'clamp', extrapolateRight: 'clamp', easing: easeInOut,
  });

  const glowOpacity = interpolate(frame, [FLIP_OUT_START, FLIP_OUT_END], [1, 0], {
    extrapolateLeft: 'clamp', extrapolateRight: 'clamp',
  });

  const showLogo = frame < FLIP_IN_END;
  const showText = frame >= FLIP_OUT_END;

  return (
    <div style={{
      width: '100%',
      height: '100%',
      background: `rgba(40,41,40,${bgOpacity})`,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      position: 'relative',
      overflow: 'hidden',
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      boxSizing: 'border-box',
    }}>

      {/* Radial glow */}
      <div style={{
        position: 'absolute',
        width: 560, height: 280,
        borderRadius: '50%',
        background: 'radial-gradient(ellipse, rgba(189,0,0,0.14) 0%, transparent 70%)',
        opacity: logoOpacity * glowOpacity,
        pointerEvents: 'none',
        top: '50%', left: '50%',
        transform: 'translate(-50%, -50%)',
      }} />

      {/* Logo */}
      {showLogo && (
        <div style={{
          position: 'absolute',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          opacity: logoOpacity,
          transform: `scale(${logoScale}) scaleX(${frame >= FLIP_OUT_START ? logoScaleX : 1})`,
          transformOrigin: 'center center',
        }}>
          <img src={logoSrc} alt="XLYOUR" style={{ width: 280, display: 'block' }} />
        </div>
      )}

      {/* Text */}
      {showText && (
        <div style={{
          position: 'absolute',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          textAlign: 'center',
          transform: `scaleX(${textScaleX})`,
          transformOrigin: 'center center',
          padding: '0 32px',
        }}>
          <div style={{
            fontSize: 13,
            fontWeight: 600,
            letterSpacing: '0.28em',
            color: '#7D7F7C',
            textTransform: 'uppercase',
            marginBottom: 12,
          }}>
            WELCOME TO THE
          </div>
          <div style={{
            fontSize: 30,
            fontWeight: 800,
            color: '#E0E1E0',
            letterSpacing: '0.06em',
            lineHeight: 1.25,
            textTransform: 'uppercase',
          }}>
            CURRENT STATE<br />ANALYSIS SITE
          </div>
          <div style={{
            marginTop: 20,
            width: 64,
            height: 3,
            borderRadius: 2,
            background: 'linear-gradient(90deg, #BD0000, #FFC300)',
          }} />
        </div>
      )}

    </div>
  );
};
