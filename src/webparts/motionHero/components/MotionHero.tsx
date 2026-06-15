import * as React from 'react';
import { WelcomeComposition } from './WelcomeComposition';
import ShaderBackground from './ShaderBackground';
import styles from './MotionHero.module.scss';

export const MotionHero: React.FC = () => (
  <div className={styles.root}>
    {/* WebGL shader fills the entire banner */}
    <ShaderBackground />

    {/* Logo flip + welcome text on top */}
    <div style={{ position: 'absolute', inset: 0, zIndex: 1 }}>
      <WelcomeComposition />
    </div>
  </div>
);

export default MotionHero;
