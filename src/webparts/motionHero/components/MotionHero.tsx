import * as React from 'react';
import { WelcomeComposition } from './WelcomeComposition';
import styles from './MotionHero.module.scss';

export const MotionHero: React.FC = () => (
  <div className={styles.root}>
    <WelcomeComposition />
  </div>
);

export default MotionHero;
