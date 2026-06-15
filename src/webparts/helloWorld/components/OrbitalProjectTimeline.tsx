/**
 * OrbitalProjectTimeline.tsx
 *
 * Adapted from the RadialOrbitalTimeline component.
 * Converted from Tailwind/shadcn to SPFx-compatible inline styles + SCSS animations.
 * Shows projects as orbital nodes. Empty state shows + placeholder nodes.
 */

import * as React from 'react';
import { FolderOpen, Plus, ExternalLink } from 'lucide-react';
import { IProject } from './IHelloWorldProps';
import styles from './OrbitalProjectTimeline.module.scss';

// ── Brand colours ─────────────────────────────────────────────────────────────
const C = {
  black:  '#282928',
  white:  '#E0E1E0',
  grey:   '#7D7F7C',
  yellow: '#FFC300',
  red:    '#BD0000',
};

// ── Props ─────────────────────────────────────────────────────────────────────

export interface IOrbitalProjectTimelineProps {
  projects: IProject[];
  selectedProjectId: string | undefined;
  onProjectSelect: (project: IProject) => void;
  onProjectDeselect: () => void;
  onAddProjectClick: () => void;
}

// ── Internal node shape ───────────────────────────────────────────────────────

interface INode {
  id: string;
  title: string;
  siteUrl: string;
  lastModified: string;
  isPlaceholder: boolean;
  project?: IProject;
}

// ── State ─────────────────────────────────────────────────────────────────────

interface IState {
  expandedId: string | undefined;
  rotationAngle: number;
  autoRotate: boolean;
}

const EMPTY_COUNT  = 5;   // placeholder nodes when no projects
const ORBIT_RADIUS = 190;

export default class OrbitalProjectTimeline extends React.Component<IOrbitalProjectTimelineProps, IState> {

  private _rafId: number = 0;
  private _lastTime: number = 0;

  constructor(props: IOrbitalProjectTimelineProps) {
    super(props);
    this.state = { expandedId: undefined, rotationAngle: 0, autoRotate: true };
  }

  public componentDidMount(): void { this._startRotation(); }
  public componentWillUnmount(): void { cancelAnimationFrame(this._rafId); }

  // ── Auto-rotation via RAF ─────────────────────────────────────────────────

  private _startRotation = (): void => {
    const tick = (now: number): void => {
      if (this.state.autoRotate) {
        const delta = now - (this._lastTime || now);
        this._lastTime = now;
        this.setState(prev => ({
          rotationAngle: (prev.rotationAngle + delta * 0.018) % 360,
        }));
      } else {
        this._lastTime = now;
      }
      this._rafId = requestAnimationFrame(tick);
    };
    this._rafId = requestAnimationFrame(tick);
  }

  // ── Node position ─────────────────────────────────────────────────────────

  private _nodePos(index: number, total: number): { x: number; y: number; zIndex: number; opacity: number } {
    const angle = ((index / total) * 360 + this.state.rotationAngle) % 360;
    const rad   = (angle * Math.PI) / 180;
    const x     = ORBIT_RADIUS * Math.cos(rad);
    const y     = ORBIT_RADIUS * Math.sin(rad);
    const zIndex  = Math.round(100 + 50 * Math.cos(rad));
    const opacity = Math.max(0.4, Math.min(1, 0.4 + 0.6 * ((1 + Math.sin(rad)) / 2)));
    return { x, y, zIndex, opacity };
  }

  // ── Toggle expanded node ──────────────────────────────────────────────────

  private _toggle = (node: INode): void => {
    if (node.isPlaceholder) {
      this.props.onAddProjectClick();
      return;
    }
    const isOpening = this.state.expandedId !== node.id;
    this.setState({ expandedId: isOpening ? node.id : undefined, autoRotate: !isOpening });
    if (isOpening && node.project) {
      this.props.onProjectSelect(node.project);
    } else {
      this.props.onProjectDeselect();
    }
  }

  private _handleContainerClick = (): void => {
    this.setState({ expandedId: undefined, autoRotate: true });
  }

  // ── Build node list ───────────────────────────────────────────────────────

  private _getNodes(): INode[] {
    const { projects } = this.props;
    const projectNodes: INode[] = projects.map(p => ({
      id: p.id,
      title: p.name,
      siteUrl: p.siteUrl || '',
      lastModified: p.lastModified || '',
      isPlaceholder: false,
      project: p,
    }));

    // Always fill up to EMPTY_COUNT total nodes, with at least 1 add-slot
    const totalNodes = Math.max(EMPTY_COUNT, projectNodes.length + 1);
    const placeholderCount = totalNodes - projectNodes.length;

    const placeholders: INode[] = Array.from({ length: placeholderCount }, (_, i) => ({
      id: `placeholder-${i}`,
      title: 'Add Project',
      siteUrl: '',
      lastModified: '',
      isPlaceholder: true,
    }));

    return [...projectNodes, ...placeholders];
  }

  // ── Render ────────────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const { selectedProjectId } = this.props;
    const { expandedId } = this.state;
    const nodes = this._getNodes();

    return (
      <div
        onClick={this._handleContainerClick}
        style={{
          width: '100%',
          height: 480,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          background: C.black,
          position: 'relative',
          overflow: 'hidden',
          borderRadius: 16,
        }}
      >
        {/* Orbit container */}
        <div style={{ position: 'relative', width: 480, height: 480, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>

          {/* Orbit ring */}
          <div style={{
            position: 'absolute',
            width: ORBIT_RADIUS * 2,
            height: ORBIT_RADIUS * 2,
            borderRadius: '50%',
            border: `1px solid rgba(224,225,224,0.1)`,
            top: '50%', left: '50%',
            transform: 'translate(-50%, -50%)',
          }} />

          {/* Centre orb */}
          <div
            className={styles.animatePulse}
            style={{
              position: 'absolute',
              width: 64, height: 64,
              borderRadius: '50%',
              background: `linear-gradient(135deg, #6366f1, #3b82f6, #14b8a6)`,
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              zIndex: 10,
            }}
          >
            <div className={styles.animatePing} style={{
              position: 'absolute', width: 80, height: 80,
              borderRadius: '50%',
              border: `1px solid rgba(224,225,224,0.2)`,
            }} />
            <div className={styles.animatePingDelay} style={{
              position: 'absolute', width: 96, height: 96,
              borderRadius: '50%',
              border: `1px solid rgba(224,225,224,0.1)`,
            }} />
            <div style={{ width: 28, height: 28, borderRadius: '50%', background: 'rgba(224,225,224,0.85)' }} />
          </div>

          {/* Nodes */}
          {nodes.map((node, index) => {
            const pos       = this._nodePos(index, nodes.length);
            const isExpanded = expandedId === node.id;
            const isSelected = !node.isPlaceholder && node.id === selectedProjectId;
            const isPulse    = node.isPlaceholder;

            return (
              <div
                key={node.id}
                onClick={e => { e.stopPropagation(); this._toggle(node); }}
                style={{
                  position: 'absolute',
                  top: '50%', left: '50%',
                  transform: `translate(calc(-50% + ${pos.x}px), calc(-50% + ${pos.y}px))`,
                  zIndex: isExpanded ? 200 : pos.zIndex,
                  opacity: isExpanded ? 1 : pos.opacity,
                  transition: 'opacity 0.3s, transform 0.7s',
                  cursor: 'pointer',
                }}
              >
                {/* Glow ring */}
                <div style={{
                  position: 'absolute',
                  borderRadius: '50%',
                  background: 'radial-gradient(circle, rgba(255,195,0,0.2) 0%, transparent 70%)',
                  width: 60, height: 60,
                  top: -10, left: -10,
                  pointerEvents: 'none',
                }} />

                {/* Icon circle */}
                <div
                  className={isPulse ? styles.nodePulse : ''}
                  style={{
                    width: 40, height: 40,
                    borderRadius: '50%',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    background: isExpanded ? C.white : isSelected ? C.yellow : 'rgba(40,41,40,0.9)',
                    border: `2px solid ${isExpanded ? C.white : isSelected ? C.yellow : node.isPlaceholder ? `rgba(255,195,0,0.5)` : 'rgba(224,225,224,0.4)'}`,
                    color: isExpanded ? C.black : isSelected ? C.black : C.white,
                    transform: isExpanded ? 'scale(1.5)' : 'scale(1)',
                    transition: 'all 0.3s',
                    boxShadow: isExpanded ? `0 0 20px rgba(224,225,224,0.3)` : isSelected ? `0 0 12px rgba(255,195,0,0.4)` : 'none',
                  }}
                >
                  {node.isPlaceholder
                    ? <Plus size={16} />
                    : <FolderOpen size={16} />
                  }
                </div>

                {/* Label */}
                <div style={{
                  position: 'absolute',
                  top: 44,
                  left: '50%',
                  transform: 'translateX(-50%)',
                  whiteSpace: 'nowrap',
                  fontSize: 11,
                  fontWeight: 600,
                  letterSpacing: '0.05em',
                  color: isExpanded ? C.white : 'rgba(224,225,224,0.7)',
                  fontFamily: "'Segoe UI', sans-serif",
                  transition: 'color 0.3s',
                  maxWidth: 100,
                  textOverflow: 'ellipsis',
                  overflow: 'hidden',
                  textAlign: 'center',
                }}>
                  {node.isPlaceholder ? '+ Add' : node.title}
                </div>

                {/* Expanded card — flips above node when in lower half of orbit */}
                {isExpanded && !node.isPlaceholder && (
                  <div
                    onClick={e => e.stopPropagation()}
                    style={{
                      position: 'absolute',
                      ...(pos.y > 40
                        ? { bottom: 60 }   // lower half → card goes up
                        : { top: 60 }),     // upper half → card goes down
                      left: '50%',
                      transform: 'translateX(-50%)',
                      width: 240,
                      background: 'rgba(20,20,20,0.95)',
                      backdropFilter: 'blur(16px)',
                      border: `1px solid rgba(224,225,224,0.2)`,
                      borderRadius: 12,
                      padding: 16,
                      boxShadow: '0 8px 32px rgba(0,0,0,0.6)',
                      zIndex: 300,
                      fontFamily: "'Segoe UI', sans-serif",
                    }}
                  >
                    {/* Connector line — adjusts side based on card direction */}
                    <div style={{
                      position: 'absolute',
                      ...(pos.y > 40
                        ? { bottom: -12 }
                        : { top: -12 }),
                      left: '50%',
                      transform: 'translateX(-50%)',
                      width: 1, height: 12,
                      background: 'rgba(224,225,224,0.4)',
                    }} />

                    {/* Title */}
                    <div style={{ fontSize: 13, fontWeight: 700, color: C.white, marginBottom: 10 }}>
                      {node.title}
                    </div>

                    {/* Open site link */}
                    {node.siteUrl && (
                      <a
                        href={node.siteUrl}
                        target="_blank"
                        rel="noreferrer"
                        onClick={e => e.stopPropagation()}
                        style={{
                          fontSize: 12, fontWeight: 600,
                          color: C.yellow,
                          textDecoration: 'none',
                          display: 'flex', alignItems: 'center', gap: 6,
                          padding: '7px 12px',
                          background: 'rgba(255,195,0,0.1)',
                          border: `1px solid rgba(255,195,0,0.3)`,
                          borderRadius: 6,
                        }}
                      >
                        <ExternalLink size={12} />
                        Open site
                      </a>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* Empty state label — only when zero real projects */}
        {this.props.projects.length === 0 && (
          <div style={{
            position: 'absolute', bottom: 24,
            fontSize: 13, color: 'rgba(224,225,224,0.4)',
            fontFamily: "'Segoe UI', sans-serif",
            letterSpacing: '0.05em',
          }}>
            You have not added any projects — click <span style={{ color: C.yellow }}>+</span> to add one
          </div>
        )}
      </div>
    );
  }
}
