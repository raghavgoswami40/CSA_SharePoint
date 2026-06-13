import * as React from 'react';
import { IProject } from './IHelloWorldProps';
import styles from './ProjectSelector.module.scss';

export interface IProjectSelectorProps {
  projects: IProject[];
  selectedProjectId: string | undefined;
  onProjectSelect: (project: IProject) => void;
  onAddProject: () => void;
}

export interface IProjectSelectorState {
  hoveredId: string | undefined;
  showAddForm: boolean;
  newProjectName: string;
  newProjectDescription: string;
}

// XLYOUR brand colours cycled across project cards
const PROJECT_COLOURS = [
  '#BD0000', // XLYOUR Red
  '#FFC300', // XLYOUR Yellow
  '#7D7F7C', // XLYOUR Grey
  '#282928', // XLYOUR Black
];

export default class ProjectSelector extends React.Component<IProjectSelectorProps, IProjectSelectorState> {

  constructor(props: IProjectSelectorProps) {
    super(props);
    this.state = {
      hoveredId: undefined,
      showAddForm: false,
      newProjectName: '',
      newProjectDescription: '',
    };
  }

  private _getInitials(name: string): string {
    return name
      .split(' ')
      .map(w => w[0])
      .join('')
      .toUpperCase()
      .slice(0, 2);
  }

  private _handleAddSubmit = (e: React.FormEvent): void => {
    e.preventDefault();
    if (this.state.newProjectName.trim()) {
      this.props.onAddProject();
      this.setState({ showAddForm: false, newProjectName: '', newProjectDescription: '' });
    }
  }

  public render(): React.ReactElement<IProjectSelectorProps> {
    const { projects, selectedProjectId, onProjectSelect } = this.props;
    const { hoveredId, showAddForm, newProjectName, newProjectDescription } = this.state;

    return (
      <div className={styles.selectorRoot}>

        {/* Header bar */}
        <div className={styles.header}>
          <div className={styles.headerLeft}>
            <span className={styles.headerLabel}>PROJECTS</span>
            <span className={styles.headerCount}>{projects.length}</span>
          </div>
          <button
            className={styles.addButton}
            onClick={() => this.setState({ showAddForm: !showAddForm })}
            title="Add new project"
          >
            <span className={styles.addIcon}>{showAddForm ? '✕' : '+'}</span>
            <span>{showAddForm ? 'Cancel' : 'New Project'}</span>
          </button>
        </div>

        {/* Add Project inline form */}
        {showAddForm && (
          <div className={styles.addFormWrapper}>
            <form className={styles.addForm} onSubmit={this._handleAddSubmit}>
              <div className={styles.formTitle}>Create a New Project</div>
              <div className={styles.formRow}>
                <input
                  className={styles.formInput}
                  type="text"
                  placeholder="Project name *"
                  value={newProjectName}
                  onChange={e => this.setState({ newProjectName: e.target.value })}
                  autoFocus
                  required
                />
              </div>
              <div className={styles.formRow}>
                <input
                  className={styles.formInput}
                  type="text"
                  placeholder="Short description (optional)"
                  value={newProjectDescription}
                  onChange={e => this.setState({ newProjectDescription: e.target.value })}
                />
              </div>
              <div className={styles.formActions}>
                <button type="submit" className={styles.submitButton} disabled={!newProjectName.trim()}>
                  Create Project
                </button>
              </div>
            </form>
          </div>
        )}

        {/* Project cards scroll rail */}
        <div className={styles.rail}>
          {projects.map((project, index) => {
            const isSelected = project.id === selectedProjectId;
            const isHovered = project.id === hoveredId;
            const colour = project.colour || PROJECT_COLOURS[index % PROJECT_COLOURS.length];

            return (
              <div
                key={project.id}
                className={`${styles.card} ${isSelected ? styles.cardSelected : ''} ${isHovered ? styles.cardHovered : ''}`}
                onClick={() => onProjectSelect(project)}
                onMouseEnter={() => this.setState({ hoveredId: project.id })}
                onMouseLeave={() => this.setState({ hoveredId: undefined })}
                role="button"
                tabIndex={0}
                onKeyDown={e => e.key === 'Enter' && onProjectSelect(project)}
                aria-pressed={isSelected}
                aria-label={`Select project ${project.name}`}
              >
                {/* Colour accent bar */}
                <div className={styles.cardAccent} style={{ background: colour }} />

                {/* Initials badge */}
                <div className={styles.initials} style={{ background: colour }}>
                  {this._getInitials(project.name)}
                </div>

                {/* Card body */}
                <div className={styles.cardBody}>
                  <div className={styles.cardName}>{project.name}</div>
                  {project.description && (
                    <div className={styles.cardDesc}>{project.description}</div>
                  )}
                  <div className={styles.cardMeta}>
                    {project.documentCount !== undefined && (
                      <span className={styles.metaChip}>
                        📄 {project.documentCount} docs
                      </span>
                    )}
                    {project.lastModified && (
                      <span className={styles.metaChip}>
                        🕐 {project.lastModified}
                      </span>
                    )}
                  </div>
                </div>

                {/* Selected indicator */}
                {isSelected && (
                  <div className={styles.selectedBadge} style={{ background: colour }}>
                    ✓
                  </div>
                )}
              </div>
            );
          })}

          {/* Empty state */}
          {projects.length === 0 && !showAddForm && (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>📁</div>
              <div className={styles.emptyText}>No projects yet</div>
              <div className={styles.emptySubText}>Click "New Project" to get started</div>
            </div>
          )}
        </div>

        {/* Selected project name indicator */}
        {selectedProjectId && (
          <div className={styles.activeBar}>
            <span className={styles.activeLabel}>ACTIVE PROJECT</span>
            <span className={styles.activeName}>
              {(projects.filter((p: IProject) => p.id === selectedProjectId)[0] || { name: '' }).name}
            </span>
          </div>
        )}
      </div>
    );
  }
}
