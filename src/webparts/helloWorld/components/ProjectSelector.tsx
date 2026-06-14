import * as React from 'react';
import { IProject, ISiteSearchResult } from './IHelloWorldProps';
import styles from './ProjectSelector.module.scss';

export interface IProjectSelectorProps {
  projects: IProject[];
  selectedProjectId: string | undefined;
  onProjectSelect: (project: IProject) => void;
  onAddProject: (siteTitle: string, siteUrl: string) => Promise<void>;
  onSearchSites: (query: string) => Promise<ISiteSearchResult[]>;
}

interface IProjectSelectorState {
  hoveredId: string | undefined;
  showAddForm: boolean;
  searchQuery: string;
  searchResults: ISiteSearchResult[];
  isSearching: boolean;
  selectedSite: ISiteSearchResult | undefined;
  isCreating: boolean;
  createError: string | undefined;
}

const PROJECT_COLOURS = ['#FFC300', '#FFC300', '#FFC300', '#FFC300'];

export default class ProjectSelector extends React.Component<IProjectSelectorProps, IProjectSelectorState> {

  private _searchTimer: ReturnType<typeof setTimeout> | undefined;

  constructor(props: IProjectSelectorProps) {
    super(props);
    this.state = {
      hoveredId: undefined,
      showAddForm: false,
      searchQuery: '',
      searchResults: [],
      isSearching: false,
      selectedSite: undefined,
      isCreating: false,
      createError: undefined,
    };
  }

  private _getInitials(name: string): string {
    return name.split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  }

  private _handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const query = e.target.value;
    this.setState({ searchQuery: query, selectedSite: undefined, searchResults: [] });

    if (this._searchTimer) clearTimeout(this._searchTimer);
    if (!query.trim()) return;

    this.setState({ isSearching: true });
    this._searchTimer = setTimeout(async () => {
      const results = await this.props.onSearchSites(query);
      this.setState({ searchResults: results, isSearching: false });
    }, 350); // debounce
  }

  private _handleSiteSelect = (site: ISiteSearchResult): void => {
    this.setState({ selectedSite: site, searchQuery: site.title, searchResults: [] });
  }

  private _handleAddSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    const { selectedSite } = this.state;
    if (!selectedSite) return;

    this.setState({ isCreating: true, createError: undefined });
    try {
      await this.props.onAddProject(selectedSite.title, selectedSite.url);
      this.setState({
        showAddForm: false, searchQuery: '', searchResults: [],
        selectedSite: undefined, isCreating: false,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to add project.';
      this.setState({ isCreating: false, createError: msg });
    }
  }

  public render(): React.ReactElement<IProjectSelectorProps> {
    const { projects, selectedProjectId, onProjectSelect } = this.props;
    const {
      hoveredId, showAddForm, searchQuery, searchResults,
      isSearching, selectedSite, isCreating, createError,
    } = this.state;

    return (
      <div className={styles.selectorRoot}>

        {/* Header */}
        <div className={styles.header}>
          <div className={styles.headerLeft}>
            <span className={styles.headerLabel}>PROJECTS</span>
            <span className={styles.headerCount}>{projects.length}</span>
          </div>
          <button
            className={styles.addButton}
            onClick={() => this.setState({ showAddForm: !showAddForm, searchQuery: '', searchResults: [], selectedSite: undefined, createError: undefined })}
            title="Link a project site"
          >
            <span className={styles.addIcon}>{showAddForm ? '✕' : '+'}</span>
            <span>{showAddForm ? 'Cancel' : 'Add Project'}</span>
          </button>
        </div>

        {/* Add form — site search picker */}
        {showAddForm && (
          <div className={styles.addFormWrapper}>
            <form className={styles.addForm} onSubmit={this._handleAddSubmit}>
              <div className={styles.formTitle}>Link a SharePoint Site</div>
              <div className={styles.formRow} style={{ position: 'relative' }}>
                <input
                  className={styles.formInput}
                  type="text"
                  placeholder="Search for a SharePoint site…"
                  value={searchQuery}
                  onChange={this._handleSearchChange}
                  autoFocus
                  autoComplete="off"
                />
                {isSearching && (
                  <div className={styles.searchSpinner}>⟳</div>
                )}
                {/* Search results dropdown */}
                {searchResults.length > 0 && (
                  <div className={styles.searchDropdown}>
                    {searchResults.map((site, i) => (
                      <div
                        key={i}
                        className={styles.searchResult}
                        onClick={() => this._handleSiteSelect(site)}
                        role="button"
                        tabIndex={0}
                        onKeyDown={e => e.key === 'Enter' && this._handleSiteSelect(site)}
                      >
                        <div className={styles.searchResultTitle}>{site.title}</div>
                        <div className={styles.searchResultUrl}>{site.url}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Selected site confirmation */}
              {selectedSite && (
                <div className={styles.selectedSite}>
                  <span className={styles.selectedSiteCheck}>✓</span>
                  <div>
                    <div className={styles.selectedSiteTitle}>{selectedSite.title}</div>
                    <div className={styles.selectedSiteUrl}>{selectedSite.url}</div>
                  </div>
                </div>
              )}

              {createError && (
                <div className={styles.formError}>{createError}</div>
              )}

              <div className={styles.formActions}>
                <button
                  type="submit"
                  className={styles.submitButton}
                  disabled={!selectedSite || isCreating}
                >
                  {isCreating ? 'Adding…' : 'Add Project'}
                </button>
              </div>
            </form>
          </div>
        )}

        {/* Project cards */}
        <div className={styles.rail}>
          {projects.map((project, index) => {
            const isSelected = project.id === selectedProjectId;
            const isHovered  = project.id === hoveredId;
            const colour     = project.colour || PROJECT_COLOURS[index % PROJECT_COLOURS.length];

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
                <div className={styles.cardAccent} style={{ background: colour }} />
                <div className={styles.initials} style={{ background: colour }}>
                  {this._getInitials(project.name)}
                </div>
                <div className={styles.cardBody}>
                  <div className={styles.cardName}>{project.name}</div>
                  {project.description && (
                    <div className={styles.cardDesc}>{project.description}</div>
                  )}
                  {project.lastModified && (
                    <div className={styles.cardMeta}>
                      <span className={styles.metaChip}>🕐 {project.lastModified}</span>
                    </div>
                  )}
                </div>
                {isSelected && (
                  <div className={styles.selectedBadge} style={{ background: colour }}>✓</div>
                )}
              </div>
            );
          })}

          {projects.length === 0 && !showAddForm && (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>📁</div>
              <div className={styles.emptyText}>You have not added any projects.</div>
              <div className={styles.emptySubText}>
                <span
                  className={styles.emptyLink}
                  onClick={() => this.setState({ showAddForm: true })}
                  role="button"
                  tabIndex={0}
                  onKeyDown={e => e.key === 'Enter' && this.setState({ showAddForm: true })}
                >
                  Add a project now
                </span>
              </div>
            </div>
          )}
        </div>

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
