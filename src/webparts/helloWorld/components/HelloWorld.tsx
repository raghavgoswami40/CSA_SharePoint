import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './HelloWorld.module.scss';
import type { IHelloWorldProps, IProject, ISiteSearchResult } from './IHelloWorldProps';
import OrbitalProjectTimeline from './OrbitalProjectTimeline';
import AddProjectForm from './AddProjectForm';
import DocumentLibraryView from './DocumentLibraryView';

const PROJECT_COLOUR = '#FFC300';
const PROJECTS_LIST  = 'CSA Projects';
const SITE_URL_FIELD = 'LinkedSiteUrl';

interface IHelloWorldState {
  projects: IProject[];
  selectedProjectId: string | undefined;
  isLoading: boolean;
  error: string | undefined;
  showAddForm: boolean;
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {
      projects: [],
      selectedProjectId: undefined,
      isLoading: true,
      error: undefined,
      showAddForm: false,
    };
  }

  public componentDidMount(): void {
    this._ensureProjectsListThenLoad().catch(err => console.error(err));
  }

  // ── Ensure list exists ───────────────────────────────────────────────────────

  private async _ensureProjectsListThenLoad(): Promise<void> {
    const { spHttpClient, webUrl } = this.props;
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
    };
    try {
      const checkRes = await spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${PROJECTS_LIST}')?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (!checkRes.ok) {
        console.log(`[Projects] Creating list...`);
        const createRes = await spHttpClient.post(
          `${webUrl}/_api/web/lists`,
          SPHttpClient.configurations.v1,
          { headers, body: JSON.stringify({ Title: PROJECTS_LIST, BaseTemplate: 100 }) }
        );
        if (!createRes.ok) throw new Error(`Failed to create list: ${createRes.status}`);

        const fieldRes = await spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${PROJECTS_LIST}')/fields`,
          SPHttpClient.configurations.v1,
          { headers, body: JSON.stringify({ FieldTypeKind: 2, Title: SITE_URL_FIELD, StaticName: SITE_URL_FIELD }) }
        );
        if (!fieldRes.ok) throw new Error(`Failed to add field: ${fieldRes.status}`);
      }
      await this._loadProjects();
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[Projects] Setup error: ${msg}`);
      this.setState({ isLoading: false, error: `Setup error: ${msg}` });
    }
  }

  // ── Load projects ────────────────────────────────────────────────────────────

  private async _loadProjects(): Promise<void> {
    this.setState({ isLoading: true, error: undefined });
    try {
      const res: SPHttpClientResponse = await this.props.spHttpClient.get(
        `${this.props.webUrl}/_api/web/lists/getbytitle('${PROJECTS_LIST}')/items`,
        SPHttpClient.configurations.v1
      );
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      const projects: IProject[] = (data.value || []).map((item: {
        Id: number; Title: string; LinkedSiteUrl: string; Modified: string;
      }) => ({
        id: String(item.Id),
        name: item.Title,
        siteUrl: item.LinkedSiteUrl || '',
        colour: PROJECT_COLOUR,
        lastModified: this._formatDate(item.Modified),
      }));
      this.setState({ projects, isLoading: false });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[Projects] Load error: ${msg}`);
      this.setState({ isLoading: false, error: `Could not load projects: ${msg}` });
    }
  }

  // ── Search sites ─────────────────────────────────────────────────────────────

  public searchSites = async (query: string): Promise<ISiteSearchResult[]> => {
    if (!query.trim()) return [];
    try {
      const res: SPHttpClientResponse = await this.props.spHttpClient.get(
        `${this.props.webUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(query)} contentclass:STS_Site'` +
        `&selectproperties='Title,Path'&rowlimit=8&trimduplicates=true`,
        SPHttpClient.configurations.v1
      );
      if (!res.ok) return [];
      const data = await res.json();
      const rows = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
      return rows.map((row: { Cells: { Key: string; Value: string }[] }) => {
        const cells: { [key: string]: string } = {};
        row.Cells.forEach((c: { Key: string; Value: string }) => { cells[c.Key] = c.Value; });
        return { title: cells.Title, url: cells.Path };
      }).filter((s: ISiteSearchResult) => s.title && s.url);
    } catch { return []; }
  }

  // ── Add project ──────────────────────────────────────────────────────────────

  public handleAddProject = async (siteTitle: string, siteUrl: string): Promise<void> => {
    const { spHttpClient, webUrl } = this.props;
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
    };

    // 1 — Save to list
    const listRes = await spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${PROJECTS_LIST}')/items`,
      SPHttpClient.configurations.v1,
      { headers, body: JSON.stringify({ Title: siteTitle, LinkedSiteUrl: siteUrl }) }
    );
    if (!listRes.ok) {
      const err = await listRes.json();
      throw new Error(err['odata.error']?.message?.value || 'Failed to add project');
    }

    // 2 — Create document library
    const libRes = await spHttpClient.post(
      `${webUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      { headers, body: JSON.stringify({ Title: siteTitle, BaseTemplate: 101 }) }
    );
    if (!libRes.ok) {
      const err = await libRes.json();
      throw new Error(err['odata.error']?.message?.value || 'Failed to create library');
    }

    // 3 — Scaffold folders
    for (const folder of ['01 Artefacts', '02 Workshops']) {
      await spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${siteTitle}')/rootfolder/folders/add('${encodeURIComponent(folder)}')`,
        SPHttpClient.configurations.v1,
        { headers, body: JSON.stringify({}) }
      );
    }

    this.setState({ showAddForm: false });
    await this._loadProjects();
  }

  private _handleProjectSelect = (project: IProject): void => {
    this.setState({ selectedProjectId: project.id });
  }

  private _formatDate(isoDate: string): string {
    if (!isoDate) return '';
    const d = new Date(isoDate);
    const diffDays = Math.floor((Date.now() - d.getTime()) / 86400000);
    if (diffDays === 0) return 'Today';
    if (diffDays === 1) return 'Yesterday';
    if (diffDays < 7) return `${diffDays} days ago`;
    if (diffDays < 30) return `${Math.floor(diffDays / 7)} week${Math.floor(diffDays / 7) > 1 ? 's' : ''} ago`;
    return d.toLocaleDateString();
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const { projects, selectedProjectId, isLoading, error, showAddForm } = this.state;
    const { hasTeamsContext } = this.props;
    const selectedProject = projects.filter((p: IProject) => p.id === selectedProjectId)[0];

    return (
      <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''}`}>

        {isLoading ? (
          <div className={styles.loadingState}>Loading projects…</div>
        ) : error ? (
          <div className={styles.errorState}>{error}</div>
        ) : (
          <>
            {/* Orbital timeline */}
            <OrbitalProjectTimeline
              projects={projects}
              selectedProjectId={selectedProjectId}
              onProjectSelect={this._handleProjectSelect}
              onProjectDeselect={() => this.setState({ selectedProjectId: undefined })}
              onAddProjectClick={() => this.setState({ showAddForm: true })}
            />

            {/* Add project form (modal overlay) */}
            {showAddForm && (
              <AddProjectForm
                onAdd={this.handleAddProject}
                onSearchSites={this.searchSites}
                onCancel={() => this.setState({ showAddForm: false })}
              />
            )}

            {/* Document library area — always shown with gap below orbital */}
            <div className={styles.contentArea} style={{ marginTop: 24 }}>
              {selectedProject ? (
                <DocumentLibraryView
                  libraryName={selectedProject.name}
                  spHttpClient={this.props.spHttpClient}
                  webUrl={this.props.webUrl}
                />
              ) : (
                <div className={styles.libraryPlaceholder}>
                  <span className={styles.libraryIcon}>🪐</span>
                  <div className={styles.libraryText}>Select a project from the orbit</div>
                  <div className={styles.libraryHint}>Click any project node above to view its document library</div>
                </div>
              )}
            </div>
          </>
        )}
      </section>
    );
  }
}
