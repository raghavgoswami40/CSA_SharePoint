import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './HelloWorld.module.scss';
import type { IHelloWorldProps, IProject, ISiteSearchResult } from './IHelloWorldProps';
import ProjectSelector from './ProjectSelector';
import DocumentLibraryView from './DocumentLibraryView';

const PROJECT_COLOUR  = '#FFC300';
const PROJECTS_LIST    = 'CSA Projects';
const SITE_URL_FIELD   = 'LinkedSiteUrl'; // avoid reserved SharePoint name "SiteUrl"

interface IHelloWorldState {
  projects: IProject[];
  selectedProjectId: string | undefined;
  isLoading: boolean;
  error: string | undefined;
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = { projects: [], selectedProjectId: undefined, isLoading: true, error: undefined };
  }

  public componentDidMount(): void {
    this._ensureProjectsListThenLoad().catch(err => console.error(err));
  }

  // ── Ensure the "CSA Projects" list exists, then load projects ───────────────

  private async _ensureProjectsListThenLoad(): Promise<void> {
    const { spHttpClient, webUrl } = this.props;
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
    };

    try {
      // Check if list exists
      const checkRes: SPHttpClientResponse = await spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${PROJECTS_LIST}')?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (!checkRes.ok) {
        console.log(`[Projects] List "${PROJECTS_LIST}" not found — creating...`);

        // Create the list
        const createRes: SPHttpClientResponse = await spHttpClient.post(
          `${webUrl}/_api/web/lists`,
          SPHttpClient.configurations.v1,
          {
            headers,
            body: JSON.stringify({
              Title: PROJECTS_LIST,
              BaseTemplate: 100,
              Description: 'Stores linked SharePoint project sites',
            }),
          }
        );

        if (!createRes.ok) {
          const errBody = await createRes.json();
          throw new Error(`Failed to create list: ${errBody['odata.error']?.message?.value || createRes.status}`);
        }

        console.log(`[Projects] List created — adding ${SITE_URL_FIELD} column...`);

        const fieldRes: SPHttpClientResponse = await spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${PROJECTS_LIST}')/fields`,
          SPHttpClient.configurations.v1,
          {
            headers,
            body: JSON.stringify({
              FieldTypeKind: 2,
              Title: SITE_URL_FIELD,
              StaticName: SITE_URL_FIELD,
            }),
          }
        );

        if (!fieldRes.ok) {
          const errBody = await fieldRes.json();
          throw new Error(`Failed to add field: ${errBody['odata.error']?.message?.value || fieldRes.status}`);
        }

        console.log(`[Projects] List setup complete.`);
      }

      await this._loadProjects();

    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[Projects] Setup error: ${msg}`);
      this.setState({ isLoading: false, error: `Setup error: ${msg}` });
    }
  }

  // ── Load projects from the CSA Projects list ────────────────────────────────

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

  // ── Search SharePoint sites ─────────────────────────────────────────────────

  public searchSites = async (query: string): Promise<ISiteSearchResult[]> => {
    if (!query.trim()) return [];
    try {
      const searchQuery = `${encodeURIComponent(query)} contentclass:STS_Site`;
      const res: SPHttpClientResponse = await this.props.spHttpClient.get(
        `${this.props.webUrl}/_api/search/query` +
        `?querytext='${searchQuery}'` +
        `&selectproperties='Title,Path'` +
        `&rowlimit=8&trimduplicates=true`,
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
    } catch {
      return [];
    }
  }

  // ── Link a selected site as a new project ───────────────────────────────────

  private _handleAddProject = async (siteTitle: string, siteUrl: string): Promise<void> => {
    const { spHttpClient, webUrl } = this.props;
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
    };

    // 1 — Save to CSA Projects list
    const listRes: SPHttpClientResponse = await spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${PROJECTS_LIST}')/items`,
      SPHttpClient.configurations.v1,
      { headers, body: JSON.stringify({ Title: siteTitle, LinkedSiteUrl: siteUrl }) }
    );

    if (!listRes.ok) {
      const err = await listRes.json();
      throw new Error(err['odata.error']?.message?.value || 'Failed to add project');
    }

    // 2 — Create document library with the same name on the current site
    const libRes: SPHttpClientResponse = await spHttpClient.post(
      `${webUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers,
        body: JSON.stringify({
          Title: siteTitle,
          BaseTemplate: 101,  // Document Library
          Description: `Document library for project: ${siteTitle}`,
        }),
      }
    );

    if (!libRes.ok) {
      const err = await libRes.json();
      throw new Error(err['odata.error']?.message?.value || 'Failed to create document library');
    }

    // 3 — Scaffold the two mandatory folders
    const folders = ['01 Artefacts', '02 Workshops'];
    for (const folder of folders) {
      await spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${siteTitle}')/rootfolder/folders/add('${encodeURIComponent(folder)}')`,
        SPHttpClient.configurations.v1,
        { headers, body: JSON.stringify({}) }
      );
    }

    console.log(`[Projects] Created library "${siteTitle}" with folders.`);
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
    const { projects, selectedProjectId, isLoading, error } = this.state;
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
            <ProjectSelector
              projects={projects}
              selectedProjectId={selectedProjectId}
              onProjectSelect={this._handleProjectSelect}
              onAddProject={this._handleAddProject}
              onSearchSites={this.searchSites}
            />
            <div className={styles.contentArea}>
              {selectedProject ? (
                <DocumentLibraryView
                  libraryName={selectedProject.name}
                  spHttpClient={this.props.spHttpClient}
                  webUrl={this.props.webUrl}
                />
              ) : (
                <div className={styles.libraryPlaceholder}>
                  <span className={styles.libraryIcon}>👆</span>
                  <div className={styles.libraryText}>Select a project above to view documents</div>
                </div>
              )}
            </div>
          </>
        )}
      </section>
    );
  }
}
