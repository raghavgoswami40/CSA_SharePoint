import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './HelloWorld.module.scss';
import type { IHelloWorldProps, IProject } from './IHelloWorldProps';
import ProjectSelector from './ProjectSelector';

// System libraries that always exist on a SharePoint site — exclude from project list
const SYSTEM_LIBRARIES = [
  'Documents', 'Form Templates', 'FormServerTemplates', 'Site Assets',
  'Site Collection Documents', 'Site Collection Images', 'Style Library',
  'Preservation Hold Library', 'Pages', 'Site Pages',
];

const PROJECT_COLOUR = '#FFC300'; // XLYOUR Yellow — all projects

interface IHelloWorldState {
  projects: IProject[];
  selectedProjectId: string | undefined;
  isLoading: boolean;
  error: string | undefined;
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {
      projects: [],
      selectedProjectId: undefined,
      isLoading: true,
      error: undefined,
    };
  }

  public componentDidMount(): void {
    this._loadProjects().catch(err => console.error('Failed to load projects', err));
  }

  // ── Load all document libraries from SharePoint ─────────────────────────────

  private async _loadProjects(): Promise<void> {
    this.setState({ isLoading: true, error: undefined });
    try {
      const response: SPHttpClientResponse = await this.props.spHttpClient.get(
        `${this.props.webUrl}/_api/web/lists` +
        `?$filter=BaseTemplate eq 101 and Hidden eq false` +
        `&$select=Id,Title,Description,ItemCount,LastItemModifiedDate` +
        `&$orderby=Title`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      const data = await response.json();

      const projects: IProject[] = (data.value || [])
        .filter((lib: { Title: string }) => !SYSTEM_LIBRARIES.includes(lib.Title))
        .map((lib: { Id: string; Title: string; Description: string; ItemCount: number; LastItemModifiedDate: string }) => ({
          id: lib.Id,
          name: lib.Title,
          description: lib.Description || undefined,
          colour: PROJECT_COLOUR,
          documentCount: lib.ItemCount,
          lastModified: this._formatDate(lib.LastItemModifiedDate),
        }));

      this.setState({ projects, isLoading: false });
    } catch {
      this.setState({ isLoading: false, error: 'Could not load projects.' });
    }
  }

  // ── Create a new document library + scaffold folders ────────────────────────

  private _handleAddProject = async (name: string, description: string): Promise<void> => {
    const { spHttpClient, webUrl } = this.props;
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
    };

    // 1 — Create the document library
    const createRes: SPHttpClientResponse = await spHttpClient.post(
      `${webUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      { headers, body: JSON.stringify({ Title: name, Description: description, BaseTemplate: 101 }) }
    );
    if (!createRes.ok) {
      const err = await createRes.json();
      throw new Error(err['odata.error']?.message?.value || 'Failed to create library');
    }

    // 2 — Scaffold the two mandatory folders
    const folders = ['01 Artefacts', '02 Workshops'];
    for (const folder of folders) {
      await spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(name)}')/rootfolder/folders/add('${encodeURIComponent(folder)}')`,
        SPHttpClient.configurations.v1,
        { headers, body: JSON.stringify({}) }
      );
    }

    // 3 — Reload so the new library appears in the list
    await this._loadProjects();
  }

  private _handleProjectSelect = (project: IProject): void => {
    this.setState({ selectedProjectId: project.id });
  }

  private _formatDate(isoDate: string): string {
    if (!isoDate) return '';
    const d = new Date(isoDate);
    const now = new Date();
    const diffDays = Math.floor((now.getTime() - d.getTime()) / 86400000);
    if (diffDays === 0) return 'Today';
    if (diffDays === 1) return 'Yesterday';
    if (diffDays < 7) return `${diffDays} days ago`;
    if (diffDays < 30) return `${Math.floor(diffDays / 7)} week${Math.floor(diffDays / 7) > 1 ? 's' : ''} ago`;
    return d.toLocaleDateString();
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const { projects, selectedProjectId, isLoading, error } = this.state;
    const { hasTeamsContext } = this.props;
    const selectedProject = projects.filter(p => p.id === selectedProjectId)[0];

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
            />

            <div className={styles.contentArea}>
              {selectedProject ? (
                <div className={styles.libraryPlaceholder}>
                  <span className={styles.libraryIcon}>📂</span>
                  <div className={styles.libraryText}>{selectedProject.name}</div>
                  <div className={styles.libraryHint}>Document library content will appear here</div>
                </div>
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
