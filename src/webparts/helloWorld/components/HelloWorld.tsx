import * as React from 'react';
import styles from './HelloWorld.module.scss';
import type { IHelloWorldProps, IProject } from './IHelloWorldProps';
import ProjectSelector from './ProjectSelector';

const PROJECT_COLOURS = ['#BD0000', '#FFC300', '#7D7F7C', '#282928'];

interface IHelloWorldState {
  projects: IProject[];
  selectedProjectId: string | undefined;
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {
      selectedProjectId: 'proj-1',
      projects: [
        {
          id: 'proj-1',
          name: 'Digital Transformation',
          description: 'Enterprise modernisation initiative',
          colour: PROJECT_COLOURS[0],
          documentCount: 42,
          lastModified: '2 days ago',
        },
        {
          id: 'proj-2',
          name: 'HR Onboarding',
          description: 'New employee experience',
          colour: PROJECT_COLOURS[1],
          documentCount: 18,
          lastModified: 'Today',
        },
        {
          id: 'proj-3',
          name: 'Finance Q3',
          description: 'Quarterly review documents',
          colour: PROJECT_COLOURS[2],
          documentCount: 31,
          lastModified: '1 week ago',
        },
        {
          id: 'proj-4',
          name: 'Client Portal',
          description: 'External-facing project assets',
          colour: PROJECT_COLOURS[3],
          documentCount: 9,
          lastModified: '3 days ago',
        },
      ],
    };
  }

  private _handleProjectSelect = (project: IProject): void => {
    this.setState({ selectedProjectId: project.id });
    console.log(`Selected project: ${project.name}`);
  }

  private _handleAddProject = (): void => {
    // TODO: Create new SharePoint document library
    console.log('Add new project document library');
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const { projects, selectedProjectId } = this.state;
    const { hasTeamsContext } = this.props;

    const selectedProject: IProject | undefined = projects.filter((p: IProject) => p.id === selectedProjectId)[0];

    return (
      <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''}`}>

        <ProjectSelector
          projects={projects}
          selectedProjectId={selectedProjectId}
          onProjectSelect={this._handleProjectSelect}
          onAddProject={this._handleAddProject}
        />

        {/* Document library area */}
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

      </section>
    );
  }
}
