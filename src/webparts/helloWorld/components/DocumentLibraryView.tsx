import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './DocumentLibraryView.module.scss';

export interface IDocumentLibraryViewProps {
  libraryName: string;
  spHttpClient: SPHttpClient;
  webUrl: string;
}

interface IFolder {
  name: string;
  serverRelativeUrl: string;
  itemCount: number;
  timeLastModified: string;
}

interface IFile {
  name: string;
  serverRelativeUrl: string;
  length: number;
  timeLastModified: string;
  uniqueId: string;
}

interface IDocumentLibraryViewState {
  currentFolderUrl: string;
  breadcrumbs: { name: string; url: string }[];
  folders: IFolder[];
  files: IFile[];
  isLoading: boolean;
  error: string | undefined;
  rootUrl: string | undefined;
  // New folder
  showNewFolder: boolean;
  newFolderName: string;
  isCreatingFolder: boolean;
  folderError: string | undefined;
  // Upload
  isUploading: boolean;
  uploadError: string | undefined;
  uploadingFileName: string | undefined;
}

export default class DocumentLibraryView extends React.Component<IDocumentLibraryViewProps, IDocumentLibraryViewState> {

  private _fileInputRef = React.createRef<HTMLInputElement>();

  constructor(props: IDocumentLibraryViewProps) {
    super(props);
    this.state = {
      currentFolderUrl: '',
      breadcrumbs: [],
      folders: [],
      files: [],
      isLoading: true,
      error: undefined,
      rootUrl: undefined,
      showNewFolder: false,
      newFolderName: '',
      isCreatingFolder: false,
      folderError: undefined,
      isUploading: false,
      uploadError: undefined,
      uploadingFileName: undefined,
    };
  }

  public componentDidMount(): void {
    this._loadRootFolder().catch(err => console.error(err));
  }

  public componentDidUpdate(prevProps: IDocumentLibraryViewProps): void {
    if (prevProps.libraryName !== this.props.libraryName) {
      this.setState({
        currentFolderUrl: '', breadcrumbs: [], folders: [], files: [],
        isLoading: true, error: undefined, rootUrl: undefined,
        showNewFolder: false, newFolderName: '', folderError: undefined,
        uploadError: undefined,
      }, () => this._loadRootFolder().catch(err => console.error(err)));
    }
  }

  // ── Load root folder ─────────────────────────────────────────────────────────

  private async _loadRootFolder(): Promise<void> {
    const { spHttpClient, webUrl, libraryName } = this.props;
    try {
      const res: SPHttpClientResponse = await spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${libraryName}')/rootfolder?$select=ServerRelativeUrl`,
        SPHttpClient.configurations.v1
      );
      if (!res.ok) {
        throw new Error(res.status === 404
          ? `Library "${libraryName}" not found. Try removing and re-adding this project.`
          : `HTTP ${res.status}`);
      }
      const data = await res.json();
      const rootUrl: string = data.ServerRelativeUrl;
      this.setState(
        { rootUrl, currentFolderUrl: rootUrl, breadcrumbs: [{ name: libraryName, url: rootUrl }] },
        () => this._loadFolder(rootUrl).catch(err => console.error(err))
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      this.setState({ isLoading: false, error: msg });
    }
  }

  // ── Load folder contents ─────────────────────────────────────────────────────

  private async _loadFolder(folderUrl: string): Promise<void> {
    const { spHttpClient, webUrl } = this.props;
    this.setState({ isLoading: true, error: undefined });
    try {
      const encoded = encodeURIComponent(folderUrl);
      const [foldersRes, filesRes] = await Promise.all([
        spHttpClient.get(
          `${webUrl}/_api/web/getFolderByServerRelativeUrl('${encoded}')/Folders` +
          `?$select=Name,ServerRelativeUrl,ItemCount,TimeLastModified&$orderby=Name`,
          SPHttpClient.configurations.v1
        ),
        spHttpClient.get(
          `${webUrl}/_api/web/getFolderByServerRelativeUrl('${encoded}')/Files` +
          `?$select=Name,ServerRelativeUrl,Length,TimeLastModified,UniqueId&$orderby=Name`,
          SPHttpClient.configurations.v1
        ),
      ]);

      if (!foldersRes.ok) throw new Error(`HTTP ${foldersRes.status}`);
      if (!filesRes.ok)   throw new Error(`HTTP ${filesRes.status}`);

      const foldersData = await foldersRes.json();
      const filesData   = await filesRes.json();

      const folders: IFolder[] = (foldersData.value || [])
        .filter((f: { Name: string }) => f.Name !== 'Forms')
        .map((f: { Name: string; ServerRelativeUrl: string; ItemCount: number; TimeLastModified: string }) => ({
          name: f.Name,
          serverRelativeUrl: f.ServerRelativeUrl,
          itemCount: f.ItemCount,
          timeLastModified: f.TimeLastModified,
        }));

      const files: IFile[] = (filesData.value || []).map((f: {
        Name: string; ServerRelativeUrl: string; Length: number; TimeLastModified: string; UniqueId: string;
      }) => ({
        name: f.Name,
        serverRelativeUrl: f.ServerRelativeUrl,
        length: f.Length,
        timeLastModified: f.TimeLastModified,
        uniqueId: f.UniqueId,
      }));

      this.setState({ folders, files, isLoading: false });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      this.setState({ isLoading: false, error: `Could not load folder: ${msg}` });
    }
  }

  // ── Create new folder ────────────────────────────────────────────────────────

  private _handleCreateFolder = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    const { newFolderName, currentFolderUrl } = this.state;
    const name = newFolderName.trim();
    if (!name) return;

    this.setState({ isCreatingFolder: true, folderError: undefined });
    try {
      const encoded = encodeURIComponent(currentFolderUrl);
      const res: SPHttpClientResponse = await this.props.spHttpClient.post(
        `${this.props.webUrl}/_api/web/getFolderByServerRelativeUrl('${encoded}')/folders/add('${encodeURIComponent(name)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body: JSON.stringify({}),
        }
      );
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err['odata.error']?.message?.value || `HTTP ${res.status}`);
      }
      this.setState({ showNewFolder: false, newFolderName: '', isCreatingFolder: false });
      await this._loadFolder(currentFolderUrl);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      this.setState({ isCreatingFolder: false, folderError: msg });
    }
  }

  // ── Upload file ──────────────────────────────────────────────────────────────

  private _handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;

    const file = fileList[0];
    this.setState({ isUploading: true, uploadError: undefined, uploadingFileName: file.name });

    try {
      const arrayBuffer = await this._readFileAsArrayBuffer(file);
      const encoded = encodeURIComponent(this.state.currentFolderUrl);

      const res: SPHttpClientResponse = await this.props.spHttpClient.post(
        `${this.props.webUrl}/_api/web/getFolderByServerRelativeUrl('${encoded}')/Files/add(url='${encodeURIComponent(file.name)}',overwrite=true)`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body: arrayBuffer,
        }
      );

      if (!res.ok) throw new Error(`HTTP ${res.status}`);

      this.setState({ isUploading: false, uploadingFileName: undefined });
      // Reset file input
      if (this._fileInputRef.current) this._fileInputRef.current.value = '';
      await this._loadFolder(this.state.currentFolderUrl);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      this.setState({ isUploading: false, uploadError: msg, uploadingFileName: undefined });
    }
  }

  private _readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = e => resolve(e.target?.result as ArrayBuffer);
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });
  }

  // ── Navigation ───────────────────────────────────────────────────────────────

  private _openFolder = (folder: IFolder): void => {
    this.setState(prev => ({
      currentFolderUrl: folder.serverRelativeUrl,
      breadcrumbs: [...prev.breadcrumbs, { name: folder.name, url: folder.serverRelativeUrl }],
      showNewFolder: false, newFolderName: '', folderError: undefined, uploadError: undefined,
    }), () => this._loadFolder(folder.serverRelativeUrl).catch(err => console.error(err)));
  }

  private _navigateTo = (url: string, index: number): void => {
    this.setState(prev => ({
      currentFolderUrl: url,
      breadcrumbs: prev.breadcrumbs.slice(0, index + 1),
      showNewFolder: false, newFolderName: '', folderError: undefined, uploadError: undefined,
    }), () => this._loadFolder(url).catch(err => console.error(err)));
  }

  // ── Helpers ──────────────────────────────────────────────────────────────────

  private _formatSize(bytes: number): string {
    if (!bytes) return '—';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  private _formatDate(isoDate: string): string {
    if (!isoDate) return '—';
    return new Date(isoDate).toLocaleDateString();
  }

  private _getFileIcon(name: string): string {
    const ext = name.split('.').pop()?.toLowerCase() || '';
    const icons: { [key: string]: string } = {
      pdf: '📄', docx: '📝', doc: '📝', xlsx: '📊', xls: '📊',
      pptx: '📑', ppt: '📑', jpg: '🖼️', jpeg: '🖼️', png: '🖼️',
      mp4: '🎬', zip: '🗜️', txt: '📃',
    };
    return icons[ext] || '📄';
  }

  // ── Render ───────────────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const {
      folders, files, isLoading, error, breadcrumbs,
      showNewFolder, newFolderName, isCreatingFolder, folderError,
      isUploading, uploadError, uploadingFileName,
    } = this.state;

    const isEmpty = !isLoading && !error && folders.length === 0 && files.length === 0;

    return (
      <div className={styles.root}>

        {/* Breadcrumb */}
        <div className={styles.breadcrumb}>
          <div className={styles.breadcrumbPath}>
            {breadcrumbs.map((crumb, i) => (
              <span key={i} className={styles.breadcrumbItem}>
                {i > 0 && <span className={styles.breadcrumbSep}>›</span>}
                <span
                  className={i === breadcrumbs.length - 1 ? styles.breadcrumbCurrent : styles.breadcrumbLink}
                  onClick={i < breadcrumbs.length - 1 ? () => this._navigateTo(crumb.url, i) : undefined}
                  role={i < breadcrumbs.length - 1 ? 'button' : undefined}
                  tabIndex={i < breadcrumbs.length - 1 ? 0 : undefined}
                  onKeyDown={i < breadcrumbs.length - 1 ? ev => ev.key === 'Enter' && this._navigateTo(crumb.url, i) : undefined}
                >
                  {crumb.name}
                </span>
              </span>
            ))}
          </div>

          {/* Toolbar actions */}
          {!isLoading && !error && (
            <div className={styles.toolbar}>
              <button
                className={styles.toolbarBtn}
                onClick={() => this.setState(prev => ({ showNewFolder: !prev.showNewFolder, folderError: undefined, newFolderName: '' }))}
                title="New folder"
              >
                📁 New Folder
              </button>
              <button
                className={`${styles.toolbarBtn} ${styles.toolbarBtnPrimary}`}
                onClick={() => this._fileInputRef.current?.click()}
                disabled={isUploading}
                title="Upload file"
              >
                {isUploading ? `Uploading ${uploadingFileName || ''}…` : '⬆ Upload'}
              </button>
              {/* Hidden file input */}
              <input
                ref={this._fileInputRef}
                type="file"
                style={{ display: 'none' }}
                onChange={this._handleFileChange}
              />
            </div>
          )}
        </div>

        {/* New folder inline form */}
        {showNewFolder && (
          <form className={styles.newFolderForm} onSubmit={this._handleCreateFolder}>
            <span className={styles.newFolderIcon}>📁</span>
            <input
              className={styles.newFolderInput}
              type="text"
              placeholder="Folder name"
              value={newFolderName}
              onChange={e => this.setState({ newFolderName: e.target.value, folderError: undefined })}
              autoFocus
            />
            <button
              type="submit"
              className={styles.newFolderSubmit}
              disabled={!newFolderName.trim() || isCreatingFolder}
            >
              {isCreatingFolder ? 'Creating…' : 'Create'}
            </button>
            <button
              type="button"
              className={styles.newFolderCancel}
              onClick={() => this.setState({ showNewFolder: false, newFolderName: '', folderError: undefined })}
            >
              Cancel
            </button>
            {folderError && <span className={styles.inlineError}>{folderError}</span>}
          </form>
        )}

        {/* Upload error */}
        {uploadError && (
          <div className={styles.uploadError}>Upload failed: {uploadError}</div>
        )}

        {/* Loading */}
        {isLoading && <div className={styles.state}>Loading…</div>}

        {/* Error */}
        {error && <div className={styles.stateError}>{error}</div>}

        {/* Empty */}
        {isEmpty && !showNewFolder && (
          <div className={styles.state}>
            <div className={styles.emptyIcon}>📂</div>
            <div>This folder is empty</div>
          </div>
        )}

        {/* Contents table */}
        {!isLoading && !error && (folders.length > 0 || files.length > 0) && (
          <div className={styles.table}>
            <div className={styles.tableHeader}>
              <div className={styles.colName}>Name</div>
              <div className={styles.colMeta}>Modified</div>
              <div className={styles.colSize}>Size</div>
            </div>

            {folders.map((folder, i) => (
              <div
                key={i}
                className={`${styles.tableRow} ${styles.tableRowFolder}`}
                onClick={() => this._openFolder(folder)}
                role="button"
                tabIndex={0}
                onKeyDown={e => e.key === 'Enter' && this._openFolder(folder)}
              >
                <div className={styles.colName}>
                  <span className={styles.itemIcon}>📁</span>
                  <span className={styles.itemName}>{folder.name}</span>
                  {folder.itemCount > 0 && <span className={styles.itemCount}>{folder.itemCount}</span>}
                </div>
                <div className={styles.colMeta}>{this._formatDate(folder.timeLastModified)}</div>
                <div className={styles.colSize}>—</div>
              </div>
            ))}

            {files.map((file, i) => (
              <div key={i} className={styles.tableRow}>
                <div className={styles.colName}>
                  <span className={styles.itemIcon}>{this._getFileIcon(file.name)}</span>
                  <a
                    className={styles.fileLink}
                    href={`${this.props.webUrl}${file.serverRelativeUrl}`}
                    target="_blank"
                    rel="noreferrer"
                  >
                    {file.name}
                  </a>
                </div>
                <div className={styles.colMeta}>{this._formatDate(file.timeLastModified)}</div>
                <div className={styles.colSize}>{this._formatSize(file.length)}</div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }
}
