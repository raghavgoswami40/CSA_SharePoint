/**
 * DocumentLibraryView.tsx
 *
 * Redesigned document library component adapted from file-upload pattern.
 * White background, list/grid toggle, search, sort, drag-and-drop upload,
 * bulk select/delete/download — all wired to the SharePoint REST API.
 */

import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  FileIcon, FileTextIcon, FileSpreadsheetIcon, FileArchiveIcon,
  ImageIcon, VideoIcon, HeadphonesIcon,
  UploadCloudIcon, UploadIcon, Trash2Icon, DownloadIcon,
  ExternalLinkIcon, SearchIcon, ListIcon, GridIcon,
  SortAscIcon, SortDescIcon, FolderIcon, FolderPlusIcon,
  AlertCircleIcon, CheckIcon,
} from 'lucide-react';

// ── Props ─────────────────────────────────────────────────────────────────────

export interface IDocumentLibraryViewProps {
  libraryName: string;
  spHttpClient: SPHttpClient;
  webUrl: string;
}

// ── Data shapes ───────────────────────────────────────────────────────────────

interface ILibFolder {
  id: string;
  name: string;
  serverRelativeUrl: string;
  itemCount: number;
  modified: string;
}

interface ILibFile {
  id: string;
  name: string;
  serverRelativeUrl: string;
  length: number;
  modified: string;
  ext: string;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

const formatBytes = (bytes: number): string => {
  if (!bytes) return '—';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

const formatDate = (iso: string): string => {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString();
};

const getExt = (name: string): string =>
  name.indexOf('.') !== -1 ? name.split('.').pop()!.toLowerCase() : '';

const FileTypeIcon: React.FC<{ ext: string; size?: number }> = ({ ext, size = 14 }) => {
  const s = { opacity: 0.6 };
  if (['pdf', 'doc', 'docx', 'txt', 'md'].indexOf(ext) !== -1)
    return <FileTextIcon size={size} style={s} />;
  if (['xls', 'xlsx', 'csv'].indexOf(ext) !== -1)
    return <FileSpreadsheetIcon size={size} style={s} />;
  if (['zip', 'rar', '7z', 'tar'].indexOf(ext) !== -1)
    return <FileArchiveIcon size={size} style={s} />;
  if (['mp4', 'mov', 'webm', 'mkv'].indexOf(ext) !== -1)
    return <VideoIcon size={size} style={s} />;
  if (['mp3', 'wav', 'flac', 'm4a'].indexOf(ext) !== -1)
    return <HeadphonesIcon size={size} style={s} />;
  if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'].indexOf(ext) !== -1)
    return <ImageIcon size={size} style={s} />;
  return <FileIcon size={size} style={s} />;
};

const getFormDigest = async (webUrl: string): Promise<string> => {
  const res = await fetch(`${webUrl}/_api/contextinfo`, {
    method: 'POST',
    headers: { 'Accept': 'application/json;odata=nometadata' },
    credentials: 'same-origin',
  });
  const data = await res.json();
  return data.FormDigestValue;
};

// ── Styles ────────────────────────────────────────────────────────────────────

const S = {
  root: {
    background: '#fff',
    borderRadius: 12,
    border: '1px solid #e5e7eb',
    fontFamily: "'Segoe UI', sans-serif",
    overflow: 'hidden',
    width: '100%',
  } as React.CSSProperties,

  toolbar: {
    display: 'flex',
    flexWrap: 'wrap' as const,
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: 8,
    padding: '12px 16px',
    borderBottom: '1px solid #f3f4f6',
    background: '#fafafa',
  } as React.CSSProperties,

  btn: (variant: 'default' | 'outline' | 'ghost' | 'danger' = 'outline', size: 'sm' | 'icon' = 'sm'): React.CSSProperties => ({
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    gap: 5, borderRadius: 7, fontWeight: 500, cursor: 'pointer',
    fontSize: 12, transition: 'all 0.15s', whiteSpace: 'nowrap' as const,
    border: variant === 'ghost' ? 'none' : '1px solid #d1d5db',
    padding: size === 'icon' ? '0' : '5px 10px',
    width: size === 'icon' ? 32 : undefined,
    height: size === 'icon' ? 32 : undefined,
    background: variant === 'default' ? '#111827' : 'transparent',
    color: variant === 'default' ? '#fff' : variant === 'danger' ? '#ef4444' : '#374151',
  }),

  searchBox: {
    position: 'relative' as const,
    display: 'flex', alignItems: 'center',
  } as React.CSSProperties,

  input: {
    border: '1px solid #d1d5db', borderRadius: 7, outline: 'none',
    fontSize: 12, background: '#fff', color: '#111827',
    padding: '5px 10px 5px 28px', height: 32,
  } as React.CSSProperties,

  dropZone: (active: boolean): React.CSSProperties => ({
    margin: '12px 16px',
    border: `1.5px dashed ${active ? '#6366f1' : '#d1d5db'}`,
    borderRadius: 10,
    padding: '12px 16px',
    background: active ? '#eef2ff' : '#fff',
    transition: 'all 0.2s',
    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
    flexWrap: 'wrap' as const, gap: 8,
  }),

  breadcrumb: {
    display: 'flex', alignItems: 'center',
    gap: 4, padding: '10px 16px',
    fontSize: 13, borderBottom: '1px solid #f3f4f6',
    flexWrap: 'wrap' as const,
  } as React.CSSProperties,

  tableHeader: {
    display: 'grid',
    gridTemplateColumns: '32px 1fr 90px 80px 90px 140px',
    padding: '7px 16px',
    background: '#f9fafb',
    borderBottom: '1px solid #f3f4f6',
    fontSize: 11, fontWeight: 700, letterSpacing: '0.05em',
    textTransform: 'uppercase' as const, color: '#9ca3af',
  } as React.CSSProperties,

  tableRow: (selected: boolean): React.CSSProperties => ({
    display: 'grid',
    gridTemplateColumns: '32px 1fr 90px 80px 90px 140px',
    padding: '8px 16px',
    alignItems: 'center',
    borderBottom: '1px solid #f9fafb',
    background: selected ? '#fef9ec' : '#fff',
    transition: 'background 0.1s',
  }),

  folderRow: (selected: boolean): React.CSSProperties => ({
    display: 'grid',
    gridTemplateColumns: '32px 1fr 90px 80px 90px 140px',
    padding: '8px 16px',
    alignItems: 'center',
    borderBottom: '1px solid #f9fafb',
    background: selected ? '#fef9ec' : '#fff',
    cursor: 'pointer',
    transition: 'background 0.1s',
  }),

  bulkBar: {
    display: 'flex', alignItems: 'center',
    justifyContent: 'space-between', gap: 8,
    padding: '8px 16px',
    background: '#f9fafb',
    borderBottom: '1px solid #f3f4f6',
    fontSize: 12,
  } as React.CSSProperties,

  emptyState: {
    display: 'flex', flexDirection: 'column' as const,
    alignItems: 'center', justifyContent: 'center',
    padding: '48px 24px', gap: 8, color: '#9ca3af', fontSize: 13,
  } as React.CSSProperties,

  gridContainer: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(150px, 1fr))',
    gap: 12, padding: 16,
  } as React.CSSProperties,

  gridCard: (selected: boolean): React.CSSProperties => ({
    border: `1px solid ${selected ? '#fbbf24' : '#e5e7eb'}`,
    borderRadius: 8, overflow: 'hidden',
    background: selected ? '#fef9ec' : '#fff',
    display: 'flex', flexDirection: 'column' as const,
  }),

  newFolderRow: {
    display: 'flex', alignItems: 'center', gap: 8,
    padding: '8px 16px',
    background: '#fafafa',
    borderBottom: '1px solid #f3f4f6',
  } as React.CSSProperties,
};

// ── Component ─────────────────────────────────────────────────────────────────

const DocumentLibraryView: React.FC<IDocumentLibraryViewProps> = ({ libraryName, spHttpClient, webUrl }) => {

  // Navigation
  const [folderUrl, setFolderUrl]   = React.useState('');
  const [breadcrumbs, setBreadcrumbs] = React.useState<{ name: string; url: string }[]>([]);

  // Content
  const [folders, setFolders] = React.useState<ILibFolder[]>([]);
  const [files, setFiles]     = React.useState<ILibFile[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError]     = React.useState('');

  // UI
  const [view, setView]       = React.useState<'list' | 'grid'>('list');
  const [query, setQuery]     = React.useState('');
  const [sortBy, setSortBy]   = React.useState<'name' | 'type' | 'size' | 'modified'>('name');
  const [sortDir, setSortDir] = React.useState<'asc' | 'desc'>('asc');
  const [selected, setSelected] = React.useState<Set<string>>(new Set());
  const [isDragging, setIsDragging] = React.useState(false);

  // Upload
  const [isUploading, setIsUploading] = React.useState(false);
  const [uploadError, setUploadError] = React.useState('');
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // New folder
  const [showNewFolder, setShowNewFolder] = React.useState(false);
  const [newFolderName, setNewFolderName] = React.useState('');
  const [isCreatingFolder, setIsCreatingFolder] = React.useState(false);

  // Misc
  const [copied, setCopied] = React.useState('');

  // ── Load root ────────────────────────────────────────────────────────────

  React.useEffect(() => {
    setIsLoading(true);
    setError('');
    setFolders([]);
    setFiles([]);
    setFolderUrl('');
    setBreadcrumbs([]);
    setSelected(new Set());

    spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${libraryName}')/rootfolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1
    ).then(async (res: SPHttpClientResponse) => {
      if (!res.ok) {
        setError(res.status === 404 ? `Library "${libraryName}" not found.` : `HTTP ${res.status}`);
        setIsLoading(false);
        return;
      }
      const data = await res.json();
      const url: string = data.ServerRelativeUrl;
      setFolderUrl(url);
      setBreadcrumbs([{ name: libraryName, url }]);
    }).catch((err: Error) => {
      setError(err.message);
      setIsLoading(false);
    });
  }, [libraryName]);

  // ── Load folder when folderUrl changes ───────────────────────────────────

  const loadFolder = (url: string): void => {
    setIsLoading(true);
    setError('');
    const enc = encodeURIComponent(url);

    Promise.all([
      spHttpClient.get(
        `${webUrl}/_api/web/getFolderByServerRelativeUrl('${enc}')/Folders?$select=Name,ServerRelativeUrl,ItemCount,TimeLastModified&$orderby=Name`,
        SPHttpClient.configurations.v1
      ),
      spHttpClient.get(
        `${webUrl}/_api/web/getFolderByServerRelativeUrl('${enc}')/Files?$select=Name,ServerRelativeUrl,Length,TimeLastModified,UniqueId&$orderby=Name`,
        SPHttpClient.configurations.v1
      ),
    ]).then(async ([fRes, fiRes]) => {
      const fData  = await fRes.json();
      const fiData = await fiRes.json();

      setFolders((fData.value || [])
        .filter((f: { Name: string }) => f.Name !== 'Forms')
        .map((f: { Name: string; ServerRelativeUrl: string; ItemCount: number; TimeLastModified: string }) => ({
          id: f.ServerRelativeUrl,
          name: f.Name,
          serverRelativeUrl: f.ServerRelativeUrl,
          itemCount: f.ItemCount,
          modified: f.TimeLastModified,
        })));

      setFiles((fiData.value || []).map((f: {
        UniqueId: string; Name: string; ServerRelativeUrl: string; Length: number; TimeLastModified: string;
      }) => ({
        id: f.UniqueId,
        name: f.Name,
        serverRelativeUrl: f.ServerRelativeUrl,
        length: f.Length,
        modified: f.TimeLastModified,
        ext: getExt(f.Name),
      })));

      setIsLoading(false);
    }).catch((err: Error) => {
      setError(err.message);
      setIsLoading(false);
    });
  };

  // ── Effect: reload when folderUrl changes ────────────────────────────────

  React.useEffect(() => {
    if (!folderUrl) return;
    loadFolder(folderUrl);
  }, [folderUrl]);

  // ── Navigation ───────────────────────────────────────────────────────────

  const openFolder = (folder: ILibFolder): void => {
    setBreadcrumbs(prev => [...prev, { name: folder.name, url: folder.serverRelativeUrl }]);
    setFolderUrl(folder.serverRelativeUrl);
    setSelected(new Set());
    setQuery('');
  };

  const navigateTo = (url: string, index: number): void => {
    setBreadcrumbs(prev => prev.slice(0, index + 1));
    setFolderUrl(url);
    setSelected(new Set());
    setQuery('');
  };

  // ── Upload ───────────────────────────────────────────────────────────────

  const uploadFile = async (file: File): Promise<void> => {
    setIsUploading(true);
    setUploadError('');
    try {
      const digest = await getFormDigest(webUrl);
      const buf = await file.arrayBuffer();
      const enc = encodeURIComponent(folderUrl);
      const res = await fetch(
        `${webUrl}/_api/web/getFolderByServerRelativeUrl('${enc}')/Files/add(url='${encodeURIComponent(file.name)}',overwrite=true)`,
        { method: 'POST', headers: { 'Accept': 'application/json;odata=nometadata', 'X-RequestDigest': digest }, body: buf, credentials: 'same-origin' }
      );
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      loadFolder(folderUrl);
    } catch (err) {
      setUploadError(err instanceof Error ? err.message : 'Upload failed');
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleDrop = (e: React.DragEvent): void => {
    e.preventDefault();
    setIsDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) uploadFile(f).catch(console.error);
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const f = e.target.files?.[0];
    if (f) uploadFile(f).catch(console.error);
  };

  // ── New folder ───────────────────────────────────────────────────────────

  const createFolder = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    const name = newFolderName.trim();
    if (!name) return;
    setIsCreatingFolder(true);
    try {
      const enc = encodeURIComponent(folderUrl);
      const res: SPHttpClientResponse = await spHttpClient.post(
        `${webUrl}/_api/web/getFolderByServerRelativeUrl('${enc}')/folders/add('${encodeURIComponent(name)}')`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata', 'Content-type': 'application/json;odata=nometadata', 'odata-version': '' }, body: JSON.stringify({}) }
      );
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setShowNewFolder(false);
      setNewFolderName('');
      loadFolder(folderUrl);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to create folder');
    } finally {
      setIsCreatingFolder(false);
    }
  };

  // ── Delete ───────────────────────────────────────────────────────────────

  const deleteFile = async (file: ILibFile): Promise<void> => {
    try {
      const digest = await getFormDigest(webUrl);
      await fetch(`${webUrl}/_api/web/getFileByServerRelativeUrl('${encodeURIComponent(file.serverRelativeUrl)}')`, {
        method: 'POST',
        headers: { 'Accept': 'application/json;odata=nometadata', 'X-RequestDigest': digest, 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' },
        credentials: 'same-origin',
      });
      loadFolder(folderUrl);
    } catch (err) { console.error(err); }
  };

  const deleteSelected = async (): Promise<void> => {
    const toDelete = files.filter(f => selected.has(f.id));
    for (const f of toDelete) await deleteFile(f);
    setSelected(new Set());
  };

  // ── Copy / download ──────────────────────────────────────────────────────

  const copyLink = async (file: ILibFile): Promise<void> => {
    const url = `${webUrl}${file.serverRelativeUrl}`;
    try {
      await navigator.clipboard.writeText(url);
      setCopied(file.id);
      setTimeout(() => setCopied(''), 1200);
    } catch { /* noop */ }
  };

  const downloadFile = (file: ILibFile): void => {
    window.open(`${webUrl}${file.serverRelativeUrl}`, '_blank', 'noopener,noreferrer');
  };

  const downloadSelected = (): void => {
    files.filter(f => selected.has(f.id)).forEach(downloadFile);
  };

  // ── Selection ────────────────────────────────────────────────────────────

  const toggleOne = (id: string): void => {
    setSelected(prev => { const n = new Set(prev); if (n.has(id)) { n.delete(id); } else { n.add(id); } return n; });
  };

  // ── Filter + sort ────────────────────────────────────────────────────────

  const filteredFolders = React.useMemo(() => {
    const q = query.toLowerCase();
    return q ? folders.filter(f => f.name.toLowerCase().indexOf(q) !== -1) : folders;
  }, [folders, query]);

  const filteredFiles = React.useMemo(() => {
    const q = query.toLowerCase();
    const base = q ? files.filter(f => f.name.toLowerCase().indexOf(q) !== -1 || f.ext.indexOf(q) !== -1) : files;
    return [...base].sort((a, b) => {
      let cmp = 0;
      if (sortBy === 'name') cmp = a.name.localeCompare(b.name);
      else if (sortBy === 'type') cmp = a.ext.localeCompare(b.ext);
      else if (sortBy === 'size') cmp = a.length - b.length;
      else cmp = new Date(a.modified).getTime() - new Date(b.modified).getTime();
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }, [files, query, sortBy, sortDir]);

  const totalSize = React.useMemo(() => files.reduce((acc, f) => acc + f.length, 0), [files]);
  const noneSelected = selected.size === 0;
  const allSelected  = filteredFiles.length > 0 && filteredFiles.every(f => selected.has(f.id));

  const toggleAll = (): void => {
    const allIds = filteredFiles.map(f => f.id);
    const everySelected = allIds.every(id => selected.has(id));
    setSelected(everySelected ? new Set() : new Set(allIds));
  };

  // ── Render ───────────────────────────────────────────────────────────────

  return (
    <div style={S.root}>

      {/* Breadcrumb */}
      <div style={S.breadcrumb}>
        {breadcrumbs.map((c, i) => (
          <React.Fragment key={c.url}>
            {i > 0 && <span style={{ color: '#d1d5db', margin: '0 2px' }}>›</span>}
            <span
              onClick={i < breadcrumbs.length - 1 ? () => navigateTo(c.url, i) : undefined}
              style={{
                cursor: i < breadcrumbs.length - 1 ? 'pointer' : 'default',
                color: i < breadcrumbs.length - 1 ? '#6366f1' : '#111827',
                fontWeight: i === breadcrumbs.length - 1 ? 600 : 400,
                textDecoration: i < breadcrumbs.length - 1 ? 'underline' : 'none',
                textUnderlineOffset: 2,
              }}
            >
              {c.name}
            </span>
          </React.Fragment>
        ))}
      </div>

      {/* Toolbar */}
      <div style={S.toolbar}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <span style={{ fontSize: 13, fontWeight: 500, color: '#111827' }}>
            Files <span style={{ color: '#9ca3af' }}>({files.length})</span>
          </span>
          <span style={{ fontSize: 11, color: '#9ca3af' }}>{formatBytes(totalSize)}</span>
        </div>

        <div style={{ display: 'flex', flexWrap: 'wrap' as const, alignItems: 'center', gap: 6 }}>
          {/* Search */}
          <div style={S.searchBox}>
            <SearchIcon size={12} style={{ position: 'absolute', left: 8, color: '#9ca3af' }} />
            <input
              type="text"
              placeholder="Search files…"
              value={query}
              onChange={e => setQuery(e.target.value)}
              style={{ ...S.input, width: 180 }}
            />
          </div>

          {/* Sort */}
          <select
            value={sortBy}
            onChange={e => setSortBy(e.target.value as typeof sortBy)}
            style={{ ...S.input, paddingLeft: 10, width: 'auto', cursor: 'pointer' }}
          >
            <option value="name">Name</option>
            <option value="type">Type</option>
            <option value="size">Size</option>
            <option value="modified">Modified</option>
          </select>
          <button
            onClick={() => setSortDir(d => d === 'asc' ? 'desc' : 'asc')}
            style={S.btn('outline', 'icon')}
            title="Toggle sort direction"
          >
            {sortDir === 'asc' ? <SortAscIcon size={14} /> : <SortDescIcon size={14} />}
          </button>

          {/* View toggle */}
          <button
            onClick={() => setView('list')}
            style={{ ...S.btn(view === 'list' ? 'default' : 'outline', 'icon') }}
            title="List view"
          >
            <ListIcon size={14} />
          </button>
          <button
            onClick={() => setView('grid')}
            style={{ ...S.btn(view === 'grid' ? 'default' : 'outline', 'icon') }}
            title="Grid view"
          >
            <GridIcon size={14} />
          </button>

          {/* Actions */}
          <button
            onClick={() => setShowNewFolder(f => !f)}
            style={S.btn('outline', 'sm')}
          >
            <FolderPlusIcon size={13} /> New Folder
          </button>
          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={isUploading}
            style={S.btn('outline', 'sm')}
          >
            <UploadCloudIcon size={13} />
            {isUploading ? 'Uploading…' : 'Upload'}
          </button>
          <input ref={fileInputRef} type="file" style={{ display: 'none' }} onChange={handleFileInput} />
        </div>
      </div>

      {/* New folder form */}
      {showNewFolder && (
        <form onSubmit={createFolder} style={S.newFolderRow}>
          <FolderIcon size={16} style={{ color: '#9ca3af', flexShrink: 0 }} />
          <input
            autoFocus
            type="text"
            placeholder="Folder name"
            value={newFolderName}
            onChange={e => setNewFolderName(e.target.value)}
            style={{ ...S.input, flex: 1 }}
          />
          <button type="submit" disabled={!newFolderName.trim() || isCreatingFolder} style={S.btn('default', 'sm')}>
            {isCreatingFolder ? 'Creating…' : 'Create'}
          </button>
          <button type="button" onClick={() => { setShowNewFolder(false); setNewFolderName(''); }} style={S.btn('outline', 'sm')}>
            Cancel
          </button>
        </form>
      )}

      {/* Drop zone */}
      <div
        onDragEnter={e => { e.preventDefault(); setIsDragging(true); }}
        onDragLeave={e => { e.preventDefault(); if (!e.currentTarget.contains(e.relatedTarget as Node)) setIsDragging(false); }}
        onDragOver={e => e.preventDefault()}
        onDrop={handleDrop}
        style={S.dropZone(isDragging)}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 36, height: 36, borderRadius: '50%', border: '1px solid #e5e7eb', background: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
            <FileIcon size={16} style={{ opacity: 0.5 }} />
          </div>
          <div>
            <div style={{ fontSize: 12, fontWeight: 500, color: '#111827' }}>Drop files to upload</div>
            <div style={{ fontSize: 11, color: '#9ca3af' }}>or use the Upload button above</div>
          </div>
        </div>
        <button onClick={() => fileInputRef.current?.click()} style={S.btn('outline', 'sm')}>
          <UploadIcon size={12} /> Select files
        </button>
      </div>

      {/* Upload error */}
      {uploadError && (
        <div style={{ padding: '6px 16px', fontSize: 12, color: '#ef4444', display: 'flex', alignItems: 'center', gap: 6, background: '#fef2f2' }}>
          <AlertCircleIcon size={12} /> {uploadError}
        </div>
      )}

      {/* Error */}
      {error && (
        <div style={{ padding: '12px 16px', fontSize: 13, color: '#ef4444', background: '#fef2f2' }}>
          {error}
        </div>
      )}

      {/* Loading */}
      {isLoading && (
        <div style={S.emptyState}>Loading…</div>
      )}

      {/* Content */}
      {!isLoading && !error && (filteredFolders.length > 0 || filteredFiles.length > 0) && (
        <>
          {/* Bulk actions */}
          <div style={S.bulkBar}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <label style={{ display: 'inline-flex', alignItems: 'center', gap: 5, cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={allSelected}
                  onChange={toggleAll}
                  style={{ accentColor: '#111827', width: 14, height: 14 }}
                />
                <span style={{ color: '#6b7280' }}>
                  {selected.size}/{filteredFiles.length} selected
                </span>
              </label>
            </div>
            <div style={{ display: 'flex', gap: 6 }}>
              <button onClick={downloadSelected} disabled={noneSelected} style={S.btn('outline', 'sm')}>
                <DownloadIcon size={12} /> Download
              </button>
              <button onClick={deleteSelected} disabled={noneSelected} style={{ ...S.btn('outline', 'sm'), color: '#ef4444', borderColor: '#fca5a5' }}>
                <Trash2Icon size={12} /> Remove
              </button>
            </div>
          </div>

          {view === 'list' ? (
            <div>
              {/* Table header */}
              <div style={S.tableHeader}>
                <div />
                <div>Name</div>
                <div>Type</div>
                <div>Size</div>
                <div>Modified</div>
                <div style={{ textAlign: 'right' }}>Actions</div>
              </div>

              {/* Folders */}
              {filteredFolders.map(folder => (
                <div
                  key={folder.id}
                  style={S.folderRow(false)}
                  onClick={() => openFolder(folder)}
                  onMouseEnter={e => (e.currentTarget.style.background = '#f9fafb')}
                  onMouseLeave={e => (e.currentTarget.style.background = '#fff')}
                >
                  <div />
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, fontWeight: 500, color: '#111827', minWidth: 0 }}>
                    <FolderIcon size={16} style={{ color: '#fbbf24', flexShrink: 0 }} />
                    <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{folder.name}</span>
                    {folder.itemCount > 0 && <span style={{ fontSize: 10, background: '#f3f4f6', color: '#6b7280', borderRadius: 10, padding: '1px 7px', flexShrink: 0 }}>{folder.itemCount}</span>}
                  </div>
                  <div style={{ fontSize: 12, color: '#9ca3af' }}>Folder</div>
                  <div style={{ fontSize: 12, color: '#9ca3af' }}>—</div>
                  <div style={{ fontSize: 12, color: '#9ca3af' }}>{formatDate(folder.modified)}</div>
                  <div />
                </div>
              ))}

              {/* Files */}
              {filteredFiles.map(file => {
                const isSel = selected.has(file.id);
                const pct   = Math.min(100, Math.round((file.length / (20 * 1024 * 1024)) * 100));
                return (
                  <div key={file.id} style={S.tableRow(isSel)}>
                    <div>
                      <input
                        type="checkbox"
                        checked={isSel}
                        onChange={() => toggleOne(file.id)}
                        style={{ accentColor: '#111827', width: 14, height: 14, cursor: 'pointer' }}
                      />
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 3, minWidth: 0 }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, fontWeight: 500, color: '#111827' }}>
                        <FileTypeIcon ext={file.ext} />
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{file.name}</span>
                      </div>
                      <div style={{ height: 4, width: 160, background: '#f3f4f6', borderRadius: 2, overflow: 'hidden' }}>
                        <div style={{ height: '100%', width: `${pct}%`, background: '#d1d5db', borderRadius: 2 }} />
                      </div>
                    </div>
                    <div style={{ fontSize: 12, color: '#6b7280', textTransform: 'uppercase', letterSpacing: '0.05em' }}>{file.ext || '—'}</div>
                    <div style={{ fontSize: 12, color: '#6b7280' }}>{formatBytes(file.length)}</div>
                    <div style={{ fontSize: 12, color: '#6b7280' }}>{formatDate(file.modified)}</div>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: 0 }}>
                      <button onClick={() => window.open(`${webUrl}${file.serverRelativeUrl}`, '_blank', 'noopener,noreferrer')} style={S.btn('ghost', 'icon')} title="Open">
                        <ExternalLinkIcon size={14} style={{ opacity: 0.6 }} />
                      </button>
                      <button onClick={() => downloadFile(file)} style={S.btn('ghost', 'icon')} title="Download">
                        <DownloadIcon size={14} style={{ opacity: 0.6 }} />
                      </button>
                      <button onClick={() => copyLink(file).catch(console.error)} style={S.btn('ghost', 'icon')} title="Copy link">
                        {copied === file.id ? <CheckIcon size={14} style={{ color: '#22c55e' }} /> : <ExternalLinkIcon size={14} style={{ opacity: 0.4 }} />}
                      </button>
                      <button onClick={() => deleteFile(file).catch(console.error)} style={{ ...S.btn('ghost', 'icon'), color: '#ef4444' }} title="Delete">
                        <Trash2Icon size={14} />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            /* Grid view */
            <div style={S.gridContainer}>
              {filteredFolders.map(folder => (
                <div
                  key={folder.id}
                  style={S.gridCard(false)}
                  onClick={() => openFolder(folder)}
                  onMouseEnter={e => (e.currentTarget.style.background = '#f9fafb')}
                  onMouseLeave={e => (e.currentTarget.style.background = '#fff')}
                >
                  <div style={{ height: 80, display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#fffbeb' }}>
                    <FolderIcon size={32} style={{ color: '#fbbf24' }} />
                  </div>
                  <div style={{ padding: 8 }}>
                    <div style={{ fontSize: 12, fontWeight: 500, color: '#111827', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{folder.name}</div>
                    <div style={{ fontSize: 11, color: '#9ca3af' }}>{folder.itemCount} items</div>
                  </div>
                </div>
              ))}

              {filteredFiles.map(file => {
                const isSel = selected.has(file.id);
                return (
                  <div key={file.id} style={S.gridCard(isSel)}>
                    <div style={{ position: 'relative', height: 80, display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#f9fafb' }}>
                      <label style={{ position: 'absolute', top: 6, left: 6, background: 'rgba(255,255,255,0.9)', borderRadius: 4, padding: '2px 4px', display: 'flex', alignItems: 'center' }}>
                        <input type="checkbox" checked={isSel} onChange={() => toggleOne(file.id)} style={{ accentColor: '#111827', width: 12, height: 12, cursor: 'pointer' }} />
                      </label>
                      <FileTypeIcon ext={file.ext} size={28} />
                    </div>
                    <div style={{ padding: 8, flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <div style={{ fontSize: 12, fontWeight: 500, color: '#111827', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={file.name}>{file.name}</div>
                      <div style={{ fontSize: 11, color: '#9ca3af' }}>{file.ext.toUpperCase()} · {formatBytes(file.length)}</div>
                      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 2, marginTop: 'auto' }}>
                        <button onClick={() => downloadFile(file)} style={S.btn('ghost', 'icon')} title="Download">
                          <DownloadIcon size={13} style={{ opacity: 0.5 }} />
                        </button>
                        <button onClick={() => deleteFile(file).catch(console.error)} style={{ ...S.btn('ghost', 'icon'), color: '#ef4444' }} title="Delete">
                          <Trash2Icon size={13} />
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </>
      )}

      {/* Empty state */}
      {!isLoading && !error && filteredFolders.length === 0 && filteredFiles.length === 0 && (
        <div style={S.emptyState}>
          <FileIcon size={32} style={{ opacity: 0.3 }} />
          <div style={{ fontWeight: 500 }}>{files.length === 0 ? 'No files yet' : 'No results match your search'}</div>
          <div style={{ fontSize: 12 }}>Drop files above or click Upload to add files</div>
        </div>
      )}

    </div>
  );
};

export default DocumentLibraryView;
