import * as React from 'react';
import { ISiteSearchResult } from './IHelloWorldProps';

const C = {
  black:  '#282928',
  white:  '#E0E1E0',
  grey:   '#7D7F7C',
  yellow: '#FFC300',
  red:    '#BD0000',
};

export interface IAddProjectFormProps {
  onAdd: (siteTitle: string, siteUrl: string) => Promise<void>;
  onSearchSites: (query: string) => Promise<ISiteSearchResult[]>;
  onCancel: () => void;
}

interface IAddProjectFormState {
  searchQuery: string;
  searchResults: ISiteSearchResult[];
  isSearching: boolean;
  selectedSite: ISiteSearchResult | undefined;
  isCreating: boolean;
  error: string | undefined;
}

export default class AddProjectForm extends React.Component<IAddProjectFormProps, IAddProjectFormState> {

  private _searchTimer: ReturnType<typeof setTimeout> | undefined;

  constructor(props: IAddProjectFormProps) {
    super(props);
    this.state = {
      searchQuery: '',
      searchResults: [],
      isSearching: false,
      selectedSite: undefined,
      isCreating: false,
      error: undefined,
    };
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
    }, 350);
  }

  private _handleSiteSelect = (site: ISiteSearchResult): void => {
    this.setState({ selectedSite: site, searchQuery: site.title, searchResults: [] });
  }

  private _handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    const { selectedSite } = this.state;
    if (!selectedSite) return;
    this.setState({ isCreating: true, error: undefined });
    try {
      await this.props.onAdd(selectedSite.title, selectedSite.url);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to add project.';
      this.setState({ isCreating: false, error: msg });
    }
  }

  public render(): React.ReactElement {
    const { searchQuery, searchResults, isSearching, selectedSite, isCreating, error } = this.state;

    return (
      // Overlay
      <div
        onClick={this.props.onCancel}
        style={{
          position: 'fixed', inset: 0,
          background: 'rgba(0,0,0,0.7)',
          zIndex: 1000,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}
      >
        {/* Modal */}
        <div
          onClick={e => e.stopPropagation()}
          style={{
            background: '#1a1b1a',
            border: `1px solid rgba(224,225,224,0.15)`,
            borderRadius: 16,
            padding: 28,
            width: 440,
            maxWidth: '90vw',
            fontFamily: "'Segoe UI', sans-serif",
          }}
        >
          <div style={{ fontSize: 16, fontWeight: 700, color: C.white, marginBottom: 4 }}>
            Link a SharePoint Site
          </div>
          <div style={{ fontSize: 12, color: C.grey, marginBottom: 20 }}>
            Search for an existing SharePoint site to add as a project
          </div>

          <form onSubmit={this._handleSubmit}>
            {/* Search input */}
            <div style={{ position: 'relative', marginBottom: 16 }}>
              <input
                type="text"
                placeholder="Search for a SharePoint site…"
                value={searchQuery}
                onChange={this._handleSearchChange}
                autoFocus
                autoComplete="off"
                style={{
                  width: '100%', boxSizing: 'border-box',
                  background: 'rgba(224,225,224,0.06)',
                  border: `1.5px solid rgba(224,225,224,0.15)`,
                  borderRadius: 8,
                  color: C.white, fontSize: 14, padding: '10px 14px',
                  outline: 'none',
                }}
              />
              {isSearching && (
                <span style={{ position: 'absolute', right: 12, top: '50%', transform: 'translateY(-50%)', color: C.grey, fontSize: 12 }}>
                  Searching…
                </span>
              )}

              {/* Dropdown results */}
              {searchResults.length > 0 && (
                <div style={{
                  position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0,
                  background: '#111', border: `1px solid rgba(224,225,224,0.15)`,
                  borderRadius: 8, overflow: 'hidden', zIndex: 10,
                  boxShadow: '0 8px 24px rgba(0,0,0,0.5)',
                }}>
                  {searchResults.map((site, i) => (
                    <div
                      key={i}
                      onClick={() => this._handleSiteSelect(site)}
                      style={{
                        padding: '10px 14px',
                        cursor: 'pointer',
                        borderBottom: i < searchResults.length - 1 ? '1px solid rgba(224,225,224,0.06)' : 'none',
                      }}
                      onMouseEnter={e => (e.currentTarget.style.background = 'rgba(255,195,0,0.08)')}
                      onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
                    >
                      <div style={{ fontSize: 13, fontWeight: 600, color: C.white, marginBottom: 2 }}>{site.title}</div>
                      <div style={{ fontSize: 11, color: C.grey, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{site.url}</div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Selected site */}
            {selectedSite && (
              <div style={{
                display: 'flex', alignItems: 'flex-start', gap: 10,
                padding: '10px 14px', marginBottom: 16,
                background: 'rgba(255,195,0,0.07)',
                border: `1px solid rgba(255,195,0,0.25)`,
                borderRadius: 8,
              }}>
                <span style={{ color: C.yellow, fontWeight: 700, marginTop: 1 }}>✓</span>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.white, marginBottom: 2 }}>{selectedSite.title}</div>
                  <div style={{ fontSize: 11, color: C.grey, wordBreak: 'break-all' }}>{selectedSite.url}</div>
                </div>
              </div>
            )}

            {/* Error */}
            {error && (
              <div style={{
                fontSize: 12, color: C.red,
                padding: '8px 12px', marginBottom: 12,
                background: 'rgba(189,0,0,0.1)',
                border: `1px solid rgba(189,0,0,0.3)`,
                borderRadius: 6,
              }}>
                {error}
              </div>
            )}

            {/* Actions */}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button
                type="button"
                onClick={this.props.onCancel}
                style={{
                  padding: '8px 18px', fontSize: 13, fontWeight: 600,
                  background: 'transparent',
                  border: `1px solid rgba(224,225,224,0.2)`,
                  borderRadius: 8, color: C.grey, cursor: 'pointer',
                }}
              >
                Cancel
              </button>
              <button
                type="submit"
                disabled={!selectedSite || isCreating}
                style={{
                  padding: '8px 18px', fontSize: 13, fontWeight: 700,
                  background: selectedSite && !isCreating ? C.yellow : 'rgba(255,195,0,0.3)',
                  border: 'none', borderRadius: 8,
                  color: C.black, cursor: selectedSite && !isCreating ? 'pointer' : 'not-allowed',
                  opacity: !selectedSite || isCreating ? 0.6 : 1,
                }}
              >
                {isCreating ? 'Adding…' : 'Add Project'}
              </button>
            </div>
          </form>
        </div>
      </div>
    );
  }
}
