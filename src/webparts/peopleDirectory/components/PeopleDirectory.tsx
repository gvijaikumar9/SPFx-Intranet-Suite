import * as React from 'react';
import { useMemo, useState, useEffect } from 'react';
import styles from './PeopleDirectory.module.scss';
import { IPeopleDirectoryProps } from './IPeopleDirectoryProps';
import { IPerson } from '../models/IPerson';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

function initials(name: string): string {
  const parts = name.trim().split(/\s+/);
  const first = parts[0] ? parts[0][0] : '';
  const last = parts.length > 1 ? parts[parts.length - 1][0] : '';
  return (first + last).toUpperCase() || '?';
}

const Avatar: React.FC<{ person: IPerson; webUrl: string; showPhoto: boolean }> = (p) => {
  const [failed, setFailed] = useState(false);
  const photoUrl = p.person.email
    ? `${p.webUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(p.person.email)}`
    : '';
  if (p.showPhoto && photoUrl && !failed) {
    return <img className={styles.avatar} src={photoUrl} alt="" aria-hidden="true" onError={() => setFailed(true)} />;
  }
  return <span className={styles.avatar} aria-hidden="true">{initials(p.person.name)}</span>;
};

const PeopleDirectory: React.FC<IPeopleDirectoryProps> = (props) => {
  const { title, layout, showTitle, frameStyle, showPhoto, webUrl, isDemo, accent, demoPeople, onSearch } = props;
  const [query, setQuery] = useState('');
  const [results, setResults] = useState<IPerson[]>([]);
  const [searching, setSearching] = useState(false);

  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  useEffect(() => {
    const q = query.trim();
    if (q.length < 2) { setResults([]); setSearching(false); return undefined; }
    setSearching(true);
    const handle = window.setTimeout(() => {
      onSearch(q).then((r) => { setResults(r); setSearching(false); }).catch(() => { setResults([]); setSearching(false); });
    }, 300);
    return () => window.clearTimeout(handle);
  }, [query, onSearch]);

  const showing: IPerson[] = query.trim().length >= 2 ? results : demoPeople;

  const card = (person: IPerson, key: string): React.ReactElement => (
    <li key={key} className={styles.card}>
      <Avatar person={person} webUrl={webUrl} showPhoto={showPhoto} />
      <div className={styles.info}>
        <p className={styles.name}>{person.name}</p>
        {person.title && <p className={styles.role}>{person.title}{person.department ? ` · ${person.department}` : ''}</p>}
        {person.email && <a className={styles.email} href={`mailto:${person.email}`}>{person.email}</a>}
      </div>
    </li>
  );

  return (
    <section className={`${styles.people} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Find a colleague'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Search returns sample people until you deploy to a tenant.</p>}

      <input
        className={styles.search}
        type="search"
        value={query}
        placeholder="Search by name"
        aria-label="Search people"
        onChange={(e) => setQuery(e.target.value)}
      />

      {searching && <p className={styles.muted}>Searching...</p>}
      {!searching && query.trim().length >= 2 && showing.length === 0 && (
        <p className={styles.muted}>No people found.</p>
      )}
      {!searching && showing.length > 0 && (
        <ul className={styles.grid}>
          {showing.map((person, i) => card(person, person.login || `p${i}`))}
        </ul>
      )}
    </section>
  );
};

export default PeopleDirectory;
