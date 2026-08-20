import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import styles from './HeaderSettings.module.scss';
import { IHeaderSettings, IHeaderConfigPatch } from '../models/IHeader';

export interface IHeaderSettingsProps {
  settings: IHeaderSettings;
  accent: string;
  onSave: (patch: IHeaderConfigPatch) => Promise<boolean>;
  onClose: () => void;
}

type SaveState = 'idle' | 'saving' | 'error';

// Owner-only settings panel, opened from the gear in the header band. Edits a working copy and
// saves the changed fields to the HeaderConfig list. Only site managers ever see this.
const HeaderSettings: React.FC<IHeaderSettingsProps> = (props) => {
  const { settings, accent, onSave, onClose } = props;
  const [enabled, setEnabled] = useState<boolean>(settings.enabled);
  const [hideNative, setHideNative] = useState<boolean>(settings.hideNativeHeader);
  const [brand, setBrand] = useState<string>(settings.brandText);
  const [showCta, setShowCta] = useState<boolean>(settings.showCta);
  const [ctaText, setCtaText] = useState<string>(settings.ctaText);
  const [ctaUrl, setCtaUrl] = useState<string>(settings.ctaUrl);
  const [accentHex, setAccentHex] = useState<string>(settings.accent || accent);
  const [state, setState] = useState<SaveState>('idle');
  const ctaOn = enabled && showCta;

  const savingRef = useRef<boolean>(false);   // synchronous double-submit guard
  const mountedRef = useRef<boolean>(true);    // ignore a resolve after unmount
  const panelRef = useRef<HTMLDivElement>(null);

  const rootStyle = { ['--accent' as string]: accent } as React.CSSProperties;

  // close on Escape, move focus into the dialog on open, and mark unmounted on cleanup
  useEffect(() => {
    mountedRef.current = true;
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') { onClose(); } };
    document.addEventListener('keydown', onKey, true);
    if (panelRef.current) { panelRef.current.focus(); }
    return (): void => { mountedRef.current = false; document.removeEventListener('keydown', onKey, true); };
  }, []);

  const save = (): void => {
    if (savingRef.current) { return; }
    savingRef.current = true;
    setState('saving');
    onSave({
      enabled: enabled,
      hideNativeHeader: hideNative,
      brandText: brand,
      showCta: showCta,
      ctaText: ctaText,
      ctaUrl: ctaUrl,
      accent: accentHex
    }).then((ok) => {
      savingRef.current = false;
      if (!mountedRef.current) { return; }
      if (ok) { onClose(); } else { setState('error'); }
    }).catch(() => {
      savingRef.current = false;
      if (mountedRef.current) { setState('error'); }
    });
  };

  // clicking the dim backdrop (not the panel) closes
  const onBackdrop = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) { onClose(); }
  };

  return (
    <div className={styles.overlay} role="dialog" aria-modal="true" aria-label="Header settings" style={rootStyle} onMouseDown={onBackdrop}>
      <div className={styles.panel} ref={panelRef} tabIndex={-1}>
        <div className={styles.head}>
          <span className={styles.title}>Header settings</span>
          <button className={styles.close} type="button" onClick={onClose} aria-label="Close">×</button>
        </div>

        <div className={styles.body}>
          <label className={styles.toggle}>
            <input type="checkbox" checked={enabled} onChange={(e) => setEnabled(e.target.checked)} />
            <span>Show the custom header</span>
          </label>

          <label className={styles.toggle}>
            <input type="checkbox" checked={hideNative} disabled={!enabled} onChange={(e) => setHideNative(e.target.checked)} />
            <span>Hide SharePoint&apos;s built-in site header</span>
          </label>

          <div className={styles.row}>
            <label className={styles.field}>
              <span className={styles.lbl}>Brand text</span>
              <input className={styles.input} type="text" value={brand} disabled={!enabled} onChange={(e) => setBrand(e.target.value)} />
            </label>
            <label className={styles.field}>
              <span className={styles.lbl}>Accent</span>
              <input className={styles.color} type="color" value={/^#[0-9a-fA-F]{6}$/.test(accentHex) ? accentHex : '#0f6cbd'} disabled={!enabled} onChange={(e) => setAccentHex(e.target.value)} />
            </label>
          </div>

          <label className={styles.toggle}>
            <input type="checkbox" checked={showCta} disabled={!enabled} onChange={(e) => setShowCta(e.target.checked)} />
            <span>Show the button</span>
          </label>

          <label className={styles.field}>
            <span className={styles.lbl}>Button label</span>
            <input className={styles.input} type="text" value={ctaText} disabled={!ctaOn} placeholder="e.g. Raise a ticket" onChange={(e) => setCtaText(e.target.value)} />
          </label>

          <label className={styles.field}>
            <span className={styles.lbl}>Button link</span>
            <input className={styles.input} type="text" value={ctaUrl} disabled={!ctaOn} placeholder="https://..." onChange={(e) => setCtaUrl(e.target.value)} />
          </label>

          <p className={styles.hint}>Navigation links come from the HeaderLinks list.</p>

          {state === 'error' && <p className={styles.error}>Could not save. You may not have permission to edit the HeaderConfig list.</p>}
        </div>

        <div className={styles.foot}>
          <button className={styles.cancel} type="button" onClick={onClose}>Cancel</button>
          <button className={styles.save} type="button" disabled={state === 'saving'} onClick={save}>
            {state === 'saving' ? 'Saving…' : 'Save'}
          </button>
        </div>
      </div>
    </div>
  );
};

export default HeaderSettings;
