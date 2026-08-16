import * as React from 'react';
import { useMemo } from 'react';
import styles from './Weather.module.scss';
import { IWeatherProps } from './IWeatherProps';
import { weatherIcon, weatherText } from '../models/IWeather';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

function dayName(iso: string): string {
  const d = new Date(iso);
  return isNaN(d.getTime()) ? '' : d.toLocaleDateString(undefined, { weekday: 'short' });
}

const Weather: React.FC<IWeatherProps> = (props) => {
  const { title, layout, showTitle, frameStyle, isDemo, accent, weather, loading, error } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  return (
    <section className={`${styles.weather} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Weather'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Set a location in the property pane.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && !weather && (
        <p className={styles.muted}>Set a location to show the weather.</p>
      )}

      {!loading && !error && weather && (
        <>
          <div className={styles.current}>
            <span className={styles.icon} aria-hidden="true">{weatherIcon(weather.currentCode)}</span>
            <div className={styles.now}>
              <p className={styles.temp}>{weather.currentTemp}&deg;{weather.unit}</p>
              <p className={styles.cond}>{weatherText(weather.currentCode)}</p>
            </div>
            <p className={styles.place}>{weather.location}</p>
          </div>

          {weather.days.length > 0 && (
            <ul className={styles.forecast}>
              {weather.days.map((d) => (
                <li key={d.date} className={styles.fday}>
                  <span className={styles.fname}>{dayName(d.date)}</span>
                  <span className={styles.ficon} aria-hidden="true">{weatherIcon(d.code)}</span>
                  <span className={styles.frange}>
                    <b>{d.max}&deg;</b> <span className={styles.fmin}>{d.min}&deg;</span>
                  </span>
                </li>
              ))}
            </ul>
          )}
        </>
      )}
    </section>
  );
};

export default Weather;
