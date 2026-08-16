export interface IWeatherDay {
  date: string;
  min: number;
  max: number;
  code: number;
}

export interface IWeather {
  location: string;
  currentTemp: number;
  currentCode: number;
  unit: string;        // 'C' or 'F'
  days: IWeatherDay[];
}

/** WMO weather code -> emoji icon. */
export function weatherIcon(code: number): string {
  if (code === 0) { return '☀️'; }
  if (code === 1 || code === 2) { return '🌤️'; }
  if (code === 3) { return '☁️'; }
  if (code === 45 || code === 48) { return '🌫️'; }
  if (code >= 51 && code <= 67) { return '🌧️'; }
  if (code >= 71 && code <= 77) { return '🌨️'; }
  if (code >= 80 && code <= 82) { return '🌦️'; }
  if (code >= 85 && code <= 86) { return '🌨️'; }
  if (code >= 95) { return '⛈️'; }
  return '🌡️';
}

/** WMO weather code -> short condition text. */
export function weatherText(code: number): string {
  if (code === 0) { return 'Clear'; }
  if (code === 1) { return 'Mainly clear'; }
  if (code === 2) { return 'Partly cloudy'; }
  if (code === 3) { return 'Overcast'; }
  if (code === 45 || code === 48) { return 'Fog'; }
  if (code >= 51 && code <= 57) { return 'Drizzle'; }
  if (code >= 61 && code <= 67) { return 'Rain'; }
  if (code >= 71 && code <= 77) { return 'Snow'; }
  if (code >= 80 && code <= 82) { return 'Showers'; }
  if (code >= 85 && code <= 86) { return 'Snow showers'; }
  if (code >= 95) { return 'Thunderstorm'; }
  return 'Weather';
}
