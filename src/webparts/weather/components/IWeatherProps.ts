import * as React from 'react';
import { IWeather } from '../models/IWeather';

export interface IWeatherProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  isDemo: boolean;
  accent: string;
  weather?: IWeather;
  loading: boolean;
  error?: string;
}
