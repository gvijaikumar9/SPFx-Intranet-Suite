import * as React from 'react';
import { IAnnouncement } from '../models/IAnnouncement';

export interface IAnnouncementsTickerProps {
  label: string;
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  items: IAnnouncement[];
  rotateSeconds: number;
  showControls: boolean;
  loading: boolean;
  error?: string;
  isDemo: boolean;
  accent: string;
  isDarkTheme: boolean;
  layout: string;      // card | minimal | bold | compact
}
