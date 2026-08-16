import * as React from 'react';
import { ICelebration } from '../models/ICelebration';

export interface ICelebrationsProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  showPhoto: boolean;
  webUrl: string;
  isDemo: boolean;
  accent: string;
  items: ICelebration[];
  loading: boolean;
  error?: string;
}
