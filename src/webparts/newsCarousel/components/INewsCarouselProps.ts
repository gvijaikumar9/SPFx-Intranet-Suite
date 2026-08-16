import * as React from 'react';
import { INews } from '../models/INews';

export interface INewsCarouselProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  rotateSeconds: number;
  isDemo: boolean;
  accent: string;
  items: INews[];
  loading: boolean;
  error?: string;
}
