import * as React from 'react';
import { IImage } from '../models/IImage';

export interface IImageGalleryProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  columns: number;
  isDemo: boolean;
  accent: string;
  items: IImage[];
  loading: boolean;
  error?: string;
}
