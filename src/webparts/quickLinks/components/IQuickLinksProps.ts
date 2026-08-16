import * as React from 'react';
import { ILink } from '../models/ILink';

export interface IQuickLinksProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  columns: number;
  isDemo: boolean;
  accent: string;
  items: ILink[];
  loading: boolean;
  error?: string;
}
