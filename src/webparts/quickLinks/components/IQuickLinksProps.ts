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
  tileColors: (string | undefined)[];   // resolved background per tile index; undefined = default
  loading: boolean;
  error?: string;
}
