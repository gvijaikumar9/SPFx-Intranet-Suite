import * as React from 'react';
import { ITabItem } from '../models/ITabItem';

export interface ITabsContainerProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  isDemo: boolean;
  accent: string;
  items: ITabItem[];
  loading: boolean;
  error?: string;
}
