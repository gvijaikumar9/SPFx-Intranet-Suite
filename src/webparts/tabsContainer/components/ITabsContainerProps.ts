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
  tabColors: (string | undefined)[];   // resolved background per tab header; undefined = default
  loading: boolean;
  error?: string;
}
