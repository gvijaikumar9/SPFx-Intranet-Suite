import * as React from 'react';
import { IChartDatum } from '../models/IChartDatum';

export interface IChartFromListProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  isDemo: boolean;
  accent: string;
  items: IChartDatum[];
  barColors: (string | undefined)[];   // resolved fill per bar; undefined = accent
  loading: boolean;
  error?: string;
}
