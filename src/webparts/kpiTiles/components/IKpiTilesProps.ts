import * as React from 'react';
import { IKpi } from '../models/IKpi';

export interface IKpiTilesProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  columns: number;
  isDemo: boolean;
  accent: string;
  items: IKpi[];
  loading: boolean;
  error?: string;
}
