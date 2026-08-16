import * as React from 'react';
import { IOrgNode } from '../models/IOrgNode';

export interface IOrgChartProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  showPhoto: boolean;
  webUrl: string;
  isDemo: boolean;
  accent: string;
  manager?: IOrgNode;
  me?: IOrgNode;
  reports: IOrgNode[];
  loading: boolean;
  error?: string;
}
