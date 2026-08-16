import * as React from 'react';
import { IHoliday } from '../models/IHoliday';

export interface IHolidaysListProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  isDemo: boolean;
  accent: string;
  items: IHoliday[];
  loading: boolean;
  error?: string;
}
