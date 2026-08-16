import * as React from 'react';
import { IEvent } from '../models/IEvent';

export interface IEventsListProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  isDemo: boolean;
  accent: string;
  items: IEvent[];
  loading: boolean;
  error?: string;
}
