import * as React from 'react';
import { IPerson } from '../models/IPerson';

export interface IPeopleDirectoryProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  showPhoto: boolean;
  webUrl: string;
  isDemo: boolean;
  accent: string;
  demoPeople: IPerson[];
  onSearch: (query: string) => Promise<IPerson[]>;
}
