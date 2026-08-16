import * as React from 'react';
import { IEmployee } from '../models/IEmployee';

export interface IEmployeeSpotlightProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  isDemo: boolean;
  accent: string;
  webUrl: string;      // used to build the profile-photo url
  showPhoto: boolean;  // true = profile photo (initials fallback); false = always initials
  items: IEmployee[];
  loading: boolean;
  error?: string;
  emptyNote?: string;   // muted guidance shown when there is nothing to feature
}
