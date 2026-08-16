import * as React from 'react';
import { IFaq } from '../models/IFaq';

export interface IFaqAccordionProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  isDemo: boolean;
  accent: string;
  items: IFaq[];
  loading: boolean;
  error?: string;
}
