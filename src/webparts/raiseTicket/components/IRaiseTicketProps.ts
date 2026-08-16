import * as React from 'react';
import { ITicket } from '../models/ITicket';

export interface IRaiseTicketProps {
  title: string;
  layout: string;      // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  categories: string[];
  isDemo: boolean;
  accent: string;
  configured: boolean;        // a real list is wired up (or demo mode is on)
  submitting: boolean;
  submitError?: string;
  submittedId?: number;       // id of the last created ticket; changes on each success
  recent: ITicket[];
  loadingRecent: boolean;
  recentError?: string;
  onSubmit: (subject: string, category: string, description: string) => void;
}
