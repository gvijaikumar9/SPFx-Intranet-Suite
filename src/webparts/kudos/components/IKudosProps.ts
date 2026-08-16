import * as React from 'react';
import { IKudos, IPeopleResult } from '../models/IKudos';

export interface IKudosProps {
  title: string;
  layout: string;      // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  showPhoto: boolean;  // true = profile photo (initials fallback); false = always initials
  webUrl: string;      // used to build the profile-photo url
  isDemo: boolean;
  accent: string;
  configured: boolean;
  canGive: boolean;             // give-kudos form is usable (demo, or a real list is set)
  items: IKudos[];
  loading: boolean;
  error?: string;
  submitting: boolean;
  submitError?: string;
  submittedKey: number;         // increments on each successful submit, clears the form
  onSearchPeople: (query: string) => Promise<IPeopleResult[]>;
  onSubmit: (recipientLogin: string | undefined, recipientName: string, message: string) => void;
}
