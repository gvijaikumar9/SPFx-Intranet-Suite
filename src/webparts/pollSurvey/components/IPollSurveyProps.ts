import * as React from 'react';

export interface IPollOption {
  option: string;
  count: number;
}

export interface IPollSurveyProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  question: string;
  options: string[];
  results: IPollOption[];
  total: number;
  voted: boolean;
  configured: boolean;
  isDemo: boolean;
  accent: string;
  submitting: boolean;
  submitError?: string;
  optionColors: (string | undefined)[];   // resolved bar fill per option; undefined = accent
  loading: boolean;
  error?: string;
  onVote: (option: string) => void;
}
