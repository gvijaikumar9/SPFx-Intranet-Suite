import * as React from 'react';

export interface IChip {
  count?: string;   // optional badge number/text
  label: string;
}

export interface IGreetingProps {
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;          // unused visually (greeting is the heading) - kept for consistency
  frameStyle: React.CSSProperties;
  name: string;                // first name of the current user
  chips: IChip[];
  accent: string;
}
