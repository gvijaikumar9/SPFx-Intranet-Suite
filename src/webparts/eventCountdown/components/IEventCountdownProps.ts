import * as React from 'react';

export interface IEventCountdownProps {
  title: string;
  layout: string;              // card | minimal | bold | compact
  showTitle: boolean;
  frameStyle: React.CSSProperties;
  eventName: string;
  targetIso: string;           // ISO date string of the event
  isDemo: boolean;             // true when no target date is set yet
  accent: string;
  unitColors: (string | undefined)[];   // resolved background per unit box (d,h,m,s); undefined = default
}
