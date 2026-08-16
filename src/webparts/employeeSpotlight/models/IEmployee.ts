export interface IEmployee {
  id: number;
  name: string;      // Employee display name, or list Title as fallback
  month: string;     // EomMonth
  citation: string;  // Citation note
  email?: string;    // used to build the profile photo url
}
