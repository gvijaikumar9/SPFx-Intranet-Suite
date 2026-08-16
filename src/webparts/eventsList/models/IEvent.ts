export interface IEvent {
  id: number;
  title: string;
  start: string;      // ISO date string (EventDate)
  end?: string;       // ISO date string (EndDate)
  location?: string;  // Location
}
