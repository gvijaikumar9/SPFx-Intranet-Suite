export interface ICelebration {
  id: number;
  name: string;    // list Title (person name)
  date: string;    // ISO date (CelebrationDate)
  type: string;    // CelebrationType, e.g. Birthday / Work anniversary
  email?: string;  // Person field email, for the profile photo
}
