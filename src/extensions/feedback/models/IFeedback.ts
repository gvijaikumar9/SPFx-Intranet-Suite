export interface IFeedback {
  rating: number;     // 1-5, or 0 if not set
  comment: string;
  pageUrl: string;    // where the feedback was given
  contact: string;    // user display name/email if they opted in, else ''
}
