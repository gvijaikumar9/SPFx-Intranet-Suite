export interface IKudos {
  id: number;
  headline: string;   // list Title
  message: string;    // KudosMsg note
  toName?: string;    // ToPerson display name
  toEmail?: string;   // ToPerson email, for the profile photo
  fromName?: string;  // FromPerson display name
}

export interface IPeopleResult {
  key: string;          // login name, used with ensureuser
  text: string;         // display name
  secondaryText?: string; // email
}
