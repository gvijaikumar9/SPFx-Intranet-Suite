export interface IImage {
  id: number;
  name: string;
  url: string;   // server-relative file url; empty for demo placeholder tiles
}

const IMAGE_EXT = ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg'];

export function isImageName(name: string): boolean {
  const dot = name.lastIndexOf('.');
  if (dot < 0) { return false; }
  const ext = name.substring(dot).toLowerCase();
  return IMAGE_EXT.indexOf(ext) >= 0;
}
