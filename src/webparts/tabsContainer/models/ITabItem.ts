export interface ITabItem {
  id: number;
  tab: string;       // the group this item belongs to (a tab label)
  title: string;
  subtitle?: string;
  href?: string;
}

/** Distinct tab labels in first-seen order. */
export function tabsOf(items: ITabItem[]): string[] {
  // Null-prototype map so a tab literally named "toString"/"constructor" cannot
  // collide with an inherited member.
  const seen: { [k: string]: boolean } = Object.create(null);
  const order: string[] = [];
  for (let i = 0; i < items.length; i++) {
    const t = items[i].tab || 'General';
    if (!seen[t]) { seen[t] = true; order.push(t); }
  }
  return order;
}
