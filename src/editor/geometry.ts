export const ROW_HEADER = 48,
  COLUMN_HEADER = 28,
  DEFAULT_ROW = 28,
  DEFAULT_COLUMN = 112;
export class Axis {
  readonly offsets: Float64Array;
  constructor(
    readonly count: number,
    sizes: Record<number, number>,
    hidden: number[],
    defaultSize: number,
  ) {
    const set = new Set(hidden);
    this.offsets = new Float64Array(count + 1);
    for (let i = 0; i < count; i++)
      this.offsets[i + 1] =
        this.offsets[i] + (set.has(i) ? 0 : (sizes[i] ?? defaultSize));
  }
  get total(): number {
    return this.offsets[this.count];
  }
  size(index: number): number {
    return this.offsets[index + 1] - this.offsets[index];
  }
  at(offset: number): number {
    let lo = 0,
      hi = this.count;
    while (lo < hi) {
      const mid = Math.ceil((lo + hi) / 2);
      if (this.offsets[mid] <= offset) lo = mid;
      else hi = mid - 1;
    }
    return Math.min(this.count - 1, lo);
  }
  visible(
    scroll: number,
    available: number,
    frozen: number,
  ): { index: number; start: number; size: number; frozen: boolean }[] {
    const cells: {
        index: number;
        start: number;
        size: number;
        frozen: boolean;
      }[] = [],
      boundary = this.offsets[frozen];
    for (let i = 0; i < frozen && this.offsets[i] < available; i++)
      if (this.size(i))
        cells.push({
          index: i,
          start: this.offsets[i],
          size: this.size(i),
          frozen: true,
        });
    for (
      let i = Math.max(frozen, this.at(scroll + boundary));
      i < this.count && this.offsets[i] - scroll < available;
      i++
    )
      if (this.size(i))
        cells.push({
          index: i,
          start: this.offsets[i] - scroll,
          size: this.size(i),
          frozen: false,
        });
    return cells;
  }
}

export function validateZoom(zoom: number): number {
  if (!Number.isFinite(zoom) || zoom < 0.5 || zoom > 2)
    throw new Error("Zoom must be between 0.5 and 2");
  return zoom;
}
