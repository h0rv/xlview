export interface BenchProvider {
  id: string;
  label: string;
  parse(bytes: Uint8Array): Promise<unknown>;
  load(bytes: Uint8Array): Promise<unknown>;
  render(): Promise<unknown>;
  scroll?(opts: {
    dx: number;
    dy: number;
    step?: number;
    totalSteps?: number;
  }): Promise<unknown>;
  resetScroll?(): Promise<void>;
  destroy?(): Promise<void>;
}

export interface BenchProviderEntry {
  id: string;
  label: string;
  create(opts: ProviderCreateOpts): Promise<BenchProvider>;
}

export interface ProviderCreateOpts {
  canvas: HTMLCanvasElement;
  container: HTMLElement;
  domRoot: HTMLElement;
  dpr: number;
  width: number;
  height: number;
}

export interface ManifestEntry {
  id: string;
  name: string;
  label: string;
  file: string;
  size: string;
}
