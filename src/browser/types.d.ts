declare global {
  interface Window {
    ActiveXObject?: new (cls: string) => unknown;
  }
}

export {};
