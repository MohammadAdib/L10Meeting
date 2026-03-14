import { loadLogo, saveLogo } from './fs-service';

let _logoUrl: string | null = null;

/** Load logo from data folder into memory */
export async function initLogo(): Promise<void> {
  _logoUrl = await loadLogo();
}

/** Get the current logo URL (or null if none) */
export function getLogoUrl(): string | null {
  return _logoUrl;
}

/** Open a file picker for the user to choose a logo, then save it and call onDone */
export function handleLogoClick(onDone: () => void): void {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    if (!file) return;
    _logoUrl = await saveLogo(file);
    onDone();
  });
  input.click();
}
