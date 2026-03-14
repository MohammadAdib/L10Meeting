import * as fs from './fs-service';
import { confirmDialog } from './utils';
import { getLogoUrl, handleLogoClick } from './logo';

let _menuEl: HTMLElement | null = null;

export function showSettingsMenu(anchor: HTMLElement): void {
  // Close if already open
  if (_menuEl) { _menuEl.remove(); _menuEl = null; return; }

  const folderName = fs.getFolderName();
  const hasLogo = !!getLogoUrl();
  const menu = document.createElement('div');
  menu.className = 'settings-menu';
  menu.innerHTML = `
    <div class="settings-menu-item settings-folder-info">
      <span class="settings-folder-label">Data folder</span>
      <span class="settings-folder-name">${folderName || 'None'}</span>
    </div>
    <button class="settings-menu-item" id="settingsChangeFolder">Change folder</button>
    <button class="settings-menu-item settings-danger" id="settingsForgetFolder">Forget folder</button>
    <div class="settings-divider"></div>
    <button class="settings-menu-item" id="settingsChangeLogo">${hasLogo ? 'Change logo' : 'Add logo'}</button>
    ${hasLogo ? `<button class="settings-menu-item settings-danger" id="settingsRemoveLogo">Remove logo</button>` : ''}
  `;

  const rect = anchor.getBoundingClientRect();
  menu.style.top = `${rect.bottom + 6}px`;
  menu.style.right = `${window.innerWidth - rect.right}px`;
  document.body.appendChild(menu);
  _menuEl = menu;

  requestAnimationFrame(() => menu.classList.add('visible'));

  document.getElementById('settingsChangeFolder')?.addEventListener('click', async () => {
    close();
    const ok = await fs.pickFolder();
    if (ok) location.reload();
  });

  document.getElementById('settingsChangeLogo')?.addEventListener('click', () => {
    close();
    handleLogoClick(() => location.reload());
  });

  document.getElementById('settingsRemoveLogo')?.addEventListener('click', async () => {
    close();
    await fs.deleteLogo();
    location.reload();
  });

  document.getElementById('settingsForgetFolder')?.addEventListener('click', async () => {
    close();
    if (!await confirmDialog('Forget the saved folder? You will need to select a folder again.', 'Forget', true)) return;
    await fs.forgetFolder();
    location.reload();
  });

  // Close on outside click
  const onClickOutside = (e: MouseEvent) => {
    if (!menu.contains(e.target as Node) && e.target !== anchor) {
      close();
    }
  };
  setTimeout(() => document.addEventListener('click', onClickOutside), 0);

  function close() {
    document.removeEventListener('click', onClickOutside);
    menu.remove();
    _menuEl = null;
  }
}
