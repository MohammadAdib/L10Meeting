export function resetAll(): void {
  if (!confirm('Reset all fields? This cannot be undone.')) return;
  location.reload();
}
