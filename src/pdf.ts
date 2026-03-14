import { showToast } from './utils';

export async function exportPDF(): Promise<void> {
  // Expand all sections for PDF
  for (let i = 1; i <= 7; i++) {
    document.getElementById(`body-${i}`)?.classList.remove('collapsed');
    document.getElementById(`chev-${i}`)?.classList.add('open');
  }

  const el = document.querySelector<HTMLElement>('.container');
  if (!el) return;

  const dateInput = document.getElementById('metaDate') as HTMLInputElement;
  const filename = `L10_Meeting_${dateInput?.value || 'draft'}.pdf`;

  showToast('Generating PDF...');

  // Dynamic import of html2pdf
  const html2pdf = (await import('html2pdf.js')).default;

  const opt = {
    margin: [10, 10, 10, 10] as [number, number, number, number],
    filename,
    image: { type: 'jpeg' as const, quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true, backgroundColor: '#f9f5f0' },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
    pagebreak: { mode: ['avoid-all', 'css', 'legacy'] },
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await html2pdf().set(opt as any).from(el).save();
  showToast('PDF exported!');
}
