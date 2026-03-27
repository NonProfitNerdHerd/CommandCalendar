/**
 * Rasterize a DOM subtree to A4/Letter-friendly PDF (preserves layout vs. text-only export).
 */
export async function downloadElementAsPdf(element: HTMLElement, filename: string): Promise<void> {
  if (!element) {
    return;
  }

  const html2canvas = (await import('html2canvas')).default;
  const { jsPDF } = await import('jspdf');

  const canvas = await html2canvas(element, {
    scale: Math.min(2, (typeof window !== 'undefined' && window.devicePixelRatio) || 2),
    useCORS: true,
    logging: false,
    backgroundColor: '#ffffff',
    scrollX: 0,
    scrollY: typeof window !== 'undefined' ? -window.scrollY : 0
  });

  const imgData: string = canvas.toDataURL('image/png', 1.0);
  const orientation: 'p' | 'l' = canvas.width >= canvas.height ? 'l' : 'p';
  const pdf = new jsPDF({
    orientation,
    unit: 'mm',
    format: 'a4'
  });

  const pageW: number = pdf.internal.pageSize.getWidth();
  const pageH: number = pdf.internal.pageSize.getHeight();
  const margin: number = 8;
  const footerFontSize: number = 8;
  pdf.setFontSize(footerFontSize);
  const generatedAt: string = new Date().toLocaleString(undefined, {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit'
  });
  const footerText: string =
    `This document is a point-in-time snapshot generated on ${generatedAt} and may not reflect any changes made after that moment.`;
  const textMaxW: number = pageW - margin * 2;
  const footerLines: string[] = pdf.splitTextToSize(footerText, textMaxW);
  const lineHeightMm: number = 3.4;
  const footerGapMm: number = 4;
  const footerBlockH: number = footerLines.length * lineHeightMm + footerGapMm;

  const maxW: number = pageW - margin * 2;
  const maxH: number = pageH - margin * 2 - footerBlockH;
  const imgRatio: number = canvas.width / canvas.height;
  let drawW: number = maxW;
  let drawH: number = drawW / imgRatio;
  if (drawH > maxH) {
    drawH = maxH;
    drawW = drawH * imgRatio;
  }
  const offsetX: number = margin + (maxW - drawW) / 2;
  const offsetY: number = margin + (maxH - drawH) / 2;

  pdf.addImage(imgData, 'PNG', offsetX, offsetY, drawW, drawH, undefined, 'FAST');

  const lastBaselineY: number = pageH - margin - 2;
  const firstBaselineY: number = lastBaselineY - (footerLines.length - 1) * lineHeightMm;
  pdf.text(footerLines, pageW / 2, firstBaselineY, { align: 'center' });

  pdf.save(filename);
}
