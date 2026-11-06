export const TASWIYAH_LOGO_PATH = "/taswiyah-logo.png"

export const REPORT_LOGO_PRINT_STYLES = `
  .report-print-logo { display: none; }
  @media print {
    .report-print-logo {
      display: block !important;
      margin: 0 auto 10px;
      text-align: center;
      break-inside: avoid;
      page-break-inside: avoid;
    }
    .report-print-logo img {
      display: inline-block;
      width: 100px;
      height: 68px;
      object-fit: contain;
    }
  }
`

export function reportLogoMarkup() {
  const logoUrl = new URL(TASWIYAH_LOGO_PATH, window.location.origin).href
  return `<div class="report-print-logo"><img src="${logoUrl}" alt="شعار تسوية هب"></div>`
}
