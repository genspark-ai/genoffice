import iconDocx from './assets/file-docx.svg'
import iconXlsx from './assets/file-xlsx.svg'
import iconPptx from './assets/file-pptx.svg'
import iconPdf from './assets/file-pdf.svg'
import iconMd from './assets/file-md.svg'
import iconHtml from './assets/file-html.svg'
import iconJson from './assets/file-json.svg'
import iconXml from './assets/file-xml.svg'
import iconYaml from './assets/file-yaml.svg'
import iconSql from './assets/file-sql.svg'
import iconCsv from './assets/file-csv.svg'
import iconTsv from './assets/file-tsv.svg'
import iconTxt from './assets/file-txt.svg'

const FILE_ICONS: Record<string, string> = {
  docx: iconDocx,
  doc: iconDocx,
  xlsx: iconXlsx,
  xls: iconXlsx,
  pptx: iconPptx,
  ppt: iconPptx,
  pdf: iconPdf,
  md: iconMd,
  markdown: iconMd,
  html: iconHtml,
  htm: iconHtml,
  json: iconJson,
  xml: iconXml,
  yaml: iconYaml,
  yml: iconYaml,
  sql: iconSql,
  csv: iconCsv,
  tsv: iconTsv,
  txt: iconTxt,
}

export function FileBadge({ ext, size }: { ext: string; size: number }) {
  const icon = FILE_ICONS[ext]
  if (icon) {
    return (
      <img
        src={icon}
        width={size}
        height={size}
        alt=""
        aria-hidden="true"
        style={{ objectFit: 'contain' }}
      />
    )
  }
  const label = ext ? ext[0].toUpperCase() : '?'
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <rect width="32" height="32" rx="7.5" fill="#98a2b3" />
      <text
        x="16"
        y="16.5"
        textAnchor="middle"
        dominantBaseline="central"
        fill="#fff"
        fontSize={17}
        fontWeight="700"
        fontFamily="system-ui, -apple-system, 'Segoe UI', sans-serif"
      >
        {label}
      </text>
    </svg>
  )
}
