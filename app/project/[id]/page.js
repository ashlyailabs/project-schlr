'use client'
import { useState, useEffect, useCallback, useRef, Fragment } from 'react'
import { useRouter, useParams } from 'next/navigation'
import ExcelJS from 'exceljs'
import { supabase } from '@/lib/supabase'
import ThemeToggle from '@/components/ThemeToggle'

const STATUS = {
  'not-started': { label: 'Not Started', color: 'text-gray-500',  bg: 'bg-cream-200 dark:bg-gray-700',  bar: 'bg-gray-300' },
  'in-progress':  { label: 'In Progress', color: 'text-blue-700 dark:text-blue-400',  bg: 'bg-blue-50 dark:bg-blue-900/40',   bar: 'bg-blue-400' },
  'done':         { label: 'Done',        color: 'text-green-700 dark:text-green-400', bg: 'bg-green-50 dark:bg-green-900/40',  bar: 'bg-green-400' },
  'delayed':      { label: 'Delayed',     color: 'text-red-600 dark:text-red-400',   bg: 'bg-red-50 dark:bg-red-900/40',    bar: 'bg-red-400' },
}

const TODAY = new Date().toISOString().split('T')[0]

function fmtFull(s) {
  return new Date(s + 'T00:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
}

function isWeekend(ds) {
  return new Date(ds + 'T00:00:00').getDay() === 0
}

function dateRange(start, end) {
  const dates = []
  const cur = new Date(start + 'T00:00:00')
  const last = new Date(end + 'T00:00:00')
  while (cur <= last) {
    dates.push(cur.toISOString().split('T')[0])
    cur.setDate(cur.getDate() + 1)
  }
  return dates
}

function groupByMonth(dates) {
  const groups = []
  dates.forEach(ds => {
    const label = new Date(ds + 'T00:00:00')
      .toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })
      .toUpperCase()
    if (!groups.length || groups[groups.length - 1].label !== label)
      groups.push({ label, dates: [] })
    groups[groups.length - 1].dates.push(ds)
  })
  return groups
}

function hlKey(taskId, date) {
  return `${taskId}|${date}`
}

function isHighlighted(highlights, taskId, date) {
  return highlights.has(hlKey(taskId, date))
}

function getProjectDates(project) {
  if (!project?.start_date || !project?.end_date) return []
  return dateRange(project.start_date, project.end_date)
}

function cellBg(ds, rowIdx, isHl) {
  if (isHl) return 'bg-blue-500 gantt-highlight'
  if (ds === TODAY) return rowIdx % 2 === 0 ? 'bg-blue-50 dark:bg-blue-900/20' : 'bg-blue-50/80 dark:bg-blue-900/15'
  if (isWeekend(ds)) return rowIdx % 2 === 0 ? 'bg-red-50 dark:bg-red-900/10' : 'bg-red-50/80 dark:bg-red-900/5'
  return rowIdx % 2 === 0 ? 'bg-cream-50 dark:bg-gray-800' : 'bg-cream-100 dark:bg-gray-800/80'
}

function GanttChart({ tasks, dates, highlights, onToggleHighlight, onEditTask }) {
  if (!tasks.length || !dates.length) return null

  const months = groupByMonth(dates)
  const cellW = 32
  const leftW = 340

  return (
    <div id="gantt-print-area" className="overflow-x-auto border border-gray-200 dark:border-gray-600 rounded-xl bg-cream-50 dark:bg-gray-800">
      <table
        className="border-collapse table-fixed gantt-table"
        style={{ width: leftW + dates.length * cellW }}
      >
        <colgroup>
          <col style={{ width: leftW }} />
          {dates.map(d => <col key={d} style={{ width: cellW }} />)}
        </colgroup>
        <thead>
          <tr className="bg-cream-100 dark:bg-gray-700">
            <th className="text-left px-4 py-2 text-[11px] font-medium text-gray-500 dark:text-gray-300 border-b border-r-2 border-gray-200 dark:border-gray-600 whitespace-nowrap">
              WORK ACTIVITY DETAILS
            </th>
            {months.map(mg => (
              <th
                key={mg.label}
                colSpan={mg.dates.length}
                className="text-center py-2 text-[11px] font-medium text-gray-700 dark:text-gray-200 tracking-wide border-b border-r border-gray-200 dark:border-gray-600 whitespace-nowrap"
              >
                {mg.label}
              </th>
            ))}
          </tr>
          <tr className="bg-cream-100 dark:bg-gray-700">
            <th className="text-left px-4 py-1 text-[10px] font-medium text-gray-400 dark:text-gray-400 border-b border-r-2 border-gray-200 dark:border-gray-600">
              SL&nbsp;&nbsp;Activity
            </th>
            {dates.map(ds => {
              const day = new Date(ds + 'T00:00:00').getDate()
              const wknd = isWeekend(ds)
              const isToday = ds === TODAY
              return (
                <th
                  key={ds}
                  className={`text-center py-1 text-[10px] border-b border-r border-gray-200 dark:border-gray-600 h-7 ${
                    isToday
                      ? 'text-blue-600 dark:text-blue-400 font-bold bg-blue-50 dark:bg-blue-900/30'
                      : wknd
                        ? 'text-red-500 dark:text-red-400 font-semibold bg-red-50 dark:bg-red-900/20'
                        : 'text-gray-500 dark:text-gray-400 font-normal'
                  }`}
                >
                  {day}
                </th>
              )
            })}
          </tr>
        </thead>
        <tbody>
          {tasks.map((task, i) => (
            <Fragment key={task.id}>
              <tr className={i % 2 === 0 ? 'bg-cream-50 dark:bg-gray-800' : 'bg-cream-100 dark:bg-gray-800/80'}>
                <td
                  onClick={() => onEditTask?.(task)}
                  className="px-4 h-11 border-b border-r-2 border-gray-200 dark:border-gray-600 text-xs font-semibold text-gray-900 dark:text-white overflow-hidden text-ellipsis whitespace-nowrap align-middle cursor-pointer hover:bg-cream-200 dark:hover:bg-gray-700"
                >
                  <span className="text-gray-400 dark:text-gray-500 font-normal mr-2">{task.sl}</span>
                  <span className="underline">{task.name}</span>
                </td>
                {dates.map(ds => {
                  const hl = isHighlighted(highlights, task.id, ds)
                  return (
                    <td
                      key={ds}
                      onClick={() => onToggleHighlight(task.id, ds)}
                      title={`${task.name} — ${ds}`}
                      className={`h-11 border-b border-r border-gray-200 dark:border-gray-600 cursor-pointer p-0 ${cellBg(ds, i, hl)}`}
                    />
                  )
                })}
              </tr>
              <tr className={i % 2 === 0 ? 'bg-cream-50 dark:bg-gray-800' : 'bg-cream-100 dark:bg-gray-800/80'}>
                <td
                  onClick={() => onEditTask?.(task)}
                  className="px-4 py-1 pb-2.5 border-b border-r-2 border-gray-200 dark:border-gray-600 text-[11px] text-gray-500 dark:text-gray-400 leading-snug align-top cursor-pointer hover:bg-cream-200 dark:hover:bg-gray-700"
                >
                  {task.description || '\u00A0'}
                </td>
                {dates.map(ds => {
                  const hl = isHighlighted(highlights, task.id, ds)
                  return (
                    <td
                      key={ds}
                      onClick={() => onToggleHighlight(task.id, ds)}
                      className={`h-6 border-b border-r border-gray-200 dark:border-gray-600 cursor-pointer p-0 ${cellBg(ds, i, hl)}`}
                    />
                  )
                })}
              </tr>
            </Fragment>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function monthPrintLabel(firstDate) {
  return new Date(firstDate + 'T00:00:00')
    .toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })
    .toUpperCase()
    .replace(' ', ' - ')
}

function buildPrintAreaHTML(project, tasks, dates, highlights, logoDataUrl = '') {
  const months = groupByMonth(dates)
  const hasDates = project.start_date && project.end_date
  const dateRangeStr = hasDates
    ? `${fmtFull(project.start_date)} – ${fmtFull(project.end_date)}`
    : ''

  let table = '<table><thead>'

  table += '<tr><th colspan="2" class="month-header"></th>'
  months.forEach(mg => {
    table += `<th colspan="${mg.dates.length}" class="month-header">${monthPrintLabel(mg.dates[0])}</th>`
  })
  table += '</tr>'

  table += '<tr><th class="gantt-sl-col">SL NO</th><th class="gantt-activity-col">WORK ACTIVITY DETAILS</th>'
  dates.forEach(ds => {
    const day = new Date(ds + 'T00:00:00').getDate()
    const sundayCls = isWeekend(ds) ? ' sunday-col sunday-day' : ''
    table += `<th class="gantt-date-col${sundayCls}">${day}</th>`
  })
  table += '</tr></thead><tbody>'

  tasks.forEach((task, i) => {
    const rowBg = i % 2 === 0 ? '#FFFFFF' : '#F5F5F5'
    table += `<tr style="background:${rowBg}">`
    table += `<td class="gantt-sl-col" style="text-align:center;vertical-align:top">${task.sl}</td>`
    table += `<td class="gantt-activity-col" style="vertical-align:top">`
    table += `<div style="font-weight:bold;text-decoration:underline">${escapeHtml(task.name)}</div>`
    if (task.description) {
      table += `<div style="font-size:6pt;color:#555;margin-top:2px;line-height:1.3">${escapeHtml(task.description)}</div>`
    }
    table += '</td>'
    dates.forEach(ds => {
      const hl = isHighlighted(highlights, task.id, ds)
      let cls = 'gantt-date-col'
      if (hl) cls += ' highlighted-cell'
      else if (isWeekend(ds)) cls += ' sunday-col'
      table += `<td class="${cls}">&nbsp;</td>`
    })
    table += '</tr>'
  })

  table += '</tbody></table>'

  return `
    <div class="print-doc-header">
      ${logoDataUrl ? `<img src="${logoDataUrl}" alt="Ashly" class="print-logo">` : ''}
      <div class="print-doc-meta">
        <div class="print-company">ASHLY INTERIORS &amp; CONSTRUCTIONS</div>
        <div class="print-project-name">${escapeHtml(project.name)}</div>
        ${project.location ? `<div class="print-meta-line">${escapeHtml(project.location)}</div>` : ''}
        ${dateRangeStr ? `<div class="print-meta-line">${dateRangeStr}</div>` : ''}
      </div>
    </div>
    <hr style="margin: 8px 0; border: 0.5px solid #ccc;">
    ${table}
  `
}

const PRINT_AREA_STYLES = `
  #print-area {
    display: none;
    background: white !important;
    color: black !important;
  }

  #print-area * {
    background-color: inherit;
    color: inherit;
  }

  #print-area table {
    background: white !important;
    color: black !important;
  }

  #print-area th,
  #print-area td {
    background-color: white !important;
    color: black !important;
    border-color: #999 !important;
  }

  #print-area .month-header {
    background-color: #D9D9D9 !important;
    color: black !important;
  }

  #print-area .highlighted-cell {
    background-color: #4472C4 !important;
    color: white !important;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }

  #print-area .sunday-col {
    background-color: #FFE0E0 !important;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }

  #print-area .sunday-day {
    color: #CC0000 !important;
  }

  #print-area .print-company,
  #print-area .print-project-name,
  #print-area .print-meta-line {
    color: black !important;
  }

  #print-area .print-doc-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
  }
  #print-area .print-logo {
    height: 120px;
    width: auto;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  #print-area .print-doc-meta {
    text-align: right;
  }
  #print-area .print-company {
    font-weight: bold;
    font-size: 14pt;
  }
  #print-area .print-project-name {
    font-weight: bold;
    font-size: 11pt;
    margin-top: 4px;
  }
  #print-area .print-meta-line {
    font-size: 9pt;
    margin-top: 2px;
  }
  @media print {
    @page { size: A4 landscape; margin: 10mm; }
    html, body {
      background: white !important;
      color: black !important;
    }
    body * { visibility: hidden; }
    #print-area, #print-area * { visibility: visible; }
    #print-area {
      display: block;
      position: absolute;
      left: 0;
      top: 0;
      width: 100%;
      background: white !important;
      color: black !important;
    }
    #print-area .print-logo {
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    .gantt-date-col { width: 18px; min-width: 18px; max-width: 18px; }
    .gantt-activity-col { width: 280px; }
    .gantt-sl-col { width: 35px; }
    table { border-collapse: collapse; width: 100%; font-size: 7pt; }
    th, td { border: 0.5px solid #999; padding: 1px 2px; }
    th { font-weight: bold; text-align: center; }
  }
`

const LOGO_PATH = '/ashly_logo_.png'

async function fetchLogoDataUrl() {
  try {
    const url = `${window.location.origin}${LOGO_PATH}`
    const response = await fetch(url)
    if (!response.ok) return ''
    const blob = await response.blob()
    if (!blob.type.startsWith('image/')) return ''
    return await new Promise((resolve) => {
      const reader = new FileReader()
      reader.onloadend = () => resolve(typeof reader.result === 'string' ? reader.result : '')
      reader.onerror = () => resolve('')
      reader.readAsDataURL(blob)
    })
  } catch {
    return ''
  }
}

function waitForPrintImages(container) {
  const images = container.querySelectorAll('img')
  if (!images.length) return Promise.resolve()
  return Promise.all([...images].map(img => new Promise(resolve => {
    if (img.complete) resolve()
    else {
      img.onload = resolve
      img.onerror = resolve
    }
  })))
}

async function downloadProjectPDF(project, tasks, dates, highlights) {
  if (!dates.length) return

  const logoDataUrl = await fetchLogoDataUrl()

  const styleEl = document.createElement('style')
  styleEl.id = 'print-area-styles'
  styleEl.textContent = PRINT_AREA_STYLES

  const printArea = document.createElement('div')
  printArea.id = 'print-area'
  printArea.innerHTML = buildPrintAreaHTML(project, tasks, dates, highlights, logoDataUrl)

  document.head.appendChild(styleEl)
  document.body.appendChild(printArea)

  await waitForPrintImages(printArea)
  await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))

  const cleanup = () => {
    printArea.remove()
    styleEl.remove()
    window.removeEventListener('afterprint', cleanup)
  }

  window.addEventListener('afterprint', cleanup)
  window.print()
}

const DEPARTMENT_BADGE = {
  corporate: 'bg-blue-50 dark:bg-blue-900/40 text-blue-700 dark:text-blue-400',
  government: 'bg-green-50 dark:bg-green-900/40 text-green-700 dark:text-green-400',
}

function deptLabel(dept) {
  const d = dept || 'corporate'
  return d.charAt(0).toUpperCase() + d.slice(1)
}

function taskPayload(form) {
  return {
    name: form.name,
    description: form.description || null,
    status: form.status,
    progress: form.progress,
  }
}

function TaskModal({ task, isNew, onSave, onDelete, onClose }) {
  const [f, setF] = useState({
    name:        task?.name        || '',
    description: task?.description || '',
    status:      task?.status      || 'not-started',
    progress:    task?.progress    ?? 0,
  })
  const inp = 'w-full border border-gray-200 dark:border-gray-600 rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-cream-50 dark:bg-gray-700 dark:text-white'

  return (
    <div className="no-print fixed inset-0 bg-black/30 backdrop-blur-sm flex items-center justify-center z-50 p-4">
      <div className="bg-cream-50 dark:bg-gray-800 rounded-2xl border border-gray-200 dark:border-gray-600 p-6 w-full max-w-sm max-h-[90vh] overflow-y-auto shadow-xl">
        <div className="flex items-center justify-between mb-5">
          <h3 className="text-base font-semibold text-gray-900 dark:text-white">{isNew ? 'New Task' : 'Edit Task'}</h3>
          <button onClick={onClose} className="text-gray-300 hover:text-gray-500 dark:hover:text-gray-400 text-lg leading-none">✕</button>
        </div>

        <div className="space-y-4">
          <div>
            <label className="text-xs text-gray-400 block mb-1.5">Activity Name</label>
            <input value={f.name} onChange={e => setF(p => ({ ...p, name: e.target.value }))} placeholder="e.g. Painting Works" className={inp} />
          </div>
          <div>
            <label className="text-xs text-gray-400 block mb-1.5">Description</label>
            <textarea value={f.description} onChange={e => setF(p => ({ ...p, description: e.target.value }))} rows={2}
              className={`${inp} resize-none`} placeholder="Optional details…" />
          </div>
          <div>
            <label className="text-xs text-gray-400 block mb-1.5">Status</label>
            <select value={f.status} onChange={e => setF(p => ({ ...p, status: e.target.value }))} className={inp}>
              {Object.entries(STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs text-gray-400 block mb-1.5">Progress — {f.progress}%</label>
            <input type="range" min={0} max={100} step={5} value={f.progress}
              onChange={e => setF(p => ({ ...p, progress: +e.target.value }))} className="w-full accent-blue-600" />
          </div>
        </div>

        <div className="flex items-center justify-between mt-6 gap-3">
          {!isNew && (
            <button onClick={() => { onDelete(); onClose() }}
              className="text-sm text-red-500 border border-red-100 dark:border-red-900/50 rounded-xl px-3 py-2 hover:bg-red-50 dark:hover:bg-red-900/20">
              Delete
            </button>
          )}
          <div className="flex gap-2 ml-auto">
            <button onClick={onClose}
              className="text-sm text-gray-500 dark:text-gray-400 border border-gray-200 dark:border-gray-600 rounded-xl px-3 py-2 hover:bg-cream-100 dark:hover:bg-gray-700">
              Cancel
            </button>
            <button onClick={() => { onSave(f); onClose() }}
              className="text-sm bg-blue-600 text-white rounded-xl px-4 py-2 hover:bg-blue-700 font-medium">
              {isNew ? 'Add Task' : 'Save'}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

function ProjectModal({ project, onSave, onClose }) {
  const [f, setF] = useState({
    name:       project.name        || '',
    location:   project.location    || '',
    start_date: project.start_date  || '',
    end_date:   project.end_date    || '',
  })
  const inp = 'w-full border border-gray-200 dark:border-gray-600 rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-cream-50 dark:bg-gray-700 dark:text-white'

  return (
    <div className="no-print fixed inset-0 bg-black/30 backdrop-blur-sm flex items-center justify-center z-50 p-4">
      <div className="bg-cream-50 dark:bg-gray-800 rounded-2xl border border-gray-200 dark:border-gray-600 p-6 w-full max-w-sm max-h-[90vh] overflow-y-auto shadow-xl">
        <div className="flex items-center justify-between mb-5">
          <h3 className="text-base font-semibold text-gray-900 dark:text-white">Edit Project</h3>
          <button onClick={onClose} className="text-gray-300 hover:text-gray-500 dark:hover:text-gray-400 text-lg leading-none">✕</button>
        </div>

        <div className="space-y-4">
          <div>
            <label className="text-xs text-gray-400 block mb-1.5">Project Name</label>
            <input value={f.name} onChange={e => setF(p => ({ ...p, name: e.target.value }))} placeholder="Project name" className={inp} />
          </div>
          <div>
            <label className="text-xs text-gray-400 block mb-1.5">Location</label>
            <input value={f.location} onChange={e => setF(p => ({ ...p, location: e.target.value }))} placeholder="e.g. Dubai Marina, Office 4B" className={inp} />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-gray-400 block mb-1.5">Start Date</label>
              <input type="date" value={f.start_date} onChange={e => setF(p => ({ ...p, start_date: e.target.value }))} className={inp} />
            </div>
            <div>
              <label className="text-xs text-gray-400 block mb-1.5">End Date</label>
              <input type="date" value={f.end_date} onChange={e => setF(p => ({ ...p, end_date: e.target.value }))} className={inp} />
            </div>
          </div>
        </div>

        <div className="flex gap-2 justify-end mt-6">
          <button onClick={onClose}
            className="text-sm text-gray-500 dark:text-gray-400 border border-gray-200 dark:border-gray-600 rounded-xl px-3 py-2 hover:bg-cream-100 dark:hover:bg-gray-700">
            Cancel
          </button>
          <button onClick={() => { onSave(f); onClose() }}
            className="text-sm bg-blue-600 text-white rounded-xl px-4 py-2 hover:bg-blue-700 font-medium">
            Save
          </button>
        </div>
      </div>
    </div>
  )
}

function DownloadMenu({ project, tasks, dates, highlights, onClose }) {
  const ref = useRef(null)

  useEffect(() => {
    function handleClick(e) {
      if (ref.current && !ref.current.contains(e.target)) onClose()
    }
    document.addEventListener('mousedown', handleClick)
    return () => document.removeEventListener('mousedown', handleClick)
  }, [onClose])

  function downloadCSV() {
    const headers = ['SL', 'Activity Name', 'Description', 'Status', 'Progress', ...dates]
    const rows = tasks.map(t => [
      t.sl,
      `"${(t.name || '').replace(/"/g, '""')}"`,
      `"${(t.description || '').replace(/"/g, '""')}"`,
      t.status,
      t.progress,
      ...dates.map(d => isHighlighted(highlights, t.id, d) ? '1' : '0'),
    ])
    const csv = [headers.join(','), ...rows.map(r => r.join(','))].join('\n')
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${project.name.replace(/[^a-z0-9]/gi, '_')}.csv`
    a.click()
    URL.revokeObjectURL(url)
    onClose()
  }

  async function downloadExcel() {
    const wb = new ExcelJS.Workbook()
    const sheetName = project.name.slice(0, 31).replace(/[*?:/\\[\]]/g, '')
    const ws = wb.addWorksheet(sheetName || 'Project')

    const months = groupByMonth(dates)
    const dateStartCol = 6

    // Row 1 — month groups across date columns
    const monthRow = ws.getRow(1)
    let col = dateStartCol
    months.forEach(mg => {
      const label = new Date(mg.dates[0] + 'T00:00:00')
        .toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })
        .toUpperCase()
      const startCol = col
      const endCol = col + mg.dates.length - 1
      monthRow.getCell(startCol).value = label
      if (mg.dates.length > 1) {
        ws.mergeCells(1, startCol, 1, endCol)
      }
      const cell = monthRow.getCell(startCol)
      cell.font = { bold: true }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5E7EB' } }
      col = endCol + 1
    })

    // Row 2 — fixed columns + day numbers only
    const dayRow = ws.getRow(2)
    dayRow.getCell(1).value = 'SL No'
    dayRow.getCell(2).value = 'Activity Name'
    dayRow.getCell(3).value = 'Description'
    dayRow.getCell(4).value = 'Status'
    dayRow.getCell(5).value = 'Progress %'
    dates.forEach((ds, i) => {
      const cell = dayRow.getCell(dateStartCol + i)
      cell.value = new Date(ds + 'T00:00:00').getDate()
      cell.numFmt = '0'
      cell.dataValidation = undefined
    })
    dayRow.font = { bold: true }
    dayRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5E7EB' } }
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } } }
    })

    tasks.forEach((t) => {
      const row = ws.addRow([
        t.sl,
        t.name,
        t.description || '',
        STATUS[t.status]?.label || t.status,
        t.progress,
        ...dates.map(d => isHighlighted(highlights, t.id, d) ? '●' : ''),
      ])
      dates.forEach((d, di) => {
        if (isHighlighted(highlights, t.id, d)) {
          const cell = row.getCell(dateStartCol + di)
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B82F6' } }
          cell.font = { color: { argb: 'FFFFFFFF' } }
          cell.alignment = { horizontal: 'center' }
        }
      })
    })

    ws.columns.forEach((col, i) => {
      col.width = i < 5 ? (i === 2 ? 30 : i === 1 ? 24 : 14) : 5
    })

    const buffer = await wb.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${project.name.replace(/[^a-z0-9]/gi, '_')}.xlsx`
    a.click()
    URL.revokeObjectURL(url)
    onClose()
  }

  async function downloadPDF() {
    await downloadProjectPDF(project, tasks, dates, highlights)
    onClose()
  }

  const items = [
    { label: 'Download as PDF', action: downloadPDF },
    { label: 'Download as Excel (.xlsx)', action: downloadExcel },
    { label: 'Download as CSV', action: downloadCSV },
  ]

  return (
    <div ref={ref} className="absolute right-0 top-full mt-1 w-52 bg-cream-50 dark:bg-gray-800 border border-gray-200 dark:border-gray-600 rounded-xl shadow-lg z-30 overflow-hidden">
      {items.map(item => (
        <button
          key={item.label}
          onClick={item.action}
          className="w-full text-left px-4 py-2.5 text-sm text-gray-700 dark:text-gray-200 hover:bg-cream-100 dark:hover:bg-gray-700 transition-colors"
        >
          {item.label}
        </button>
      ))}
    </div>
  )
}

export default function ProjectPage() {
  const router = useRouter()
  const params = useParams()
  const id = params.id

  const [project, setProject] = useState(null)
  const [tasks, setTasks] = useState([])
  const [highlights, setHighlights] = useState(new Set())
  const [loading, setLoading] = useState(true)
  const [editTask, setEditTask] = useState(null)
  const [editProject, setEditProject] = useState(false)
  const [downloadOpen, setDownloadOpen] = useState(false)
  const [darkMode, setDarkMode] = useState(false)

  useEffect(() => {
    const saved = localStorage.getItem('theme')
    if (saved === 'dark') setDarkMode(true)
  }, [])

  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add('dark')
      localStorage.setItem('theme', 'dark')
    } else {
      document.documentElement.classList.remove('dark')
      localStorage.setItem('theme', 'light')
    }
  }, [darkMode])

  const load = useCallback(async () => {
    setLoading(true)
    const [{ data: proj }, { data: tsk }, { data: hl }] = await Promise.all([
      supabase.from('projects').select('*').eq('id', id).single(),
      supabase.from('tasks').select('*').eq('project_id', id).order('sl'),
      supabase.from('highlights').select('task_id, date').eq('project_id', id),
    ])
    setProject(proj)
    setTasks(tsk || [])
    setHighlights(new Set((hl || []).map(h => hlKey(h.task_id, h.date))))
    setLoading(false)
  }, [id])

  useEffect(() => { load() }, [load])

  async function toggleHighlight(taskId, ds) {
    const key = hlKey(taskId, ds)
    const next = new Set(highlights)
    if (next.has(key)) {
      next.delete(key)
      await supabase.from('highlights').delete().eq('project_id', id).eq('task_id', taskId).eq('date', ds)
    } else {
      next.add(key)
      await supabase.from('highlights').upsert({ project_id: id, task_id: taskId, date: ds })
    }
    setHighlights(next)
  }

  async function saveTask(form) {
    const payload = taskPayload(form)
    if (editTask === 'new') {
      const { data } = await supabase.from('tasks')
        .insert({ project_id: id, sl: tasks.length + 1, ...payload }).select().single()
      if (data) setTasks(ts => [...ts, data])
    } else {
      const { data } = await supabase.from('tasks')
        .update(payload).eq('id', editTask.id).select().single()
      if (data) setTasks(ts => ts.map(t => t.id === editTask.id ? data : t))
    }
  }

  async function saveProject(form) {
    const payload = {
      name: form.name.trim(),
      location: form.location.trim() || null,
      start_date: form.start_date || null,
      end_date: form.end_date || null,
    }
    const { data } = await supabase.from('projects').update(payload).eq('id', id).select().single()
    if (data) setProject(data)
  }

  async function deleteTask(taskId) {
    await supabase.from('tasks').delete().eq('id', taskId)
    const remaining = tasks.filter(t => t.id !== taskId).map((t, i) => ({ ...t, sl: i + 1 }))
    setTasks(remaining)
    await Promise.all(remaining.map(t => supabase.from('tasks').update({ sl: t.sl }).eq('id', t.id)))
    setHighlights(prev => new Set([...prev].filter(k => !k.startsWith(`${taskId}|`))))
  }

  if (loading) return (
    <div className="min-h-screen bg-cream-100 dark:bg-gray-900 flex items-center justify-center text-sm text-gray-400">
      Loading project…
    </div>
  )

  if (!project) return (
    <div className="min-h-screen bg-cream-100 dark:bg-gray-900 flex flex-col items-center justify-center gap-4">
      <p className="text-sm text-gray-400">Project not found.</p>
      <button onClick={() => router.push('/')} className="text-sm text-blue-600 underline no-print">← Back</button>
    </div>
  )

  const dates = getProjectDates(project)
  const hasProjectDates = Boolean(project.start_date && project.end_date)
  const highlightCount = highlights.size

  const stats = [
    { label: 'Total Tasks', val: tasks.length, cls: 'text-gray-900 dark:text-white' },
    { label: 'In Progress', val: tasks.filter(t => t.status === 'in-progress').length, cls: 'text-blue-600 dark:text-blue-400' },
    { label: 'Completed', val: tasks.filter(t => t.status === 'done').length, cls: 'text-green-600 dark:text-green-400' },
    { label: 'Delayed', val: tasks.filter(t => t.status === 'delayed').length, cls: 'text-red-500 dark:text-red-400' },
  ]

  return (
    <div className="min-h-screen bg-cream-100 dark:bg-gray-900">
      <ThemeToggle darkMode={darkMode} onToggle={() => setDarkMode(d => !d)} />

      <div className="max-w-screen-xl mx-auto px-4 sm:px-6 py-10 print:py-4 print:px-2">
        <button onClick={() => router.push(`/department/${project.department || 'corporate'}`)}
          className="no-print text-sm text-gray-400 hover:text-gray-600 dark:hover:text-gray-300 mb-5 flex items-center gap-1 transition-colors">
          ← All Projects
        </button>

        <div className="flex items-start justify-between mb-8 print:mb-4">
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h1 className="text-2xl font-semibold text-gray-900 dark:text-white tracking-tight print:text-xl">{project.name}</h1>
              <span className={`no-print text-xs px-2.5 py-0.5 rounded-full font-medium ${DEPARTMENT_BADGE[project.department || 'corporate']}`}>
                {deptLabel(project.department)}
              </span>
              <button
                onClick={() => setEditProject(true)}
                className="no-print w-8 h-8 rounded-lg border border-gray-200 dark:border-gray-600 text-gray-400 dark:text-gray-500 hover:text-gray-600 dark:hover:text-gray-300 hover:bg-cream-100 dark:hover:bg-gray-700 flex items-center justify-center transition-colors flex-shrink-0"
                title="Edit project"
                aria-label="Edit project"
              >
                ✎
              </button>
            </div>
            {project.location && (
              <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">📍 {project.location}</p>
            )}
            {hasProjectDates && (
              <p className="text-sm text-gray-400 mt-1">{fmtFull(project.start_date)} – {fmtFull(project.end_date)}</p>
            )}
          </div>
          <div className="no-print flex items-center gap-2 flex-shrink-0">
            <div className="relative">
              <button
                onClick={() => setDownloadOpen(o => !o)}
                className="text-sm px-4 py-2.5 rounded-xl border border-gray-200 dark:border-gray-600 text-gray-700 dark:text-gray-200 hover:bg-cream-100 dark:hover:bg-gray-700 font-medium transition-colors"
              >
                Download ↓
              </button>
              {downloadOpen && (
                <DownloadMenu
                  project={project}
                  tasks={tasks}
                  dates={dates}
                  highlights={highlights}
                  onClose={() => setDownloadOpen(false)}
                />
              )}
            </div>
            <button onClick={() => setEditTask('new')}
              className="bg-blue-600 text-white text-sm px-4 py-2.5 rounded-xl hover:bg-blue-700 font-medium">
              + Add Task
            </button>
          </div>
        </div>

        <div className="no-print grid grid-cols-2 sm:grid-cols-4 gap-3 mb-10">
          {stats.map(s => (
            <div key={s.label} className="bg-cream-50 dark:bg-gray-800 border border-gray-200 dark:border-gray-600 rounded-2xl p-4">
              <p className="text-xs text-gray-400 mb-1">{s.label}</p>
              <p className={`text-3xl font-semibold leading-none ${s.cls}`}>{s.val}</p>
            </div>
          ))}
        </div>

        {tasks.length > 0 && (
          <div className="no-print flex items-center justify-between flex-wrap gap-3 mb-4">
            <div className="flex gap-5 flex-wrap items-center">
              <span className="flex items-center gap-1.5 text-xs text-gray-500 dark:text-gray-400">
                <span className="w-4 h-3 rounded-sm inline-block bg-blue-500" /> Highlighted cell
              </span>
              <span className="flex items-center gap-1.5 text-xs text-red-400">
                <span className="font-bold text-xs">Sun</span> Non-working day
              </span>
            </div>
            <div className="flex items-center gap-3">
              {highlightCount > 0 && (
                <span className="text-xs text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/30 border border-blue-200 dark:border-blue-800 px-3 py-1 rounded-full font-medium">
                  {highlightCount} cell{highlightCount !== 1 ? 's' : ''} highlighted
                </span>
              )}
              <span className="text-xs text-gray-400">Click activity to edit · Click any date cell to highlight</span>
            </div>
          </div>
        )}

        {tasks.length > 0 && hasProjectDates ? (
          <GanttChart
            tasks={tasks}
            dates={dates}
            highlights={highlights}
            onToggleHighlight={toggleHighlight}
            onEditTask={setEditTask}
          />
        ) : tasks.length > 0 ? (
          <div className="text-center py-20 border border-dashed border-gray-200 dark:border-gray-600 rounded-2xl text-gray-400 text-sm">
            Set a project start and end date to see the Gantt chart
          </div>
        ) : (
          <div className="text-center py-20 border border-dashed border-gray-200 dark:border-gray-600 rounded-2xl text-gray-400 text-sm">
            No activities yet — click &quot;+ Add Task&quot; to start scheduling.
          </div>
        )}
      </div>

      {editProject && (
        <ProjectModal
          project={project}
          onSave={saveProject}
          onClose={() => setEditProject(false)}
        />
      )}

      {editTask && (
        <TaskModal
          task={editTask === 'new' ? null : editTask}
          isNew={editTask === 'new'}
          onSave={saveTask}
          onDelete={editTask !== 'new' ? () => deleteTask(editTask.id) : undefined}
          onClose={() => setEditTask(null)}
        />
      )}
    </div>
  )
}
