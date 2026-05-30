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


  async function downloadExcel() {
    let logoBase64 = null
    try {
      const logoResponse = await fetch('/ashly_logo.png')
      if (logoResponse.ok) {
        const logoBlob = await logoResponse.blob()
        logoBase64 = await new Promise((resolve) => {
          const reader = new FileReader()
          reader.onloadend = () => {
            const result = reader.result
            resolve(typeof result === 'string' ? result.split(',')[1] : null)
          }
          reader.onerror = () => resolve(null)
          reader.readAsDataURL(logoBlob)
        })
      }
    } catch {
      logoBase64 = null
    }

    const thinBorder = { style: 'thin', color: { argb: 'FFB0BEC5' } }
    const allSidesThin = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder }
    const whiteBorder = { style: 'thin', color: { argb: 'FFFFFFFF' } }
    const allSidesWhite = { top: whiteBorder, left: whiteBorder, bottom: whiteBorder, right: whiteBorder }

    const wb = new ExcelJS.Workbook()
    const sheetName = project.name.slice(0, 31).replace(/[*?:/\\[\]]/g, '')
    const ws = wb.addWorksheet(sheetName || 'Project')

    const months = groupByMonth(dates)
    const dateStartCol = 6
    const lastCol = Math.max(dateStartCol + dates.length - 1, 5)

    if (logoBase64) {
      const logoId = wb.addImage({ base64: logoBase64, extension: 'png' })
      ws.addImage(logoId, {
        tl: { col: 0, row: 0 },
        br: { col: 2, row: 6 },
      })
    }

    for (let r = 1; r <= 6; r++) {
      ws.getRow(r).height = 20
    }

    const companyRow = ws.getRow(7)
    companyRow.height = 18
    companyRow.getCell(1).value = 'ASHLY GROUP OF COMPANIES'
    companyRow.getCell(1).font = { bold: true, size: 14, color: { argb: 'FF1F3864' } }
    ws.mergeCells(7, 1, 7, 5)

    const projectRow = ws.getRow(8)
    projectRow.height = 18
    projectRow.getCell(1).value = project.name
    projectRow.getCell(1).font = { bold: true, size: 12, color: { argb: 'FF1F3864' } }
    ws.mergeCells(8, 1, 8, 5)

    const locationRow = ws.getRow(9)
    locationRow.height = 18
    locationRow.getCell(1).value = project.location || ''
    locationRow.getCell(1).font = { size: 10, color: { argb: 'FF444444' } }
    ws.mergeCells(9, 1, 9, 5)

    const dateRangeStr = project.start_date && project.end_date
      ? `${fmtFull(project.start_date)} – ${fmtFull(project.end_date)}`
      : ''
    const dateRangeRow = ws.getRow(10)
    dateRangeRow.height = 18
    dateRangeRow.getCell(1).value = dateRangeStr
    dateRangeRow.getCell(1).font = { size: 10, color: { argb: 'FF444444' } }
    ws.mergeCells(10, 1, 10, 5)

    // Row 11 — empty separator

    const monthRowNum = 12
    const dayRowNum = 13
    const monthRow = ws.getRow(monthRowNum)
    monthRow.height = 20

    let col = dateStartCol
    months.forEach(mg => {
      const label = new Date(mg.dates[0] + 'T00:00:00')
        .toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })
        .toUpperCase()
      const startCol = col
      const endCol = col + mg.dates.length - 1
      monthRow.getCell(startCol).value = label
      if (mg.dates.length > 1) {
        ws.mergeCells(monthRowNum, startCol, monthRowNum, endCol)
      }
      const cell = monthRow.getCell(startCol)
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } }
      cell.border = allSidesWhite
      col = endCol + 1
    })

    const dayRow = ws.getRow(dayRowNum)
    dayRow.height = 16

    const fixedHeaders = ['SL No', 'Activity Name', 'Description', 'Status', 'Progress %']
    fixedHeaders.forEach((label, i) => {
      const cell = dayRow.getCell(i + 1)
      cell.value = label
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } }
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 }
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      cell.border = allSidesThin
    })

    dates.forEach((ds, i) => {
      const cell = dayRow.getCell(dateStartCol + i)
      cell.value = new Date(ds + 'T00:00:00').getDate()
      cell.numFmt = '0'
      cell.dataValidation = undefined
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E5FA3' } }
      cell.font = {
        bold: true,
        color: { argb: isWeekend(ds) ? 'FFFF6B6B' : 'FFFFFFFF' },
        size: 9,
      }
      cell.alignment = { horizontal: 'center' }
      cell.border = allSidesThin
    })

    tasks.forEach((t, taskIdx) => {
      const rowNum = 14 + taskIdx
      const row = ws.getRow(rowNum)
      row.height = 32
      const isEven = rowNum % 2 === 0
      const rowFill = isEven
        ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }
        : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EDF5' } }

      row.getCell(1).value = t.sl
      row.getCell(2).value = t.name
      row.getCell(3).value = t.description || ''
      row.getCell(4).value = STATUS[t.status]?.label || t.status
      row.getCell(5).value = t.progress

      for (let c = 1; c <= 5; c++) {
        const cell = row.getCell(c)
        cell.fill = rowFill
        cell.font = {
          size: 9,
          color: { argb: 'FF222222' },
          bold: c === 2,
        }
        cell.border = allSidesThin
        if (c === 3) cell.alignment = { wrapText: true, vertical: 'top' }
      }

      dates.forEach((d, di) => {
        const cell = row.getCell(dateStartCol + di)
        cell.border = allSidesThin
        if (isHighlighted(highlights, t.id, d)) {
          cell.value = '●'
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } }
          cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 9 }
          cell.alignment = { horizontal: 'center', vertical: 'middle' }
        } else if (isWeekend(d)) {
          cell.value = ''
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4E4' } }
          cell.font = { size: 9, color: { argb: 'FF222222' } }
        } else {
          cell.value = ''
          cell.fill = rowFill
          cell.font = { size: 9, color: { argb: 'FF222222' } }
        }
      })
    })

    ws.getColumn(1).width = 6
    ws.getColumn(2).width = 22
    ws.getColumn(3).width = 35
    ws.getColumn(4).width = 14
    ws.getColumn(5).width = 10
    for (let i = dateStartCol; i <= lastCol; i++) {
      ws.getColumn(i).width = 4.5
    }

    ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 13 }]

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

  const items = [
    { label: 'Download as Excel (.xlsx)', action: downloadExcel }
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
