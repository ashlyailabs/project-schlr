'use client'
import { useState, useEffect, useCallback } from 'react'
import { useRouter, useParams } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import ThemeToggle from '@/components/ThemeToggle'

const STATUS = {
  'not-started': { label: 'Not Started', color: 'text-gray-500 dark:text-gray-400',  bg: 'bg-gray-100 dark:bg-gray-700',  bar: 'bg-gray-300',   accent: '#D1D5DB' },
  'in-progress':  { label: 'In Progress', color: 'text-blue-700 dark:text-blue-400',  bg: 'bg-blue-50 dark:bg-blue-900/40',   bar: 'bg-blue-400',   accent: '#60A5FA' },
  'done':         { label: 'Done',        color: 'text-green-700 dark:text-green-400', bg: 'bg-green-50 dark:bg-green-900/40',  bar: 'bg-green-400',  accent: '#4ADE80' },
  'delayed':      { label: 'Delayed',     color: 'text-red-600 dark:text-red-400',   bg: 'bg-red-50 dark:bg-red-900/40',    bar: 'bg-red-400',    accent: '#F87171' },
}

function fmtShort(s) {
  return new Date(s + 'T00:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })
}

function overallStatus(tasks) {
  if (!tasks.length) return 'not-started'
  if (tasks.some(t => t.status === 'delayed'))     return 'delayed'
  if (tasks.some(t => t.status === 'in-progress')) return 'in-progress'
  if (tasks.every(t => t.status === 'done'))        return 'done'
  return 'not-started'
}

function ProjectCard({ project, onClick, onDelete }) {
  const tasks   = project.tasks || []
  const avgProg = tasks.length
    ? Math.round(tasks.reduce((a, t) => a + (t.progress || 0), 0) / tasks.length)
    : 0
  const status = overallStatus(tasks)
  const st     = STATUS[status]

  const counts = {}
  tasks.forEach(t => { counts[t.status] = (counts[t.status] || 0) + 1 })

  const hasDates = project.start_date && project.end_date

  return (
    <div
      onClick={onClick}
      className="bg-white dark:bg-gray-800 rounded-2xl border border-gray-200 dark:border-gray-600 p-5 cursor-pointer hover:border-blue-300 dark:hover:border-blue-500 hover:shadow-sm transition-all relative overflow-hidden group"
    >
      <div className="absolute top-0 left-0 right-0 h-1 rounded-t-2xl" style={{ background: st.accent }} />

      <div className="flex items-start justify-between mt-1 mb-2">
        <h2 className="text-sm font-semibold text-gray-900 dark:text-white leading-snug flex-1 mr-3">{project.name}</h2>
        <div className="flex items-center gap-2 flex-shrink-0">
          <span className={`text-xs px-2.5 py-0.5 rounded-full font-medium ${st.bg} ${st.color}`}>
            {st.label}
          </span>
          <button
            onClick={e => { e.stopPropagation(); onDelete() }}
            className="text-gray-200 dark:text-gray-600 hover:text-red-400 transition-colors text-sm leading-none opacity-0 group-hover:opacity-100"
          >
            ✕
          </button>
        </div>
      </div>

      {(project.location || hasDates) && (
        <div className="text-xs text-gray-400 mb-3 space-y-0.5">
          {project.location && <p>📍 {project.location}</p>}
          {hasDates && (
            <p>{fmtShort(project.start_date)} – {fmtShort(project.end_date)}</p>
          )}
        </div>
      )}

      <p className="text-xs text-gray-400 mb-4">
        {tasks.length} task{tasks.length !== 1 ? 's' : ''}
      </p>

      <div className="mb-4">
        <div className="flex justify-between text-xs text-gray-400 mb-1.5">
          <span>Overall progress</span>
          <span className="font-medium text-gray-600 dark:text-gray-300">{avgProg}%</span>
        </div>
        <div className="h-1.5 bg-gray-100 dark:bg-gray-700 rounded-full overflow-hidden">
          <div className={`h-full rounded-full transition-all ${st.bar}`} style={{ width: `${avgProg}%` }} />
        </div>
      </div>

      <div className="flex items-center justify-between">
        <div className="flex gap-3 flex-wrap">
          {Object.entries(counts).filter(([, v]) => v > 0).map(([k, v]) => (
            <span key={k} className={`text-xs ${STATUS[k].color}`}>
              {v}&nbsp;{k === 'not-started' ? 'pending' : k === 'in-progress' ? 'active' : k}
            </span>
          ))}
        </div>
        <span className="text-gray-300 dark:text-gray-600 text-sm">›</span>
      </div>
    </div>
  )
}

function NewProjectForm({ name, setName, location, setLocation, startDate, setStartDate, endDate, setEndDate, saving, onCreate, onCancel }) {
  const inp = 'border border-gray-200 dark:border-gray-600 dark:bg-gray-700 dark:text-white rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-full'

  return (
    <div className="bg-white dark:bg-gray-800 rounded-2xl border border-gray-200 dark:border-gray-600 p-5 flex flex-col gap-3">
      <p className="text-sm font-semibold text-gray-800 dark:text-white">New Project</p>
      <input
        value={name}
        onChange={e => setName(e.target.value)}
        placeholder="Project name…"
        autoFocus
        onKeyDown={e => { if (e.key === 'Enter') onCreate(); if (e.key === 'Escape') onCancel() }}
        className={inp}
      />
      <input
        value={location}
        onChange={e => setLocation(e.target.value)}
        placeholder="e.g. Dubai Marina, Office 4B"
        className={inp}
      />
      <div className="grid grid-cols-2 gap-2">
        <div>
          <label className="text-xs text-gray-400 block mb-1">Start Date</label>
          <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className={inp} />
        </div>
        <div>
          <label className="text-xs text-gray-400 block mb-1">End Date</label>
          <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className={inp} />
        </div>
      </div>
      <div className="flex gap-2">
        <button
          onClick={onCreate}
          disabled={saving}
          className="flex-1 bg-blue-600 text-white text-sm py-2 rounded-xl hover:bg-blue-700 disabled:opacity-50 font-medium transition-colors"
        >
          {saving ? 'Creating…' : 'Create project'}
        </button>
        <button
          onClick={onCancel}
          className="px-3 py-2 text-sm text-gray-500 dark:text-gray-400 border border-gray-200 dark:border-gray-600 rounded-xl hover:bg-gray-50 dark:hover:bg-gray-700"
        >
          Cancel
        </button>
      </div>
    </div>
  )
}

function AddProjectCard({ onClick }) {
  return (
    <button
      onClick={onClick}
      className="bg-white dark:bg-gray-800 rounded-2xl border border-dashed border-gray-300 dark:border-gray-600 p-5 flex flex-col items-center justify-center gap-3 cursor-pointer hover:border-blue-400 dark:hover:border-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition-all min-h-[160px] text-gray-400 hover:text-blue-500 group"
    >
      <div className="w-10 h-10 rounded-full bg-gray-100 dark:bg-gray-700 group-hover:bg-blue-100 dark:group-hover:bg-blue-900/40 flex items-center justify-center text-xl transition-colors">
        +
      </div>
      <span className="text-sm font-medium">Add new project</span>
    </button>
  )
}

export default function DepartmentPage() {
  const router = useRouter()
  const params = useParams()
  const dept = params.dept

  const [department, setDepartment] = useState(null)
  const [projects, setProjects] = useState([])
  const [loading, setLoading] = useState(true)
  const [adding, setAdding] = useState(false)
  const [name, setName] = useState('')
  const [location, setLocation] = useState('')
  const [startDate, setStartDate] = useState('')
  const [endDate, setEndDate] = useState('')
  const [saving, setSaving] = useState(false)
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
    const [{ data: deptInfo }, { data: projs }, { data: tasks }] = await Promise.all([
      supabase.from('departments').select('*').eq('name', dept).single(),
      supabase.from('projects').select('*').eq('department', dept).order('created_at', { ascending: false }),
      supabase.from('tasks').select('*'),
    ])
    setDepartment(deptInfo)
    const merged = (projs || []).map(p => ({
      ...p,
      tasks: (tasks || []).filter(t => t.project_id === p.id),
    }))
    setProjects(merged)
    setLoading(false)
  }, [dept])

  useEffect(() => { load() }, [load])

  function resetForm() {
    setName('')
    setLocation('')
    setStartDate('')
    setEndDate('')
    setAdding(false)
  }

  async function createProject() {
    if (!name.trim()) return
    setSaving(true)
    const { data } = await supabase.from('projects').insert({
      name: name.trim(),
      location: location.trim() || null,
      start_date: startDate || null,
      end_date: endDate || null,
      department: dept,
    }).select().single()
    setSaving(false)
    if (data) router.push(`/project/${data.id}`)
    resetForm()
  }

  async function deleteProject(id) {
    if (!confirm('Delete this project and all its tasks?')) return
    await supabase.from('projects').delete().eq('id', id)
    setProjects(ps => ps.filter(p => p.id !== id))
  }

  const activeCount = projects.filter(p => overallStatus(p.tasks) === 'in-progress').length
  const delayedCount = projects.filter(p => overallStatus(p.tasks) === 'delayed').length

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50 dark:bg-gray-900 flex items-center justify-center text-sm text-gray-400">
        Loading…
      </div>
    )
  }

  if (!department) {
    return (
      <div className="min-h-screen bg-gray-50 dark:bg-gray-900 flex flex-col items-center justify-center gap-4">
        <p className="text-sm text-gray-400">Department not found.</p>
        <button onClick={() => router.push('/')} className="text-sm text-blue-600 underline">← All Departments</button>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-50 dark:bg-gray-900">
      <ThemeToggle darkMode={darkMode} onToggle={() => setDarkMode(d => !d)} />

      <div className="max-w-5xl mx-auto px-4 sm:px-6 py-10">
        <button onClick={() => router.push('/')}
          className="text-sm text-gray-400 hover:text-gray-600 dark:hover:text-gray-300 mb-5 flex items-center gap-1 transition-colors">
          ← All Departments
        </button>

        <div className="flex items-start justify-between mb-8 pb-6 border-b border-gray-200 dark:border-gray-600">
          <div>
            <h1 className="text-2xl font-semibold text-gray-900 dark:text-white tracking-tight">{department.display_name}</h1>
            <p className="text-sm text-gray-400 mt-1">{projects.length} project{projects.length !== 1 ? 's' : ''}</p>
          </div>
          <div className="flex gap-6 text-right mr-12">
            {activeCount > 0 && (
              <div>
                <div className="text-2xl font-semibold text-blue-600 dark:text-blue-400 leading-none">{activeCount}</div>
                <div className="text-xs text-gray-400 mt-0.5">active</div>
              </div>
            )}
            {delayedCount > 0 && (
              <div>
                <div className="text-2xl font-semibold text-red-500 dark:text-red-400 leading-none">{delayedCount}</div>
                <div className="text-xs text-gray-400 mt-0.5">delayed</div>
              </div>
            )}
          </div>
        </div>

        {projects.length === 0 && !adding && (
          <p className="text-sm text-gray-400 mb-4">No {department.name} projects yet</p>
        )}

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {projects.map(p => (
            <ProjectCard
              key={p.id}
              project={p}
              onClick={() => router.push(`/project/${p.id}`)}
              onDelete={() => deleteProject(p.id)}
            />
          ))}

          {adding ? (
            <NewProjectForm
              name={name}
              setName={setName}
              location={location}
              setLocation={setLocation}
              startDate={startDate}
              setStartDate={setStartDate}
              endDate={endDate}
              setEndDate={setEndDate}
              saving={saving}
              onCreate={createProject}
              onCancel={resetForm}
            />
          ) : (
            <AddProjectCard onClick={() => setAdding(true)} />
          )}
        </div>
      </div>
    </div>
  )
}
