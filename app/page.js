'use client'
import { useState, useEffect, useCallback } from 'react'
import { useRouter } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import ThemeToggle from '@/components/ThemeToggle'

const TILE_ACCENTS = ['#60A5FA', '#4ADE80', '#F59E0B', '#A78BFA', '#F472B6', '#FB7185']

function toSlug(name) {
  return name.trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '')
}

function toDisplayName(name) {
  return name.trim().split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ')
}

function DepartmentTile({ department, projectCount, accent, onClick }) {
  return (
    <button
      onClick={onClick}
      className="bg-cream-50 dark:bg-gray-800 rounded-2xl border border-gray-200 dark:border-gray-600 p-6 min-h-[160px] w-full text-left cursor-pointer hover:border-blue-300 dark:hover:border-blue-500 hover:shadow-md transition-all relative overflow-hidden group flex flex-col justify-between"
    >
      <div className="absolute top-0 left-0 bottom-0 w-1 rounded-l-2xl" style={{ background: accent }} />
      <div className="pl-3">
        <h2 className="text-lg font-bold text-gray-900 dark:text-white">{department.display_name}</h2>
        <p className="text-sm text-gray-400 mt-2">
          {projectCount} project{projectCount !== 1 ? 's' : ''}
        </p>
      </div>
      <span className="text-gray-300 dark:text-gray-600 text-xl self-end group-hover:text-blue-400 transition-colors pl-3">›</span>
    </button>
  )
}

export default function Home() {
  const router = useRouter()
  const [departments, setDepartments] = useState([])
  const [projectCounts, setProjectCounts] = useState({})
  const [loading, setLoading] = useState(true)
  const [adding, setAdding] = useState(false)
  const [newDeptName, setNewDeptName] = useState('')
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
    const [{ data: depts }, { data: projs }] = await Promise.all([
      supabase.from('departments').select('*').order('display_name'),
      supabase.from('projects').select('department'),
    ])
    const counts = {}
    ;(projs || []).forEach(p => {
      const d = p.department || 'corporate'
      counts[d] = (counts[d] || 0) + 1
    })
    setDepartments(depts || [])
    setProjectCounts(counts)
    setLoading(false)
  }, [])

  useEffect(() => { load() }, [load])

  async function addDepartment() {
    const slug = toSlug(newDeptName)
    if (!slug) return
    setSaving(true)
    const { data, error } = await supabase.from('departments').insert({
      name: slug,
      display_name: toDisplayName(newDeptName),
    }).select().single()
    setSaving(false)
    if (error) {
      alert(error.message.includes('unique') ? 'Department already exists.' : error.message)
      return
    }
    if (data) {
      setDepartments(ds => [...ds, data].sort((a, b) => a.display_name.localeCompare(b.display_name)))
      setProjectCounts(c => ({ ...c, [data.name]: 0 }))
    }
    setNewDeptName('')
    setAdding(false)
  }

  return (
    <div className="min-h-screen bg-cream-100 dark:bg-gray-900">
      <ThemeToggle darkMode={darkMode} onToggle={() => setDarkMode(d => !d)} />

      <div className="max-w-5xl mx-auto px-4 sm:px-6 py-10">
        <div className="mb-10 pb-6 border-b border-gray-200 dark:border-gray-600 mr-12">
          <h1 className="text-2xl font-semibold text-gray-900 dark:text-white tracking-tight">Project Scheduler</h1>
          <p className="text-sm text-gray-400 mt-1">Ashly Group of Comapnies</p>
        </div>

        {loading ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
            {[1, 2, 3].map(i => (
              <div key={i} className="bg-cream-50 dark:bg-gray-800 rounded-2xl border border-gray-200 dark:border-gray-600 min-h-[160px] animate-pulse" />
            ))}
          </div>
        ) : (
          <>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 mb-6">
              {departments.map((dept, i) => (
                <DepartmentTile
                  key={dept.id}
                  department={dept}
                  projectCount={projectCounts[dept.name] || 0}
                  accent={TILE_ACCENTS[i % TILE_ACCENTS.length]}
                  onClick={() => router.push(`/department/${dept.name}`)}
                />
              ))}
            </div>

            {adding ? (
              <div className="bg-cream-50 dark:bg-gray-800 rounded-2xl border border-gray-200 dark:border-gray-600 p-5 max-w-sm flex flex-col gap-3">
                <p className="text-sm font-semibold text-gray-800 dark:text-white">New Department</p>
                <input
                  value={newDeptName}
                  onChange={e => setNewDeptName(e.target.value)}
                  placeholder="Department name…"
                  autoFocus
                  onKeyDown={e => { if (e.key === 'Enter') addDepartment(); if (e.key === 'Escape') { setAdding(false); setNewDeptName('') } }}
                  className="border border-gray-200 dark:border-gray-600 dark:bg-gray-700 dark:text-white rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
                <div className="flex gap-2">
                  <button
                    onClick={addDepartment}
                    disabled={saving || !newDeptName.trim()}
                    className="flex-1 bg-blue-600 text-white text-sm py-2 rounded-xl hover:bg-blue-700 disabled:opacity-50 font-medium"
                  >
                    {saving ? 'Saving…' : 'Save'}
                  </button>
                  <button
                    onClick={() => { setAdding(false); setNewDeptName('') }}
                    className="px-3 py-2 text-sm text-gray-500 dark:text-gray-400 border border-gray-200 dark:border-gray-600 rounded-xl hover:bg-cream-100 dark:hover:bg-gray-700"
                  >
                    Cancel
                  </button>
                </div>
              </div>
            ) : (
              <button
                onClick={() => setAdding(true)}
                className="text-sm text-gray-500 dark:text-gray-400 border border-dashed border-gray-300 dark:border-gray-600 rounded-xl px-4 py-2.5 hover:border-blue-400 dark:hover:border-blue-500 hover:text-blue-500 transition-colors"
              >
                + Add Department
              </button>
            )}
          </>
        )}
      </div>
    </div>
  )
}
