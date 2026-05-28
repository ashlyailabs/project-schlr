'use client'

export default function ThemeToggle({ darkMode, onToggle }) {
  return (
    <button
      onClick={onToggle}
      className="no-print fixed top-4 right-4 z-40 w-9 h-9 rounded-xl border border-gray-200 dark:border-gray-600 bg-white dark:bg-gray-800 text-gray-600 dark:text-gray-300 hover:bg-gray-50 dark:hover:bg-gray-700 flex items-center justify-center transition-colors shadow-sm"
      title="Toggle dark mode"
      aria-label="Toggle dark mode"
    >
      {darkMode ? (
        <span className="text-base leading-none">☽</span>
      ) : (
        <span className="text-base leading-none">☀</span>
      )}
    </button>
  )
}
