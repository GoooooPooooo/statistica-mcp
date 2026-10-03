export function requirePath(a) {
  // In attach mode a path is optional: the worker then uses the active sheet.
  if (a.attach && (a.path === undefined || a.path === null || a.path === '')) return undefined
  if (typeof a.path !== 'string' || a.path.trim() === '') throw new Error('`path` is required')
  return a.path
}
