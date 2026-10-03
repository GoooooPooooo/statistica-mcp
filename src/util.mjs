export function requirePath(a) {
  if (typeof a.path !== 'string' || a.path.trim() === '') throw new Error('`path` is required')
  return a.path
}
