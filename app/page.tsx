'use client'

import { useState, useEffect, useRef, useCallback } from 'react'

interface Entry {
  id: string
  group: string
  imageUrl: string
  date: string
}

type Tab = 'upload' | 'gallery' | 'generate'

export default function Home() {
  const [tab, setTab] = useState<Tab>('upload')
  const [entries, setEntries] = useState<Entry[]>([])
  const [loadingEntries, setLoadingEntries] = useState(false)

  // Upload state
  const [file, setFile] = useState<File | null>(null)
  const [preview, setPreview] = useState<string | null>(null)
  const [group, setGroup] = useState('')
  const [saving, setSaving] = useState(false)
  const [saveMsg, setSaveMsg] = useState('')
  const [dragOver, setDragOver] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  // Generate state
  const [generating, setGenerating] = useState(false)
  const [genMsg, setGenMsg] = useState('')

  const fetchEntries = useCallback(async () => {
    setLoadingEntries(true)
    try {
      const res = await fetch('/api/entries', { cache: 'no-store' })
      const data = await res.json()
      setEntries(data)
    } catch {}
    setLoadingEntries(false)
  }, [])

  useEffect(() => { fetchEntries() }, [fetchEntries])

  function handleFile(f: File) {
    setFile(f)
    const reader = new FileReader()
    reader.onload = (e) => setPreview(e.target?.result as string)
    reader.readAsDataURL(f)
  }

  function clearFile() {
    setFile(null)
    setPreview(null)
    if (fileInputRef.current) fileInputRef.current.value = ''
  }

  async function handleSave() {
    if (!file || !group.trim()) return
    setSaving(true)
    setSaveMsg('')
    try {
      const fd = new FormData()
      fd.append('file', file)
      fd.append('group', group.trim())
      const res = await fetch('/api/upload', { method: 'POST', body: fd })
      if (!res.ok) throw new Error()
      const entry = await res.json()
      setEntries(prev => [entry, ...prev])
      setSaveMsg('Saved!')
      clearFile()
      setGroup('')
    } catch {
      setSaveMsg('Upload failed. Try again.')
    }
    setSaving(false)
    setTimeout(() => setSaveMsg(''), 3000)
  }

  async function handleDelete(entry: Entry) {
    if (!confirm(`Delete entry for "${entry.group}"?`)) return
    try {
      await fetch('/api/delete', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: entry.id, imageUrl: entry.imageUrl }),
      })
      setEntries(prev => prev.filter(e => e.id !== entry.id))
    } catch {}
  }

  async function handleGenerate() {
    if (!entries.length) return
    setGenerating(true)
    setGenMsg('Loading library…')

    // Dynamically import pptxgenjs
    const PptxGenJS = (await import('pptxgenjs')).default
    const pptx = new PptxGenJS()
    pptx.layout = 'LAYOUT_WIDE'

    setGenMsg('Fetching images…')
    for (const entry of entries) {
      // Convert image URL to base64
      let imgData: string
      try {
        const res = await fetch(entry.imageUrl)
        const buf = await res.arrayBuffer()
        const b64 = btoa(String.fromCharCode(...new Uint8Array(buf)))
        const mime = entry.imageUrl.match(/\.(png)$/i) ? 'image/png' : 'image/jpeg'
        imgData = `data:${mime};base64,${b64}`
      } catch {
        imgData = entry.imageUrl
      }

      const slide = pptx.addSlide()
      slide.background = { color: 'F7F7F6' }

      // Header bar
      slide.addShape('rect' as any, {
        x: 0, y: 0, w: 13.33, h: 0.72,
        fill: { color: '1D3461' },
        line: { color: '1D3461' },
      })

      slide.addText(entry.group, {
        x: 0.4, y: 0.1, w: 12.5, h: 0.52,
        fontSize: 24, bold: true, color: 'FFFFFF', fontFace: 'Calibri',
        valign: 'middle',
      })

      // Photo
      slide.addImage({
        data: imgData,
        x: 1.2, y: 0.9, w: 10.93, h: 5.5,
        sizing: { type: 'contain', w: 10.93, h: 5.5 },
      })

      // Footer
      const d = new Date(entry.date).toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric',
      })
      slide.addText(`Uploaded: ${d}`, {
        x: 0.4, y: 6.85, w: 12.5, h: 0.25,
        fontSize: 10, color: '999999', fontFace: 'Calibri',
      })
    }

    setGenMsg('Downloading…')
    await pptx.writeFile({ fileName: 'group-photos.pptx' })
    setGenMsg(`✓ Downloaded group-photos.pptx (${entries.length} slides)`)
    setGenerating(false)
  }

  return (
    <div style={{ maxWidth: 720, margin: '0 auto', padding: '2rem 1rem' }}>
      {/* Header */}
      <div style={{ marginBottom: '2rem' }}>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: '#1d3461', letterSpacing: '-0.3px' }}>
          📸 Group Photo Uploader
        </h1>
        <p style={{ color: '#6b6a67', fontSize: 14, marginTop: 4 }}>
          Upload group photos, view all entries, and export to PowerPoint.
        </p>
      </div>

      {/* Tabs */}
      <div style={{ display: 'flex', borderBottom: '1.5px solid var(--border)', marginBottom: '1.5rem', gap: 0 }}>
        {(['upload', 'gallery', 'generate'] as Tab[]).map(t => (
          <button
            key={t}
            onClick={() => { setTab(t); if (t === 'gallery') fetchEntries() }}
            style={{
              padding: '9px 20px',
              fontSize: 14,
              fontWeight: tab === t ? 600 : 400,
              color: tab === t ? '#1d3461' : '#6b6a67',
              background: 'none',
              border: 'none',
              borderBottom: tab === t ? '2px solid #1d3461' : '2px solid transparent',
              cursor: 'pointer',
              textTransform: 'capitalize',
              transition: 'all 0.15s',
              marginBottom: -1.5,
            }}
          >
            {t === 'upload' ? '⬆ Upload' : t === 'gallery' ? `🖼 Gallery (${entries.length})` : '📊 Generate'}
          </button>
        ))}
      </div>

      {/* Upload Panel */}
      {tab === 'upload' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div
            onClick={() => fileInputRef.current?.click()}
            onDragOver={e => { e.preventDefault(); setDragOver(true) }}
            onDragLeave={() => setDragOver(false)}
            onDrop={e => {
              e.preventDefault(); setDragOver(false)
              const f = e.dataTransfer.files[0]
              if (f?.type.startsWith('image/')) handleFile(f)
            }}
            style={{
              border: `2px dashed ${dragOver ? '#1d3461' : '#d1d0cd'}`,
              borderRadius: var_radius,
              padding: '2.5rem',
              textAlign: 'center',
              cursor: 'pointer',
              background: dragOver ? '#e8eef6' : '#fafaf9',
              transition: 'all 0.15s',
            }}
          >
            <div style={{ fontSize: 40, marginBottom: 8 }}>📁</div>
            <p style={{ fontWeight: 600, fontSize: 15 }}>Click to upload or drag & drop</p>
            <p style={{ color: '#6b6a67', fontSize: 13, marginTop: 4 }}>PNG, JPG, WEBP supported</p>
          </div>
          <input ref={fileInputRef} type="file" accept="image/*" style={{ display: 'none' }}
            onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f) }} />

          {preview && (
            <div style={{ position: 'relative', borderRadius: var_radius, overflow: 'hidden', border: '1px solid var(--border)' }}>
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img src={preview} alt="Preview" style={{ width: '100%', maxHeight: 260, objectFit: 'contain', display: 'block', background: '#f0f0ef' }} />
              <button
                onClick={clearFile}
                style={{ position: 'absolute', top: 8, right: 8, background: 'rgba(0,0,0,0.6)', border: 'none', borderRadius: '50%', width: 30, height: 30, cursor: 'pointer', color: '#fff', fontSize: 16, display: 'flex', alignItems: 'center', justifyContent: 'center' }}
              >✕</button>
            </div>
          )}

          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: var_radius, padding: '1rem' }}>
            <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#6b6a67', marginBottom: 6 }}>
              GROUP NUMBER / IDENTIFIER
            </label>
            <input
              type="text"
              value={group}
              onChange={e => setGroup(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && handleSave()}
              placeholder="e.g. Group 3, Section A, Team Alpha"
              style={{ width: '100%', padding: '10px 12px', border: '1px solid var(--border)', borderRadius: var_radius_sm, fontSize: 15, outline: 'none' }}
            />
          </div>

          <button
            onClick={handleSave}
            disabled={!file || !group.trim() || saving}
            style={{
              width: '100%', padding: '12px', fontSize: 15, fontWeight: 600,
              background: (!file || !group.trim() || saving) ? '#e5e5e3' : '#1d3461',
              color: (!file || !group.trim() || saving) ? '#999' : '#fff',
              border: 'none', borderRadius: var_radius, cursor: (!file || !group.trim() || saving) ? 'not-allowed' : 'pointer',
              transition: 'all 0.15s',
            }}
          >
            {saving ? 'Saving…' : '💾 Save to database'}
          </button>
          {saveMsg && (
            <p style={{ textAlign: 'center', color: saveMsg.includes('failed') ? '#dc2626' : '#16a34a', fontSize: 14, fontWeight: 500 }}>
              {saveMsg}
            </p>
          )}
        </div>
      )}

      {/* Gallery Panel */}
      {tab === 'gallery' && (
        <div>
          {loadingEntries ? (
            <p style={{ textAlign: 'center', color: '#6b6a67', padding: '3rem' }}>Loading…</p>
          ) : entries.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '4rem 1rem', color: '#6b6a67' }}>
              <div style={{ fontSize: 48, marginBottom: 12 }}>🖼</div>
              <p style={{ fontSize: 15 }}>No entries yet. Upload some group photos!</p>
            </div>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', gap: 14 }}>
              {entries.map(entry => (
                <div key={entry.id} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: var_radius, overflow: 'hidden', position: 'relative' }}>
                  {/* eslint-disable-next-line @next/next/no-img-element */}
                  <img src={entry.imageUrl} alt={entry.group} style={{ width: '100%', height: 130, objectFit: 'cover', display: 'block' }} />
                  <button
                    onClick={() => handleDelete(entry)}
                    style={{ position: 'absolute', top: 6, right: 6, background: 'rgba(0,0,0,0.55)', border: 'none', borderRadius: '50%', width: 26, height: 26, cursor: 'pointer', color: '#fff', fontSize: 13, display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                    title="Delete"
                  >✕</button>
                  <div style={{ padding: '8px 10px' }}>
                    <span style={{ display: 'inline-block', background: '#e8eef6', color: '#1d3461', fontSize: 12, fontWeight: 600, padding: '2px 8px', borderRadius: 4 }}>
                      {entry.group}
                    </span>
                    <p style={{ fontSize: 11, color: '#9b9a97', marginTop: 4 }}>
                      {new Date(entry.date).toLocaleDateString()}
                    </p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Generate Panel */}
      {tab === 'generate' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: var_radius, padding: '1.25rem' }}>
            <p style={{ fontSize: 14, color: '#6b6a67', marginBottom: 12 }}>
              Compile all saved group photos into a PowerPoint file — one slide per entry with the group label and image.
            </p>
            <div style={{ background: '#f5f5f4', borderRadius: var_radius_sm, padding: '10px 14px', fontSize: 13, color: '#6b6a67' }}>
              {entries.length === 0
                ? 'ℹ No entries saved yet.'
                : `📋 ${entries.length} entr${entries.length === 1 ? 'y' : 'ies'} ready: ${[...new Set(entries.map(e => e.group))].join(', ')}`
              }
            </div>
          </div>

          <button
            onClick={handleGenerate}
            disabled={entries.length === 0 || generating}
            style={{
              width: '100%', padding: '13px', fontSize: 15, fontWeight: 600,
              background: (entries.length === 0 || generating) ? '#e5e5e3' : '#1d3461',
              color: (entries.length === 0 || generating) ? '#999' : '#fff',
              border: 'none', borderRadius: var_radius, cursor: (entries.length === 0 || generating) ? 'not-allowed' : 'pointer',
              transition: 'all 0.15s',
            }}
          >
            {generating ? '⏳ Building…' : '📥 Generate PowerPoint'}
          </button>

          {genMsg && (
            <p style={{ textAlign: 'center', fontSize: 13, color: genMsg.startsWith('✓') ? '#16a34a' : '#6b6a67', fontWeight: genMsg.startsWith('✓') ? 600 : 400 }}>
              {genMsg}
            </p>
          )}
        </div>
      )}
    </div>
  )
}

const var_radius = 10
const var_radius_sm = 6
