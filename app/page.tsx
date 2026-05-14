'use client'

import { useState, useEffect, useRef, useCallback } from 'react'

interface Entry {
  id: string
  group: string
  imageUrl: string
  date: string
}

type Tab = 'upload' | 'gallery' | 'generate'

const R = '10px'
const RS = '6px'

export default function Home() {
  const [tab, setTab] = useState<Tab>('upload')
  const [entries, setEntries] = useState<Entry[]>([])
  const [loadingEntries, setLoadingEntries] = useState(false)

  const [file, setFile] = useState<File | null>(null)
  const [preview, setPreview] = useState<string | null>(null)
  const [group, setGroup] = useState('')
  const [saving, setSaving] = useState(false)
  const [saveMsg, setSaveMsg] = useState('')
  const [dragOver, setDragOver] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

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

  useEffect(() => {
    fetchEntries()
  }, [fetchEntries])

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

      const res = await fetch('/api/upload', {
        method: 'POST',
        body: fd,
      })

      if (!res.ok) throw new Error()

      const entry = await res.json()
      setEntries((prev) => [entry, ...prev])
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
        body: JSON.stringify({
          id: entry.id,
          imageUrl: entry.imageUrl,
        }),
      })

      setEntries((prev) => prev.filter((e) => e.id !== entry.id))
    } catch {}
  }

  async function handleGenerate() {
    if (!entries.length) return

    setGenerating(true)
    setGenMsg('Loading library…')

    const PptxGenJS = (await import('pptxgenjs')).default
    const pptx = new PptxGenJS()
    pptx.layout = 'LAYOUT_WIDE'

    async function toDataUri(url: string): Promise<string> {
      const res = await fetch(url)
      const buf = await res.arrayBuffer()
      const bytes = new Uint8Array(buf)
      const b64 = btoa(Array.from(bytes, (b) => String.fromCharCode(b)).join(''))
      const mime = url.match(/\.png$/i) ? 'image/png' : 'image/jpeg'
      return `data:${mime};base64,${b64}`
    }

    setGenMsg('Loading poster…')

    let posterData: string | null = null

    try {
      posterData = await toDataUri('/groufie-poster.png')
    } catch {
      posterData = null
    }

    const cover = pptx.addSlide()

    if (posterData) {
      cover.addImage({
        data: posterData,
        x: 0,
        y: 0,
        w: 13.33,
        h: 7.5,
        sizing: {
          type: 'cover',
          w: 13.33,
          h: 7.5,
        },
      })
    } else {
      cover.background = { color: '001A33' }
      cover.addText('MAKE A GROUFIE! 📸', {
        x: 1,
        y: 2.5,
        w: 11.33,
        h: 1.5,
        fontSize: 48,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Impact',
        align: 'center',
      })
    }

    setGenMsg('Building slides…')

    for (const entry of entries) {
      let imgData: string

      try {
        imgData = await toDataUri(entry.imageUrl)
      } catch {
        imgData = entry.imageUrl
      }

      const slide = pptx.addSlide()
      slide.background = { color: '0A0A2E' }

      slide.addShape('rect' as any, {
        x: 0,
        y: 0,
        w: 13.33,
        h: 1.0,
        fill: { color: '00B4D8' },
        line: { color: '00B4D8' },
      })

      slide.addShape('rect' as any, {
        x: 0,
        y: 0.9,
        w: 13.33,
        h: 0.12,
        fill: { color: '00C851' },
        line: { color: '00C851' },
      })

      slide.addShape('ellipse' as any, {
        x: 0.15,
        y: 0.1,
        w: 0.7,
        h: 0.7,
        fill: { color: 'FFD700' },
        line: { color: '007A33', width: 2 },
      })

      slide.addText('😊', {
        x: 0.15,
        y: 0.1,
        w: 0.7,
        h: 0.7,
        fontSize: 20,
        align: 'center',
        valign: 'middle',
        margin: 0,
      })

      slide.addText('MAKE A GROUFIE! 📸', {
        x: 1.0,
        y: 0.08,
        w: 11.0,
        h: 0.48,
        fontSize: 26,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Impact',
        align: 'left',
        valign: 'middle',
        margin: 0,
      })

      slide.addText('The Funniest but Hungriest Accountability Group (AG)', {
        x: 1.0,
        y: 0.56,
        w: 10.0,
        h: 0.3,
        fontSize: 12,
        color: 'FFD700',
        fontFace: 'Calibri',
        align: 'left',
        valign: 'middle',
        italic: true,
        margin: 0,
      })

      slide.addShape('rect' as any, {
        x: 0.4,
        y: 1.18,
        w: 8.5,
        h: 0.58,
        fill: { color: '00C851' },
        line: { color: '007A33', width: 2 },
      })

      slide.addText(entry.group, {
        x: 0.4,
        y: 1.18,
        w: 8.5,
        h: 0.58,
        fontSize: 24,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Impact',
        align: 'center',
        valign: 'middle',
        margin: 0,
      })

      slide.addImage({
        data: imgData,
        x: 0.4,
        y: 1.88,
        w: 8.5,
        h: 5.1,
        sizing: {
          type: 'contain',
          w: 8.5,
          h: 5.1,
        },
      })

      slide.addShape('rect' as any, {
        x: 9.15,
        y: 1.18,
        w: 3.78,
        h: 5.8,
        fill: { color: '0D1547' },
        line: { color: '00B4D8', width: 2 },
      })

      slide.addShape('rect' as any, {
        x: 9.15,
        y: 1.18,
        w: 3.78,
        h: 0.65,
        fill: { color: '00B4D8' },
        line: { color: '00B4D8' },
      })

      slide.addText('HEAD OFFICE', {
        x: 9.15,
        y: 1.2,
        w: 3.78,
        h: 0.3,
        fontSize: 13,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Arial Black',
        align: 'center',
        margin: 0,
      })

      slide.addText('CORPORATE FELLOWSHIP', {
        x: 9.15,
        y: 1.5,
        w: 3.78,
        h: 0.28,
        fontSize: 9,
        bold: true,
        color: 'FFD700',
        fontFace: 'Calibri',
        align: 'center',
        margin: 0,
      })

      const cards = [
        { emoji: '🍕', label: 'Best Food Pic' },
        { emoji: '🍩', label: 'Most Creative' },
        { emoji: '📚', label: 'Hustle Mode' },
        { emoji: '☕', label: 'Coffee Addict' },
        { emoji: '😂', label: 'Funniest AG' },
      ]

      cards.forEach((card, i) => {
        const yPos = 2.05 + i * 0.92

        slide.addShape('rect' as any, {
          x: 9.35,
          y: yPos,
          w: 3.38,
          h: 0.75,
          fill: { color: '152060' },
          line: { color: '00B4D8', width: 1 },
        })

        slide.addText(card.emoji, {
          x: 9.38,
          y: yPos + 0.05,
          w: 0.65,
          h: 0.65,
          fontSize: 22,
          align: 'center',
          valign: 'middle',
          margin: 0,
        })

        slide.addText(card.label, {
          x: 10.08,
          y: yPos + 0.1,
          w: 2.5,
          h: 0.55,
          fontSize: 13,
          bold: true,
          color: 'FFFFFF',
          fontFace: 'Calibri',
          align: 'left',
          valign: 'middle',
          margin: 0,
        })
      })

      slide.addShape('rect' as any, {
        x: 0,
        y: 6.95,
        w: 13.33,
        h: 0.55,
        fill: { color: '00C851' },
        line: { color: '00C851' },
      })

      slide.addText('Kabalikat para sa Maunlad na Buhay, Inc. (A Microfinance NGO)  🌿', {
        x: 0.3,
        y: 6.95,
        w: 12.73,
        h: 0.55,
        fontSize: 12,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Calibri',
        align: 'center',
        valign: 'middle',
        margin: 0,
      })
    }

    setGenMsg('Downloading…')

    await pptx.writeFile({
      fileName: 'groufie-photos.pptx',
    })

    setGenMsg(`✓ Downloaded groufie-photos.pptx (${entries.length + 1} slides)`)
    setGenerating(false)
  }

  const canSave = !!file && !!group.trim() && !saving

  const readyGroups = entries
    .map((e) => e.group)
    .filter((group, index, arr) => arr.indexOf(group) === index)
    .join(', ')

  return (
    <div style={{ maxWidth: 720, margin: '0 auto', padding: '2rem 1rem' }}>
      <div style={{ marginBottom: '2rem' }}>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: '#1d3461', letterSpacing: '-0.3px' }}>
          📸 Group Photo Uploader
        </h1>
        <p style={{ color: '#6b6a67', fontSize: 14, marginTop: 4 }}>
          Upload group photos, view all entries, and export to PowerPoint.
        </p>
      </div>

      {/* Tabs */}
      <div style={{ display: 'flex', borderBottom: '1.5px solid #e5e5e3', marginBottom: '1.5rem' }}>
        {(['upload', 'gallery', 'generate'] as Tab[]).map((t) => (
          <button
            key={t}
            onClick={() => {
              setTab(t)
              if (t === 'gallery') fetchEntries()
            }}
            style={{
              padding: '9px 20px',
              fontSize: 14,
              fontWeight: tab === t ? 600 : 400,
              color: tab === t ? '#1d3461' : '#6b6a67',
              background: 'none',
              border: 'none',
              borderBottom: tab === t ? '2.5px solid #1d3461' : '2.5px solid transparent',
              cursor: 'pointer',
              marginBottom: -1.5,
              transition: 'all 0.15s',
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
            onDragOver={(e) => {
              e.preventDefault()
              setDragOver(true)
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault()
              setDragOver(false)

              const f = e.dataTransfer.files[0]

              if (f?.type.startsWith('image/')) handleFile(f)
            }}
            style={{
              border: `2px dashed ${dragOver ? '#1d3461' : '#d1d0cd'}`,
              borderRadius: R,
              padding: '2.5rem',
              textAlign: 'center',
              cursor: 'pointer',
              background: dragOver ? '#e8eef6' : '#fafaf9',
              transition: 'all 0.15s',
            }}
          >
            <div style={{ fontSize: 40, marginBottom: 8 }}>📁</div>
            <p style={{ fontWeight: 600, fontSize: 15 }}>Click to upload or drag &amp; drop</p>
            <p style={{ color: '#6b6a67', fontSize: 13, marginTop: 4 }}>PNG, JPG, WEBP supported</p>
          </div>

          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            style={{ display: 'none' }}
            onChange={(e) => {
              const f = e.target.files?.[0]
              if (f) handleFile(f)
            }}
          />

          {preview && (
            <div style={{ position: 'relative', borderRadius: R, overflow: 'hidden', border: '1px solid #e5e5e3' }}>
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img
                src={preview}
                alt="Preview"
                style={{
                  width: '100%',
                  maxHeight: 260,
                  objectFit: 'contain',
                  display: 'block',
                  background: '#f0f0ef',
                }}
              />

              <button
                onClick={clearFile}
                style={{
                  position: 'absolute',
                  top: 8,
                  right: 8,
                  background: 'rgba(0,0,0,0.6)',
                  border: 'none',
                  borderRadius: '50%',
                  width: 30,
                  height: 30,
                  cursor: 'pointer',
                  color: '#fff',
                  fontSize: 16,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                }}
              >
                ✕
              </button>
            </div>
          )}

          <div style={{ background: '#fff', border: '1px solid #e5e5e3', borderRadius: R, padding: '1rem' }}>
            <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#6b6a67', marginBottom: 6 }}>
              GROUP NUMBER / IDENTIFIER
            </label>

            <input
              type="text"
              value={group}
              onChange={(e) => setGroup(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && handleSave()}
              placeholder="e.g. Group 3, Section A, Team Alpha"
              style={{
                width: '100%',
                padding: '10px 12px',
                border: '1px solid #e5e5e3',
                borderRadius: RS,
                fontSize: 15,
                outline: 'none',
                fontFamily: 'inherit',
              }}
            />
          </div>

          <button
            onClick={handleSave}
            disabled={!canSave}
            style={{
              width: '100%',
              padding: '12px',
              fontSize: 15,
              fontWeight: 600,
              background: canSave ? '#1d3461' : '#e5e5e3',
              color: canSave ? '#fff' : '#999',
              border: 'none',
              borderRadius: R,
              cursor: canSave ? 'pointer' : 'not-allowed',
              transition: 'all 0.15s',
            }}
          >
            {saving ? 'Saving…' : '💾 Save to database'}
          </button>

          {saveMsg && (
            <p
              style={{
                textAlign: 'center',
                color: saveMsg.includes('failed') ? '#dc2626' : '#16a34a',
                fontSize: 14,
                fontWeight: 500,
              }}
            >
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
              <p>No entries yet. Upload some group photos!</p>
            </div>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', gap: 14 }}>
              {entries.map((entry) => (
                <div
                  key={entry.id}
                  style={{
                    background: '#fff',
                    border: '1px solid #e5e5e3',
                    borderRadius: R,
                    overflow: 'hidden',
                    position: 'relative',
                  }}
                >
                  {/* eslint-disable-next-line @next/next/no-img-element */}
                  <img
                    src={entry.imageUrl}
                    alt={entry.group}
                    style={{
                      width: '100%',
                      height: 130,
                      objectFit: 'cover',
                      display: 'block',
                    }}
                  />

                  <button
                    onClick={() => handleDelete(entry)}
                    style={{
                      position: 'absolute',
                      top: 6,
                      right: 6,
                      background: 'rgba(0,0,0,0.55)',
                      border: 'none',
                      borderRadius: '50%',
                      width: 26,
                      height: 26,
                      cursor: 'pointer',
                      color: '#fff',
                      fontSize: 13,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                    }}
                    title="Delete"
                  >
                    ✕
                  </button>

                  <div style={{ padding: '8px 10px' }}>
                    <span
                      style={{
                        display: 'inline-block',
                        background: '#e8eef6',
                        color: '#1d3461',
                        fontSize: 12,
                        fontWeight: 600,
                        padding: '2px 8px',
                        borderRadius: 4,
                      }}
                    >
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
          <div style={{ background: '#fff', border: '1px solid #e5e5e3', borderRadius: R, padding: '1.25rem' }}>
            <p style={{ fontSize: 14, color: '#6b6a67', marginBottom: 12 }}>
              Compile all saved group photos into a PowerPoint file — one slide per entry with the group label and image.
            </p>

            <div style={{ background: '#f5f5f4', borderRadius: RS, padding: '10px 14px', fontSize: 13, color: '#6b6a67' }}>
              {entries.length === 0
                ? 'ℹ No entries saved yet.'
                : `📋 ${entries.length} entr${entries.length === 1 ? 'y' : 'ies'} ready: ${readyGroups}`}
            </div>
          </div>

          <button
            onClick={handleGenerate}
            disabled={entries.length === 0 || generating}
            style={{
              width: '100%',
              padding: '13px',
              fontSize: 15,
              fontWeight: 600,
              background: entries.length > 0 && !generating ? '#1d3461' : '#e5e5e3',
              color: entries.length > 0 && !generating ? '#fff' : '#999',
              border: 'none',
              borderRadius: R,
              cursor: entries.length > 0 && !generating ? 'pointer' : 'not-allowed',
              transition: 'all 0.15s',
            }}
          >
            {generating ? '⏳ Building…' : '📥 Generate PowerPoint'}
          </button>

          {genMsg && (
            <p
              style={{
                textAlign: 'center',
                fontSize: 13,
                color: genMsg.startsWith('✓') ? '#16a34a' : '#6b6a67',
                fontWeight: genMsg.startsWith('✓') ? 600 : 400,
              }}
            >
              {genMsg}
            </p>
          )}
        </div>
      )}
    </div>
  )
}
