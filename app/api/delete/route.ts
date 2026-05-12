import { del, list } from '@vercel/blob'
import { NextRequest, NextResponse } from 'next/server'

export async function DELETE(request: NextRequest) {
  try {
    const { id, imageUrl } = await request.json()
    if (!id) return NextResponse.json({ error: 'Missing id' }, { status: 400 })

    // Delete metadata blob
    const { blobs } = await list({ prefix: `meta/${id}` })
    for (const blob of blobs) await del(blob.url)

    // Delete image blob if URL provided
    if (imageUrl) {
      try { await del(imageUrl) } catch {}
    }

    return NextResponse.json({ ok: true })
  } catch (err) {
    console.error(err)
    return NextResponse.json({ error: 'Delete failed' }, { status: 500 })
  }
}
