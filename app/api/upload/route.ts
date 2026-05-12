import { put } from '@vercel/blob'
import { NextRequest, NextResponse } from 'next/server'

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()
    const file = formData.get('file') as File
    const group = formData.get('group') as string

    if (!file || !group) {
      return NextResponse.json({ error: 'Missing file or group' }, { status: 400 })
    }

    const id = crypto.randomUUID()
    const ext = file.name.split('.').pop() || 'jpg'
    const filename = `photos/${id}.${ext}`

    const blob = await put(filename, file, { access: 'public' })

    const meta = {
      id,
      group,
      imageUrl: blob.url,
      date: new Date().toISOString(),
    }
    await put(`meta/${id}.json`, JSON.stringify(meta), {
      access: 'public',
      contentType: 'application/json',
    })

    return NextResponse.json(meta)
  } catch (err) {
    console.error(err)
    return NextResponse.json({ error: 'Upload failed' }, { status: 500 })
  }
}
