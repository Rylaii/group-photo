import { list } from '@vercel/blob'
import { NextResponse } from 'next/server'

export const revalidate = 0

export async function GET() {
  try {
    const { blobs } = await list({ prefix: 'meta/' })

    const entries = await Promise.all(
      blobs.map(async (blob) => {
        try {
          const res = await fetch(blob.url)
          return await res.json()
        } catch {
          return null
        }
      })
    )

    const valid = entries
      .filter(Boolean)
      .sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())

    return NextResponse.json(valid)
  } catch (err) {
    console.error(err)
    return NextResponse.json([], { status: 500 })
  }
}
