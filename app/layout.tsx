import './globals.css'
import type { Metadata } from 'next'

export const metadata: Metadata = {
  title: 'Group Photo Uploader',
  description: 'Upload group photos and generate a PowerPoint presentation',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}
