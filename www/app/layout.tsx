import type { Metadata } from 'next'
import { Inter } from 'next/font/google'
import { Toaster } from '@/components/ui/sonner'
import './globals.css'

const inter = Inter({ subsets: ['latin'] })

export const metadata: Metadata = {
	title: 'DataCraft',
	description:
		'Merge Excel templates with JSON data. Fast, accurate, and designed to streamline your workflow.',
}

export default function RootLayout({
	children,
}: Readonly<{
	children: React.ReactNode
}>) {
	return (
		<html lang='en'>
			<body className={`${inter.className}`}>
				<main>{children}</main>
				<Toaster />
			</body>
		</html>
	)
}
