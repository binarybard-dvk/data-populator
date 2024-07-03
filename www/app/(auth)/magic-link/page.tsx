'use client'

import { Button } from '@/components/ui/button'
import Link from 'next/link'
import { useSearchParams } from 'next/navigation'
import React from 'react'

export default function MagicLink() {
	const params = useSearchParams()
	return (
		<div className='m-auto max-w-sm space-y-6'>
			<span className='text-muted-foreground italic text-sm'>
				logging in as {params.get('email')}
			</span>
			<div className='space-y-2 my-4 text-center border border-neutral-200 rounded-lg p-8'>
				<h1 className='text-3xl font-bold'>Check your inbox</h1>
				<p className='text-muted-foreground'>
					We sent a magic link ✨ to your email. Click the link to login.
				</p>
			</div>
			<Link href='/login' className='w-full'>
				<Button className='w-full mt-4' variant='link'>
					Back to login
				</Button>
			</Link>
		</div>
	)
}
