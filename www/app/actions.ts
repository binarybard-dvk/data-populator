'use server'

import { createClient } from '@/lib/supabase'
import { Provider } from '@supabase/supabase-js'
import { headers } from 'next/headers'
import { redirect } from 'next/navigation'

export async function emailLogin(email: string) {
	console.log(email)
	const supabase = createClient()
	const origin = headers().get('origin')

	const { data, error } = await supabase.auth.signInWithOtp({
		email,
		options: {
			shouldCreateUser: true,
			emailRedirectTo: `${origin}`,
		},
	})

	if (error) {
		console.log(error)
		throw new Error(error.message || 'Error sending magic link')
	} else {
		console.log(data)
		redirect(`/magic-link?email=${email}`)
	}
}

export async function oAuthLogin(provider: Provider) {
	const supabase = createClient()
	const origin = headers().get('origin')
	const { data, error } = await supabase.auth.signInWithOAuth({
		provider,
		options: {
			redirectTo: `${origin}`,
		},
	})

	if (error) {
		console.log(error)
		throw new Error(error.message || 'Error logging in')
	} else {
		return redirect(data.url)
	}
}
