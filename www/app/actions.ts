"use server";

import { createClient } from "@/lib/supabase";
import { Provider } from "@supabase/supabase-js";
import { headers } from "next/headers";
import { redirect } from "next/navigation";

export async function emailLogin(email: string) {
  console.log(email);
  const supabase = await createClient();
  const origin = headers().get("origin");

  const { data, error } = await supabase.auth.signInWithOtp({
    email,
    options: {
      emailRedirectTo: `${origin}/app`,
    },
  });

  if (error) {
    console.log(error);
    redirect(`/error?error=${error.message}`);
  } else {
    console.log(data);
    redirect(`/magic-link?email=${email}`);
  }
}

export async function oAuthLogin(provider: Provider) {
  const supabase = createClient();
  const origin = headers().get("origin");
  const { data, error } = await supabase.auth.signInWithOAuth({
    provider,
    options: {
      redirectTo: `${origin}/app`,
    },
  });

  if (error) {
    console.log(error);
    redirect(`/error?error=${error.message}`);
  } else {
    return redirect(data.url);
  }
}
