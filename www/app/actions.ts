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
      emailRedirectTo: `${origin}/auth/callback`,
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
      redirectTo: `${origin}/auth/callback`,
    },
  });

  if (error) {
    console.log(error);
    redirect(`/error?error=${error.message}`);
  } else {
    return redirect(data.url);
  }
}

export async function getUserData() {
  const supabase = createClient();
  const { data, error } = await supabase.auth.getUser();
  const {
    data: profile,
    error: profileError,
    status,
  } = await supabase
    .from("profiles")
    .select(`username, website, avatar_url`)
    .eq("id", data?.user?.id)
    .single();

  if (error || (profileError && status !== 406)) {
    console.log(error);
    redirect(`/error?error=${error?.message}`);
  } else {
    return { ...profile, user_id: data?.user?.id, email: data?.user?.email };
  }
}

export async function updateProfile({
  username,
  website,
  avatar_url,
  user_id,
}: {
  username: string;
  website: string;
  avatar_url: string;
  user_id: string;
}) {
  const supabase = createClient();
  const updates = {
    id: user_id,
    username,
    website,
    avatar_url,
    updated_at: new Date(),
  };

  const { error } = await supabase.from("profiles").upsert(updates);

  if (error) {
    throw error;
  }
}
