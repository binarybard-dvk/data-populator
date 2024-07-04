"use client";

import { Button } from "@/components/ui/button";
import { useRouter, useSearchParams } from "next/navigation";
import React from "react";

export default function ErrorPage() {
  const params = useSearchParams();
  const router = useRouter();

  return (
    <div className="m-auto max-w-sm space-y-6">
      <div className="space-y-2 my-4 text-center border border-neutral-200 rounded-lg p-8">
        <h1 className="text-3xl font-bold">Something went wrong!</h1>
        <p className="text-muted-foreground">{params.get("error")}</p>
      </div>
      <Button
        className="w-full mt-4"
        variant="link"
        onClick={() => router.back()}
      >
        Go back
      </Button>
    </div>
  );
}
