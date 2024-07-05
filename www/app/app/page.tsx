"use client";

import { Separator } from "@/components/ui/separator";
import { useQuery } from "@tanstack/react-query";
import { getUserData } from "../actions";

export default function HomePage() {
  const { data, error, isLoading } = useQuery({
    queryKey: ["user"],
    queryFn: getUserData,
  });

  if (isLoading) {
    return <div>Loading...</div>;
  }

  return (
    <div className="flex flex-col space-y-4">
      <div className="space-y-0.5">
        <h2 className="text-2xl font-bold tracking-tight">
          Hi {data?.username || "User"}!
        </h2>
        <p className="text-muted-foreground">
          Here are all the actions you can do with this app.
        </p>
      </div>
      <Separator className="my-6" />
    </div>
  );
}
