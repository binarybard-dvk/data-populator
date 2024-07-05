import { SidebarNav } from "@/components/sidebar-nav";
import { Separator } from "@/components/ui/separator";
import { getUserData } from "../actions";
import {
  HydrationBoundary,
  QueryClient,
  dehydrate,
} from "@tanstack/react-query";
import { get } from "http";
import { getQueryClient } from "@/lib/get-query-client";

const sidebarNavItems = [
  {
    title: "Home",
    href: "/app",
  },
  {
    title: "Templates",
    href: "/app/templates",
  },
  {
    title: "API connections",
    href: "/app/api",
  },
  {
    title: "Account",
    href: "/app/account",
  },
];

interface AppLayoutProps {
  children: React.ReactNode;
}

export default async function AppLayout({ children }: AppLayoutProps) {
  const queryClient = getQueryClient();
  await queryClient.prefetchQuery({
    queryKey: ["user"],
    queryFn: getUserData,
  });

  return (
    <>
      <div className="space-y-6 p-10 pb-16 md:block">
        <div className="flex flex-col space-y-8 lg:flex-row lg:space-x-12 lg:space-y-0">
          <aside className="-mx-4 lg:w-1/5">
            <SidebarNav items={sidebarNavItems} />
          </aside>
          <div className="flex-1 lg:max-w-2xl">
            <HydrationBoundary state={dehydrate(queryClient)}>
              {children}
            </HydrationBoundary>
          </div>
        </div>
      </div>
    </>
  );
}
