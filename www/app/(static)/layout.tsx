export default function StaticLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return <main className="max-w-2xl mx-auto px-4">{children}</main>;
}
