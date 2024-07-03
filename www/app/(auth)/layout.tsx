export default function AuthLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <main className="max-w-2xl h-screen mx-auto flex flex-col">{children}</main>
  );
}
