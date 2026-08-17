import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "US Phone Cleaner",
  description: "Clean, validate, and enrich US phone number lead files",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className="antialiased">{children}</body>
    </html>
  );
}
