import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Asaan Excel – Lead Data Processor",
  description: "Upload Excel lead sheets, apply macros, and download results",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
