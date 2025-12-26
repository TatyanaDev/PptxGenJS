import { ReactNode } from "react";
import { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "PPTXGenJS",
  description: "A simple demo that generates a PowerPoint slide using pptxgenjs and lets you download it as a PPTX file",
};

export default function RootLayout({ children }: Readonly<{ children: ReactNode }>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
