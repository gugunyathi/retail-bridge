import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useEffect, useState } from "react";
import { ScaledSlide } from "@/components/ScaledSlide";
import { SlideIndexProvider } from "@/components/slide-kit";
import { slides } from "@/slides/deck";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "TM Pick n Pay Express — Diaspora-to-Door Pitch Deck" },
      {
        name: "description",
        content:
          "Executive pitch deck evolving TM Pick n Pay's Click & Collect platform into a diaspora-to-door delivery engine, with revenue models and commercial options.",
      },
      { property: "og:title", content: "TM Pick n Pay Express — Diaspora-to-Door Pitch Deck" },
      {
        property: "og:description",
        content:
          "Revenue projections, business model options and integration roadmap for TM Pick n Pay's cross-border delivery platform.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Deck,
});

function Deck() {
  const [i, setI] = useState(0);
  const [grid, setGrid] = useState(false);

  const go = useCallback(
    (n: number) => setI((c) => Math.min(slides.length - 1, Math.max(0, c + n))),
    [],
  );

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "ArrowRight" || e.key === " ") go(1);
      if (e.key === "ArrowLeft") go(-1);
      if (e.key.toLowerCase() === "g") setGrid((g) => !g);
      if (e.key === "F5") {
        e.preventDefault();
        document.documentElement.requestFullscreen?.();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [go]);

  const current = slides[i] ?? slides[0]!;

  useEffect(() => {
    document.title = `${i + 1}/${slides.length} — ${current.title}`;
  }, [i, current.title]);

  const Current = current.Component;


  return (
    <div className="flex min-h-screen flex-col bg-pnp-blue-deep">
      <header className="flex flex-wrap items-center gap-3 border-b border-white/10 px-6 py-4">
        <h1 className="mr-auto text-sm font-semibold tracking-tight text-white">
          TM Pick n Pay Express · Pitch Deck
        </h1>
        <button
          onClick={() => setGrid((g) => !g)}
          className="rounded-full border border-white/25 px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-white/10"
        >
          {grid ? "Slide view" : "Grid view"}
        </button>
        <a
          href="/TM-Pick-n-Pay-Express.pdf"
          download="TM-Pick-n-Pay-Express.pdf"
          className="rounded-full bg-white px-4 py-2 text-xs font-semibold text-pnp-blue-deep transition-opacity hover:opacity-90"
        >
          Download PDF
        </a>
        <a
          href="/TM-Pick-n-Pay-Express.pptx"
          download="TM-Pick-n-Pay-Express.pptx"
          className="rounded-full bg-pnp-red px-4 py-2 text-xs font-semibold text-white transition-opacity hover:opacity-90"
        >
          Download PPTX
        </a>
        <a
          href="/TM-Pick-n-Pay-Express.ppt"
          download="TM-Pick-n-Pay-Express.ppt"
          className="rounded-full bg-pnp-red px-4 py-2 text-xs font-semibold text-white transition-opacity hover:opacity-90"
        >
          Download PPT
        </a>


        <button
          onClick={() => document.documentElement.requestFullscreen?.()}
          className="rounded-full border border-white/25 px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-white/10"
        >
          Present
        </button>
      </header>

      {grid ? (
        <div className="grid flex-1 grid-cols-1 gap-6 overflow-auto p-6 sm:grid-cols-2 xl:grid-cols-3">
          {slides.map((s, idx) => {
            const C = s.Component;
            return (
              <button
                key={s.id}
                onClick={() => {
                  setI(idx);
                  setGrid(false);
                }}
                className="group text-left"
              >
                <div className="aspect-video overflow-hidden rounded-xl border border-white/15">
                  <ScaledSlide>
                    <SlideIndexProvider value={idx + 1}>
                      <C />
                    </SlideIndexProvider>
                  </ScaledSlide>
                </div>
                <div className="mt-2 text-xs font-medium text-white/70">
                  {idx + 1}. {s.title}
                </div>
              </button>
            );
          })}
        </div>
      ) : (
        <>
          <main className="flex-1 p-4">
            <div className="mx-auto aspect-video h-full max-h-[calc(100vh-160px)] w-full max-w-[1600px] overflow-hidden rounded-2xl">
              <ScaledSlide>
                <SlideIndexProvider value={i + 1}>
                  <Current />
                </SlideIndexProvider>
              </ScaledSlide>
            </div>
          </main>
          <footer className="flex items-center justify-center gap-4 pb-6 text-white">
            <button
              onClick={() => go(-1)}
              className="rounded-full border border-white/25 px-5 py-2 text-sm font-semibold transition-colors hover:bg-white/10"
            >
              ←
            </button>
            <span className="text-sm tabular-nums text-white/70">
              {i + 1} / {slides.length} · {current.title}
            </span>
            <button
              onClick={() => go(1)}
              className="rounded-full border border-white/25 px-5 py-2 text-sm font-semibold transition-colors hover:bg-white/10"
            >
              →
            </button>
          </footer>
        </>
      )}
    </div>
  );
}
