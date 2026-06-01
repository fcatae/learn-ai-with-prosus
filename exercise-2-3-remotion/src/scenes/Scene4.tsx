import React from "react";
import { useCurrentFrame, useVideoConfig, spring, interpolate } from "remotion";

// Scene 4 — The bridge (portrait 1080×1920)
// Layout: Sales box → vertical bar → Technology box → tagline
export const Scene4: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();

  const topP     = spring({ frame,        fps, config: { damping: 70 } });
  const barP     = spring({ frame: frame - 10, fps, config: { damping: 60 } });
  const bottomP  = spring({ frame: frame - 18, fps, config: { damping: 70 } });
  const taglineP = spring({ frame: frame - 28, fps, config: { damping: 80 } });

  const barHeight = interpolate(barP, [0, 1], [0, 120]);

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        background: "#fff",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'Helvetica Neue', Arial, sans-serif",
        padding: "0 80px",
        gap: 0,
        overflow: "hidden",
      }}
    >
      {/* Sales box */}
      <div
        style={{
          background: "#032D60",
          borderRadius: 24,
          padding: "36px 80px",
          color: "#fff",
          fontSize: 52,
          fontWeight: 800,
          opacity: topP,
          transform: `translateY(${interpolate(topP, [0, 1], [-60, 0])}px)`,
        }}
      >
        Sales
      </div>

      {/* Vertical bridge bar */}
      <div
        style={{
          width: 10,
          height: barHeight,
          background: "linear-gradient(180deg, #032D60, #00A1E0)",
          borderRadius: 5,
        }}
      />

      {/* Technology box */}
      <div
        style={{
          background: "#00A1E0",
          borderRadius: 24,
          padding: "36px 60px",
          color: "#fff",
          fontSize: 52,
          fontWeight: 800,
          opacity: bottomP,
          transform: `translateY(${interpolate(bottomP, [0, 1], [60, 0])}px)`,
        }}
      >
        Technology
      </div>

      {/* Tagline */}
      <div
        style={{
          marginTop: 80,
          opacity: taglineP,
          transform: `translateY(${interpolate(taglineP, [0, 1], [40, 0])}px)`,
          textAlign: "center",
        }}
      >
        <div
          style={{
            color: "#032D60",
            fontSize: 52,
            fontWeight: 800,
            lineHeight: 1.15,
          }}
        >
          The bridge between<br />
          <span style={{ color: "#00A1E0" }}>commercial results</span><br />
          and technical excellence
        </div>
      </div>
    </div>
  );
};
