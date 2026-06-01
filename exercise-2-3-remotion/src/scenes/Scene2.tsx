import React from "react";
import { useCurrentFrame, useVideoConfig, spring, interpolate } from "remotion";

const COUNTRIES = ["MX", "BR", "CO", "AR", "CL"];

// Scene 2 — Latin America role (portrait 1080×1920)
export const Scene2: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();

  const headingP = spring({ frame,        fps, config: { damping: 80 } });
  const titleP   = spring({ frame: frame - 10, fps, config: { damping: 80 } });
  const descP    = spring({ frame: frame - 20, fps, config: { damping: 80 } });

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        background: "#00A1E0",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'Helvetica Neue', Arial, sans-serif",
        padding: "0 80px",
        overflow: "hidden",
        gap: 0,
      }}
    >
      {/* Label */}
      <div
        style={{
          color: "#032D60",
          fontSize: 28,
          fontWeight: 700,
          letterSpacing: 4,
          textTransform: "uppercase",
          marginBottom: 28,
          opacity: headingP,
        }}
      >
        Technical Sales
      </div>

      {/* Big title */}
      <div
        style={{
          color: "#fff",
          fontSize: 96,
          fontWeight: 800,
          lineHeight: 1.0,
          textAlign: "center",
          opacity: titleP,
          transform: `translateY(${interpolate(titleP, [0, 1], [50, 0])}px)`,
        }}
      >
        Leading<br />
        <span style={{ color: "#032D60" }}>Latin<br />America</span>
      </div>

      {/* Description */}
      <div
        style={{
          color: "#032D60",
          fontSize: 30,
          lineHeight: 1.5,
          textAlign: "center",
          marginTop: 40,
          marginBottom: 64,
          opacity: descP,
          transform: `translateY(${interpolate(descP, [0, 1], [20, 0])}px)`,
        }}
      >
        Connecting enterprise technology<br />to business outcomes
      </div>

      {/* Country tags — horizontal row */}
      <div
        style={{
          display: "flex",
          flexDirection: "row",
          flexWrap: "wrap",
          gap: 18,
          justifyContent: "center",
        }}
      >
        {COUNTRIES.map((country, i) => {
          const p = spring({ frame: frame - 20 - i * 5, fps, config: { damping: 80 } });
          return (
            <div
              key={country}
              style={{
                background: "rgba(255,255,255,0.25)",
                border: "2px solid rgba(255,255,255,0.6)",
                borderRadius: 16,
                padding: "14px 36px",
                color: "#fff",
                fontSize: 28,
                fontWeight: 700,
                letterSpacing: 2,
                opacity: p,
                transform: `scale(${interpolate(p, [0, 1], [0.7, 1])})`,
              }}
            >
              {country}
            </div>
          );
        })}
      </div>
    </div>
  );
};
