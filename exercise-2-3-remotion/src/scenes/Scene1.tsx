import React from "react";
import { useCurrentFrame, useVideoConfig, spring, interpolate } from "remotion";

// Scene 1 — Title card (portrait 1080×1920)
export const Scene1: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();

  const badgeP  = spring({ frame,        fps, config: { damping: 80 } });
  const titleP  = spring({ frame: frame - 8,  fps, config: { damping: 80 } });
  const subP    = spring({ frame: frame - 18, fps, config: { damping: 80 } });

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        background: "linear-gradient(180deg, #032D60 0%, #0B4F8A 60%, #032D60 100%)",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'Helvetica Neue', Arial, sans-serif",
        overflow: "hidden",
        position: "relative",
        padding: "0 80px",
      }}
    >
      {/* Grid background */}
      <div
        style={{
          position: "absolute",
          inset: 0,
          backgroundImage:
            "linear-gradient(rgba(0,161,224,0.06) 1px, transparent 1px)," +
            "linear-gradient(90deg, rgba(0,161,224,0.06) 1px, transparent 1px)",
          backgroundSize: "80px 80px",
        }}
      />

      {/* Salesforce badge */}
      <div
        style={{
          background: "#00A1E0",
          borderRadius: 20,
          padding: "16px 48px",
          marginBottom: 60,
          fontSize: 32,
          fontWeight: 700,
          color: "#fff",
          letterSpacing: 3,
          textTransform: "uppercase",
          opacity: badgeP,
          transform: `scale(${interpolate(badgeP, [0, 1], [0.7, 1])})`,
        }}
      >
        Salesforce
      </div>

      {/* Title */}
      <div
        style={{
          color: "#fff",
          fontSize: 88,
          fontWeight: 800,
          lineHeight: 1.05,
          textAlign: "center",
          opacity: titleP,
          transform: `translateY(${interpolate(titleP, [0, 1], [60, 0])}px)`,
        }}
      >
        Director of<br />
        <span style={{ color: "#00A1E0" }}>Solutioning<br />Engineering</span>
      </div>

      {/* Subtitle */}
      <div
        style={{
          marginTop: 48,
          color: "rgba(255,255,255,0.65)",
          fontSize: 36,
          fontWeight: 400,
          letterSpacing: 1,
          textAlign: "center",
          opacity: subP,
          transform: `translateY(${interpolate(subP, [0, 1], [30, 0])}px)`,
        }}
      >
        Technical Sales Leader<br />Latin America
      </div>
    </div>
  );
};
