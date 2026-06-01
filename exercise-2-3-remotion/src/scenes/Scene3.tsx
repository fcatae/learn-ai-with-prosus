import React from "react";
import { useCurrentFrame, useVideoConfig, spring, interpolate } from "remotion";

const ITEMS = [
  { icon: "🎯", label: "Building CRM Discipline",          delay: 8  },
  { icon: "⚙️", label: "Driving Sales Process Adoption",   delay: 16 },
  { icon: "📈", label: "Growing Presales Team Capability", delay: 24 },
];

const Item: React.FC<{ icon: string; label: string; delay: number }> = ({ icon, label, delay }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const p = spring({ frame: frame - delay, fps, config: { damping: 70 } });

  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: 32,
        opacity: p,
        transform: `translateX(${interpolate(p, [0, 1], [-60, 0])}px)`,
        background: "rgba(255,255,255,0.08)",
        borderRadius: 24,
        padding: "32px 48px",
        border: "1px solid rgba(255,255,255,0.15)",
        width: "100%",
      }}
    >
      <span style={{ fontSize: 52, lineHeight: 1 }}>{icon}</span>
      <span style={{ color: "#fff", fontSize: 34, fontWeight: 600, lineHeight: 1.2 }}>
        {label}
      </span>
    </div>
  );
};

// Scene 3 — Mission (portrait 1080×1920)
export const Scene3: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const headerP = spring({ frame, fps, config: { damping: 80 } });

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        background: "linear-gradient(180deg, #0B4F8A 0%, #032D60 100%)",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'Helvetica Neue', Arial, sans-serif",
        padding: "0 80px",
        gap: 28,
        overflow: "hidden",
      }}
    >
      {/* Section label */}
      <div
        style={{
          color: "#00A1E0",
          fontSize: 28,
          fontWeight: 700,
          letterSpacing: 4,
          textTransform: "uppercase",
          marginBottom: 16,
          opacity: headerP,
        }}
      >
        Mission
      </div>

      {ITEMS.map((item) => (
        <Item key={item.label} {...item} />
      ))}
    </div>
  );
};
