import React from "react";
import { AbsoluteFill, Sequence, useVideoConfig, useCurrentFrame, interpolate } from "remotion";
import { Scene1 } from "./scenes/Scene1";
import { Scene2 } from "./scenes/Scene2";
import { Scene3 } from "./scenes/Scene3";
import { Scene4 } from "./scenes/Scene4";

const SCENE_DURATION = 60; // 2 seconds at 30fps
const OVERLAP = 10;        // frames of crossfade overlap

const FadeWrapper: React.FC<{ children: React.ReactNode; durationInFrames: number }> = ({
  children,
  durationInFrames,
}) => {
  const frame = useCurrentFrame();
  const fadeIn = interpolate(frame, [0, 10], [0, 1], { extrapolateRight: "clamp" });
  const fadeOut = interpolate(frame, [durationInFrames - 10, durationInFrames], [1, 0], {
    extrapolateLeft: "clamp",
  });
  const opacity = Math.min(fadeIn, fadeOut);

  return (
    <AbsoluteFill style={{ opacity }}>
      {children}
    </AbsoluteFill>
  );
};

export const Intro: React.FC = () => {
  return (
    <AbsoluteFill>
      <Sequence from={0} durationInFrames={SCENE_DURATION + OVERLAP}>
        <FadeWrapper durationInFrames={SCENE_DURATION + OVERLAP}>
          <Scene1 />
        </FadeWrapper>
      </Sequence>

      <Sequence from={SCENE_DURATION} durationInFrames={SCENE_DURATION + OVERLAP}>
        <FadeWrapper durationInFrames={SCENE_DURATION + OVERLAP}>
          <Scene2 />
        </FadeWrapper>
      </Sequence>

      <Sequence from={SCENE_DURATION * 2} durationInFrames={SCENE_DURATION + OVERLAP}>
        <FadeWrapper durationInFrames={SCENE_DURATION + OVERLAP}>
          <Scene3 />
        </FadeWrapper>
      </Sequence>

      <Sequence from={SCENE_DURATION * 3} durationInFrames={SCENE_DURATION + OVERLAP}>
        <FadeWrapper durationInFrames={SCENE_DURATION + OVERLAP}>
          <Scene4 />
        </FadeWrapper>
      </Sequence>
    </AbsoluteFill>
  );
};
