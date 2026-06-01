import React from "react";
import { Composition } from "remotion";
import { Intro } from "./Intro";

export const RemotionRoot: React.FC = () => {
  return (
    <>
      <Composition
        id="Intro"
        component={Intro}
        durationInFrames={240}  // 4 scenes × 60 frames = 8 seconds at 30fps
        fps={30}
        width={1080}
        height={1920}
      />
    </>
  );
};
