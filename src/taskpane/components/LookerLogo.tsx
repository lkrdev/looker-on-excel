import React, { useId } from "react";
import { lookerColors } from "../theme/lookerTheme";

export interface LookerIconProps {
  size?: number;
  variant?: "color" | "blue" | "monochrome";
  monochromeColor?: string;
  className?: string;
  style?: React.CSSProperties;
}

/**
 * Official Looker Logo Icon
 * Vector graphic accurately reproduced from official Looker assets.
 */
export const LookerIcon: React.FC<LookerIconProps> = ({
  size = 24,
  variant = "color",
  monochromeColor = "#1D5288",
  className,
  style,
}) => {
  const rawId = useId();
  const clipId = `looker-clip-${rawId.replace(/[^a-zA-Z0-9_-]/g, "")}`;

  // 1. Full 4-color Google/Looker logo (looker-32-color.svg)
  if (variant === "color") {
    return (
      <svg
        viewBox="0 0 32 32"
        width={size}
        height={size}
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        className={className}
        style={{
          display: "inline-block",
          verticalAlign: "middle",
          flexShrink: 0,
          ...style,
        }}
        role="img"
        aria-label="Looker"
      >
        <defs>
          <clipPath id={clipId}>
            <path
              d="M18.5,29c-3.584,0-6.5-2.916-6.5-6.5s2.916-6.5,6.5-6.5,6.5,2.916,6.5,6.5-2.916,6.5-6.5,6.5ZM18.5,18c-2.481,0-4.5,2.019-4.5,4.5s2.019,4.5,4.5,4.5,4.5-2.019,4.5-4.5-2.019-4.5-4.5-4.5Z"
              fill="none"
            />
          </clipPath>
        </defs>
        {/* White backing line */}
        <line x1="12.381" y1="11.854" x2="16.619" y2="8.146" stroke="#ffffff" strokeWidth="1" />
        {/* Red upper connector */}
        <path
          d="M12.381,12.854c-.278,0-.555-.115-.753-.342-.364-.415-.322-1.047.094-1.411l4.238-3.708c.416-.365,1.047-.321,1.411.094.364.416.322,1.047-.094,1.411l-4.238,3.709c-.19.166-.425.247-.658.247Z"
          fill={lookerColors.red}
        />
        {/* Blue main circle */}
        <path
          d="M18.5,29c-3.584,0-6.5-2.916-6.5-6.5s2.916-6.5,6.5-6.5,6.5,2.916,6.5,6.5-2.916,6.5-6.5,6.5ZM18.5,18c-2.481,0-4.5,2.019-4.5,4.5s2.019,4.5,4.5,4.5,4.5-2.019,4.5-4.5-2.019-4.5-4.5-4.5Z"
          fill={lookerColors.blue}
        />
        {/* Red upper node */}
        <path
          d="M18.5,10c-.935,0-1.814-.364-2.475-1.025s-1.025-1.54-1.025-2.475.364-1.814,1.025-2.475c1.321-1.321,3.627-1.322,4.95,0,.661.661,1.025,1.54,1.025,2.475s-.364,1.813-1.025,2.475h0c-.661.661-1.54,1.025-2.475,1.025ZM18.5,5c-.4,0-.777.156-1.06.439-.284.283-.44.66-.44,1.061s.156.777.439,1.06c.566.567,1.555.567,2.121,0t0,0c.283-.283.439-.66.439-1.06s-.156-.777-.439-1.06c-.283-.284-.66-.44-1.061-.44Z"
          fill={lookerColors.red}
        />
        {/* Green arc & diagonal connectors */}
        <path
          d="M25,22.5h-2c0-2.481-2.019-4.5-4.5-4.5-1.104,0-2.166.403-2.989,1.137l-1.329-1.494c1.19-1.06,2.724-1.643,4.318-1.643,3.584,0,6.5,2.916,6.5,6.5Z"
          fill={lookerColors.green}
        />
        <rect
          x="12.574"
          y="14.901"
          width="2"
          height="4.115"
          transform="translate(-7.838 13.303) rotate(-41.637)"
          fill={lookerColors.green}
        />
        <g clipPath={`url(#${clipId})`}>
          <rect
            x="14.354"
            y="14.225"
            width="2"
            height="9.472"
            transform="translate(-8.719 14.99) rotate(-41.636)"
            fill={lookerColors.green}
          />
        </g>
        {/* Yellow left node */}
        <path
          d="M10.5,17c-.935,0-1.814-.364-2.475-1.025-1.364-1.364-1.364-3.585,0-4.949,1.321-1.322,3.627-1.322,4.95,0,1.364,1.364,1.364,3.585,0,4.949h0c-.661.661-1.54,1.025-2.475,1.025ZM10.5,12c-.4,0-.777.156-1.06.439-.585.585-.585,1.536,0,2.121.566.566,1.555.566,2.121,0h0c.584-.585.584-1.536,0-2.121-.283-.283-.66-.439-1.061-.439Z"
          fill={lookerColors.yellow}
        />
      </svg>
    );
  }

  // 2. Monochrome or Blue Looker Variant (looker.svg specs)
  const cTop = variant === "blue" ? "#D2E3FC" : monochromeColor;
  const cMid = variant === "blue" ? "#5E97F6" : monochromeColor;
  const cBase = variant === "blue" ? "#4285F4" : monochromeColor;

  return (
    <svg
      viewBox="0 0 120 120"
      width={size}
      height={size}
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      className={className}
      style={{
        display: "inline-block",
        verticalAlign: "middle",
        flexShrink: 0,
        ...style,
      }}
      role="img"
      aria-label="Looker"
    >
      <path
        d="M72.4,22.1c-4.9,0-8.8,4-8.8,8.8c0,1.7,0.5,3.5,1.5,4.9l3.8-3.7c-0.1-0.4-0.2-0.8-0.2-1.1c0-2.1,1.7-3.8,3.7-3.8c2.1,0,3.8,1.7,3.8,3.7s-1.7,3.8-3.7,3.8c0,0,0,0-0.1,0c-0.4,0-0.8-0.1-1.1-0.2l-3.7,3.7c1.4,1,3.1,1.5,4.9,1.5c4.9,0,8.8-3.9,8.8-8.8C81.2,26.1,77.2,22.1,72.4,22.1z"
        fill={cTop}
      />
      <path
        d="M68.9,47.8c0-3-1-5.9-2.8-8.3l-4.7,5c0.6,1,0.9,2.2,0.9,3.4c-0.1,1.9-0.9,3.7-2.3,5l2.7,6.5C66.5,56.8,68.9,52.4,68.9,47.8z"
        fill={cMid}
      />
      <path
        d="M55.2,54.8c-3.9,0-7-3.1-7-7s3.1-7,7-7c1.4,0,2.7,0.4,3.8,1.1l4.8-4.8c-2.4-2-5.5-3.1-8.7-3.1c-7.6,0-13.8,6.2-13.8,13.8s6.2,13.8,13.8,13.8c0.9,0,1.9-0.1,2.8-0.3L55.2,54.8z"
        fill={cMid}
      />
      <path
        d="M72.6,60.4c-3.1,0-6.1,0.4-9,1.3l3.9,9.5c11.3-2.8,22.8,4,25.6,15.3c2.8,11.3-4,22.8-15.3,25.6c-11.3,2.8-22.8-4-25.6-15.3c-2.4-9.4,2-19.2,10.6-23.8l-3.9-9.4c-15.5,7.6-22,26.3-14.4,41.8s26.3,22,41.8,14.4s22-26.3,14.4-41.8C95.4,67.2,84.5,60.4,72.6,60.4z"
        fill={cBase}
      />
    </svg>
  );
};

/**
 * Signature 4-Color Accent Line
 */
export const LookerAccentBar: React.FC<{
  height?: number;
  mode?: "gradient" | "stripe";
  style?: React.CSSProperties;
}> = ({ height = 3, mode = "gradient", style }) => {
  return (
    <div
      style={{
        height: `${height}px`,
        width: "100%",
        background: mode === "gradient" ? lookerColors.accentGradient : lookerColors.accentStripe,
        flexShrink: 0,
        ...style,
      }}
    />
  );
};

export interface LookerBrandLockupProps {
  iconSize?: number;
  layout?: "horizontal" | "vertical";
  showSubtitle?: boolean;
  subtitleText?: string;
  style?: React.CSSProperties;
}

/**
 * Full Brand Lockup with Looker typography and Excel tag
 */
export const LookerBrandLockup: React.FC<LookerBrandLockupProps> = ({
  iconSize = 24,
  layout = "horizontal",
  showSubtitle = true,
  subtitleText = "for Microsoft Excel",
  style,
}) => {
  if (layout === "vertical") {
    return (
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          gap: "10px",
          textAlign: "center",
          ...style,
        }}
      >
        <div
          style={{
            padding: "12px",
            borderRadius: "50%",
            backgroundColor: "#F8FAFD",
            border: "1px solid #E8F0FE",
            boxShadow: "0 2px 8px rgba(66, 133, 244, 0.12)",
            display: "inline-flex",
          }}
        >
          <LookerIcon size={iconSize} variant="color" />
        </div>
        <div>
          <div
            style={{
              fontFamily: "'Google Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
              fontSize: "18px",
              fontWeight: 600,
              color: "#202124",
              letterSpacing: "-0.2px",
            }}
          >
            Looker
          </div>
          {showSubtitle && (
            <div
              style={{
                fontSize: "12px",
                color: lookerColors.navy,
                fontWeight: 500,
                marginTop: "2px",
              }}
            >
              {subtitleText}
            </div>
          )}
        </div>
      </div>
    );
  }

  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: "10px",
        ...style,
      }}
    >
      <LookerIcon size={iconSize} variant="color" />
      <div style={{ display: "flex", alignItems: "baseline", gap: "6px" }}>
        <span
          style={{
            fontFamily: "'Google Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
            fontSize: "16px",
            fontWeight: 600,
            color: "#202124",
            letterSpacing: "-0.2px",
          }}
        >
          Looker
        </span>
        {showSubtitle && (
          <span
            style={{
              fontSize: "12px",
              fontWeight: 500,
              color: lookerColors.navy,
              backgroundColor: lookerColors.blueLight,
              padding: "1px 6px",
              borderRadius: "4px",
            }}
          >
            Excel
          </span>
        )}
      </div>
    </div>
  );
};
