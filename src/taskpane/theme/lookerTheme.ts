import { Theme, webLightTheme } from "@fluentui/react-components";

/**
 * Official Looker / Google Cloud BI Brand Colors
 */
export const lookerColors = {
  // Core Looker & Google Palette
  blue: "#4285F4",
  blueDark: "#1A73E8",
  blueDeep: "#1557B0",
  navy: "#1D5288",
  navyDark: "#143C64",
  red: "#EA4335",
  yellow: "#FBBC04",
  yellowDark: "#F29900",
  green: "#34A853",
  greenDark: "#1E8E3E",

  // Background tints for badges & categories
  blueLight: "#E8F0FE",
  greenLight: "#E6F4EA",
  yellowLight: "#FEF7E0",
  redLight: "#FCE8E6",

  // Gradients
  accentGradient:
    "linear-gradient(90deg, #4285F4 0%, #EA4335 33%, #FBBC04 66%, #34A853 100%)",
  accentStripe:
    "linear-gradient(90deg, #4285F4 0%, #4285F4 25%, #EA4335 25%, #EA4335 50%, #FBBC04 50%, #FBBC04 75%, #34A853 75%, #34A853 100%)",
};

/**
 * Custom Fluent UI v9 Theme configured with Looker's signature brand colors.
 */
export const lookerLightTheme: Theme = {
  ...webLightTheme,
  // Primary buttons and brand accents
  colorBrandBackground: lookerColors.blueDark,
  colorBrandBackgroundHover: lookerColors.blueDeep,
  colorBrandBackgroundPressed: "#0D47A1",
  colorBrandBackgroundSelected: lookerColors.blueDark,

  // Brand foreground / text
  colorBrandForeground1: lookerColors.blueDark,
  colorBrandForeground2: lookerColors.blueDeep,
  colorBrandForegroundLink: lookerColors.blueDark,
  colorBrandForegroundLinkHover: lookerColors.blueDeep,
  colorBrandForegroundLinkPressed: "#0D47A1",
  colorBrandForegroundLinkSelected: lookerColors.blueDark,

  // Strokes and outlines
  colorBrandStroke1: lookerColors.blueDark,
  colorBrandStroke2: lookerColors.blue,
  colorBrandStroke2Hover: lookerColors.blueDeep,
  colorBrandStroke2Pressed: "#0D47A1",

  // Focus and shadows
  colorCompoundBrandStroke: lookerColors.blueDark,
  colorCompoundBrandStrokeHover: lookerColors.blueDeep,
  colorCompoundBrandStrokePressed: "#0D47A1",
};
