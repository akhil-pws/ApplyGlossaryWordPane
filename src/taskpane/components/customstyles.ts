export interface CustomStyleProperties {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  fontFamily: string;
  size: string; // e.g., "11pt", "12pt", "14pt"
  fontColor: string; // Hex color code, e.g., "#1B365D"
  backgroundColor: string; // Hex color code or "transparent", e.g., "#F5F7FA"
}

export interface CustomTextStyle {
  id: string;
  name: string;
  properties: CustomStyleProperties;
}

// 6 default customized styles for the Proof of Concept (POC)
export const customizedStyles: CustomTextStyle[] = [
  {
    id: "style_corporate_navy",
    name: "Corporate Navy",
    properties: {
      bold: true,
      italic: false,
      underline: false,
      fontFamily: "Calibri",
      size: "11pt",
      fontColor: "#1B365D",
      backgroundColor: "#F0F4F8"
    }
  },
  {
    id: "style_warning_red",
    name: "Warning Highlight",
    properties: {
      bold: true,
      italic: false,
      underline: true,
      fontFamily: "Segoe UI",
      size: "10.5pt",
      fontColor: "#C62828",
      backgroundColor: "#FFEBEE"
    }
  },
  {
    id: "style_emerald_editorial",
    name: "Emerald Editorial",
    properties: {
      bold: false,
      italic: true,
      underline: false,
      fontFamily: "Georgia",
      size: "12pt",
      fontColor: "#1B5E20",
      backgroundColor: "#E8F5E9"
    }
  },
  {
    id: "style_sunshine_accent",
    name: "Sunshine Highlight",
    properties: {
      bold: true,
      italic: false,
      underline: false,
      fontFamily: "Arial",
      size: "11pt",
      fontColor: "#E65100",
      backgroundColor: "#FFFDE7"
    }
  },
  {
    id: "style_royal_emphasis",
    name: "Royal Emphasis",
    properties: {
      bold: true,
      italic: true,
      underline: false,
      fontFamily: "Times New Roman",
      size: "36pt",
      fontColor: "#4A148C",
      backgroundColor: "#F3E5F5"
    }
  },
  {
    id: "style_modern_charcoal",
    name: "Modern Charcoal",
    properties: {
      bold: false,
      italic: false,
      underline: false,
      fontFamily: "Arial",
      size: "10pt",
      fontColor: "#333333",
      backgroundColor: "#ECEFF1"
    }
  }
];
