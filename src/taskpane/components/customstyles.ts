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

export function mapApiStyleToCustomStyle(apiStyle: any): CustomTextStyle {
  const props = apiStyle.Properties || {};
  return {
    id: apiStyle.Name,
    name: apiStyle.DisplayName,
    properties: {
      bold: props.Bold ?? false,
      italic: props.Italic ?? false,
      underline: props.Underline ?? false,
      fontFamily: props.FontFamily ?? "Calibri",
      size: props.Size ?? "11pt",
      fontColor: props.FontColor ?? props.fontColor ?? "#000000",
      backgroundColor: props.BackgroundColor ?? props.backgroundColor ?? "transparent"
    }
  };
}
