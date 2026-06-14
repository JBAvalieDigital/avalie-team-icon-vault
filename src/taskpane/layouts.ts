export type FieldType = "text" | "textarea" | "color";

export interface LayoutField {
  id: string;
  type: FieldType;
  label: string;
  default: string;
}

export interface LayoutTemplate {
  id: string;
  name: string;
  fields: LayoutField[];
  renderSvg: (data: Record<string, string>, iconSvg: string) => string;
}

// Helper to wrap text into multiple <tspan> elements for SVG
function wrapText(text: string, maxCharsPerLine: number, x: number, initialY: number, lineHeight: number = 24): string {
  const words = text.split(" ");
  const lines: string[] = [];
  let currentLine = "";

  words.forEach((word) => {
    if ((currentLine + word).length > maxCharsPerLine) {
      lines.push(currentLine.trim());
      currentLine = word + " ";
    } else {
      currentLine += word + " ";
    }
  });
  if (currentLine.trim()) {
    lines.push(currentLine.trim());
  }

  return lines
    .map((line, index) => {
      // For the first line we position it at initialY, subsequent lines use dy
      if (index === 0) {
        return `<tspan x="${x}" y="${initialY}">${escapeXml(line)}</tspan>`;
      }
      return `<tspan x="${x}" dy="${lineHeight}">${escapeXml(line)}</tspan>`;
    })
    .join("");
}

function escapeXml(unsafe: string): string {
  if (!unsafe) return "";
  return unsafe.replace(/[<>&'"]/g, (c) => {
    switch (c) {
      case "<": return "&lt;";
      case ">": return "&gt;";
      case "&": return "&amp;";
      case "'": return "&apos;";
      case '"': return "&quot;";
      default: return c;
    }
  });
}

// Helper to extract the inner content of an SVG (stripping the outer <svg> tag)
function getSvgInner(svgStr: string): string {
  if (!svgStr) return "";
  const match = svgStr.match(/<svg[^>]*>([\s\S]*?)<\/svg>/i);
  return match ? match[1] : svgStr;
}

export const layouts: LayoutTemplate[] = [
  {
    id: "banner-card",
    name: "Banner Card",
    fields: [
      { id: "title", type: "text", label: "Title", default: "Important Notice" },
      { id: "desc", type: "textarea", label: "Description", default: "This is a brief description of the notice. It can span a couple of lines." },
      { id: "color", type: "color", label: "Accent Color", default: "#6B21A8" }
    ],
    renderSvg: (data, iconSvg) => {
      const innerIcon = getSvgInner(iconSvg);
      return `
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 160" width="600" height="160">
          <rect x="0" y="0" width="600" height="160" rx="16" fill="#F3F4F6" />
          
          <!-- Icon Background & Icon -->
          <g transform="translate(32, 40)">
            <rect x="0" y="0" width="80" height="80" rx="16" fill="${data.color}" />
            <g fill="#ffffff" stroke="#ffffff" transform="translate(16, 16) scale(2)">
              ${innerIcon}
            </g>
          </g>
          
          <!-- Text -->
          <text x="144" font-family="system-ui, -apple-system, sans-serif" font-weight="bold" font-size="28" fill="${data.color}">
            ${wrapText(data.title || '', 40, 144, 70, 32)}
          </text>
          
          <text x="144" font-family="system-ui, -apple-system, sans-serif" font-weight="normal" font-size="20" fill="#374151">
            ${wrapText(data.desc || '', 45, 144, 110, 26)}
          </text>
        </svg>
      `.trim();
    }
  },
  {
    id: "step-card",
    name: "Step Card",
    fields: [
      { id: "stepNum", type: "text", label: "Step Number", default: "1" },
      { id: "title", type: "text", label: "Title", default: "File Migration" },
      { id: "color", type: "color", label: "Primary Color", default: "#45B08C" }
    ],
    renderSvg: (data, iconSvg) => {
      const innerIcon = getSvgInner(iconSvg);
      return `
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 220" width="320" height="220">
          <!-- Main Card -->
          <rect x="20" y="30" width="280" height="160" rx="24" fill="${data.color}" />
          
          <!-- Icon Box -->
          <g transform="translate(110, 50)">
            <rect x="0" y="0" width="100" height="80" rx="16" fill="transparent" stroke="#ffffff" stroke-width="4" />
            <g fill="#ffffff" stroke="#ffffff" transform="translate(26, 16) scale(2)">
              ${innerIcon}
            </g>
          </g>
          
          <!-- Title -->
          <text x="160" y="165" font-family="system-ui, -apple-system, sans-serif" font-weight="bold" font-size="24" fill="#ffffff" text-anchor="middle">
            ${escapeXml(data.title || '')}
          </text>

          <!-- Step Badge -->
          <g transform="translate(10, 0)">
            <circle cx="40" cy="40" r="30" fill="${data.color}" stroke="#ffffff" stroke-width="2" />
            <text x="40" y="52" font-family="system-ui, -apple-system, sans-serif" font-weight="bold" font-size="32" fill="#ffffff" text-anchor="middle">
              ${escapeXml(data.stepNum || '')}
            </text>
          </g>
        </svg>
      `.trim();
    }
  },
  {
    id: "info-card",
    name: "Info Card",
    fields: [
      { id: "title", type: "text", label: "Title", default: "Always Current" },
      { id: "bullet1", type: "text", label: "Bullet 1", default: "One cloud source" },
      { id: "bullet2", type: "text", label: "Bullet 2", default: "No more guessing" },
      { id: "color", type: "color", label: "Primary Color", default: "#2E8B57" }
    ],
    renderSvg: (data, iconSvg) => {
      const innerIcon = getSvgInner(iconSvg);
      return `
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 320" width="400" height="320">
          <!-- Main Card -->
          <rect x="20" y="50" width="360" height="240" rx="24" fill="${data.color}" />
          
          <!-- Title -->
          <text x="200" y="120" font-family="system-ui, -apple-system, sans-serif" font-weight="bold" font-size="32" fill="#ffffff" text-anchor="middle">
            ${escapeXml(data.title || '')}
          </text>
          
          <!-- Bullets -->
          <circle cx="50" cy="160" r="5" fill="#ffffff" />
          <text x="70" font-family="system-ui, -apple-system, sans-serif" font-weight="normal" font-size="24" fill="#ffffff">
            ${wrapText(data.bullet1 || '', 25, 70, 168, 28)}
          </text>

          <circle cx="50" cy="220" r="5" fill="#ffffff" />
          <text x="70" font-family="system-ui, -apple-system, sans-serif" font-weight="normal" font-size="24" fill="#ffffff">
            ${wrapText(data.bullet2 || '', 25, 70, 228, 28)}
          </text>

          <!-- Top Icon Badge -->
          <g transform="translate(150, 0)">
            <circle cx="50" cy="50" r="40" fill="#ffffff" stroke="${data.color}" stroke-width="4" />
            <g fill="${data.color}" stroke="${data.color}" transform="translate(26, 26) scale(2)">
              ${innerIcon}
            </g>
          </g>
        </svg>
      `.trim();
    }
  }
];
