import { AlignmentType, Document, HeadingLevel, Paragraph, TextRun, ImageRun } from 'docx';
import got from 'got';
import sharp from 'sharp';
import path from 'path';
import fs from 'fs';

// Maximum image size in bytes (2MB)
const MAX_IMAGE_SIZE = 2 * 1024 * 1024;

// Default header styles (used as fallback)
export const DEFAULT_HEADER_STYLES = {
  h1: {
    size: 14,
    color: '000000',
    font: 'Outfit',
    bold: true,
  },
  h2: {
    size: 12,
    color: '000000',
    font: 'Outfit',
    bold: true,
  },
  h3: {
    size: 10,
    color: '000000',
    font: 'Outfit',
    bold: true,
  },
} as const;

// Template paths
export const TEMPLATE_PATHS = {
  'traffic-builders': path.join(process.cwd(), 'src/routes/wordGenerator/word-templates', 'traffic-builders.dotx'),
  'shoq': path.join(process.cwd(), 'src/routes/wordGenerator/word-templates', 'shoq.dotx'),
  'datahive': path.join(process.cwd(), 'src/routes/wordGenerator/word-templates', 'datahive.dotx'),
  'unbound-group': path.join(process.cwd(), 'src/routes/wordGenerator/word-templates', 'unbound-group.dotx'),
} as const;

// Validate and process image
export async function processImage(imageUrl: string): Promise<Buffer | null> {
  try {
    const response = await got(imageUrl, { responseType: 'buffer' });
    const imageBuffer = response.body;

    // Check file size
    if (imageBuffer.length > MAX_IMAGE_SIZE) {
      console.warn(`Image ${imageUrl} exceeds 2MB limit, skipping...`);
      return null;
    }

    // Process image with sharp
    const processedImage = await sharp(imageBuffer)
      .resize(800, 600, { fit: 'inside', withoutEnlargement: true })
      .toBuffer();

    return processedImage;
  } catch (error) {
    console.error(`Error processing image ${imageUrl}:`, error);
    return null;
  }
}

// Apply header styles
export function applyHeaderStyles(
  text: string,
  level: number,
  customStyles?: typeof DEFAULT_HEADER_STYLES
): Paragraph {
  const styles = customStyles || DEFAULT_HEADER_STYLES;
  const styleKey = `h${level}` as keyof typeof styles;
  const style = styles[styleKey];

  return new Paragraph({
    text,
    heading: getHeadingLevel(level),
    spacing: {
      before: 400,
      after: 400,
    },
    children: [
      new TextRun({
        text,
        bold: style.bold,
        size: style.size * 2, // docx uses half-points
        font: style.font,
        color: style.color,
      }),
    ],
  });
}

// Get heading level
function getHeadingLevel(level: number): HeadingLevel {
  switch (level) {
    case 1:
      return HeadingLevel.HEADING_1;
    case 2:
      return HeadingLevel.HEADING_2;
    case 3:
      return HeadingLevel.HEADING_3;
    default:
      return HeadingLevel.HEADING_1;
  }
}

// Load template
export function loadTemplate(templateName: string): Document | null {
  try {
    const templatePath = TEMPLATE_PATHS[templateName as keyof typeof TEMPLATE_PATHS];
    if (!templatePath || !fs.existsSync(templatePath)) {
      console.warn(`Template ${templateName} not found, using default document...`);
      return null;
    }
    return Document.load(templatePath);
  } catch (error) {
    console.error(`Error loading template ${templateName}:`, error);
    return null;
  }
}

// Validate header styles
export function validateHeaderStyles(styles: unknown): boolean {
  try {
    if (!styles || typeof styles !== 'object') return false;
    
    const requiredKeys = ['h1', 'h2', 'h3'] as const;
    const requiredProps = ['size', 'color', 'font', 'bold'] as const;
    
    return requiredKeys.every(key => {
      const style = (styles as Record<string, unknown>)[key];
      if (!style || typeof style !== 'object') return false;
      
      return requiredProps.every(prop => {
        const value = (style as Record<string, unknown>)[prop];
        return value !== undefined && value !== null;
      });
    });
  } catch (error) {
    console.error('Error validating header styles:', error);
    return false;
  }
}

// Modified generateSectionContent function to handle images and custom header styles
const generateSectionContent = async (section: any, config: any) => {
  const content: any[] = [];

  // Add heading if present
  if (section.heading) {
    try {
      if (config.headerStyles && validateHeaderStyles(config.headerStyles)) {
        content.push(applyHeaderStyles(section.heading, section.headingLevel || 1, config.headerStyles));
      } else {
        // Fallback to existing heading logic
        content.push(
          new Paragraph({
            text: section.heading,
            heading: getHeadingLevel(section.headingLevel || 1),
            spacing: {
              before: 400,
              after: 400,
            },
          })
        );
      }
    } catch (error) {
      console.error('Error applying header styles, using default:', error);
      content.push(
        new Paragraph({
          text: section.heading,
          heading: getHeadingLevel(section.headingLevel || 1),
          spacing: {
            before: 400,
            after: 400,
          },
        })
      );
    }
  }

  // Process content
  for (const item of section.content) {
    try {
      switch (item.type) {
        case 'image':
          if (item.url) {
            const imageBuffer = await processImage(item.url);
            if (imageBuffer) {
              content.push(
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: imageBuffer,
                      transformation: {
                        width: item.width || 400,
                        height: item.height || 300,
                      },
                    }),
                  ],
                  spacing: {
                    after: 200,
                  },
                })
              );
            } else {
              // Insert placeholder if image cannot be loaded
              content.push(
                new Paragraph({
                  text: '<place image here>',
                  spacing: {
                    after: 200,
                  },
                })
              );
            }
          }
          break;
        // ... existing code ...
      }
    } catch (error) {
      console.error(`Error processing item ${item.type}:`, error);
    }
  }

  return content;
} 