import { extendZodWithOpenApi } from '@asteasolutions/zod-to-openapi';
import { z } from 'zod';

extendZodWithOpenApi(z);

// Define Word Generator Response Schema
export type WordGeneratorResponse = z.infer<typeof WordGeneratorResponseSchema>;
export const WordGeneratorResponseSchema = z.object({
  filepath: z.string().openapi({
    description: 'The file path where the generated Word document is saved.',
  }),
});

// Define Cell Schema
const CellSchema = z.object({
  text: z.string().optional().openapi({
    description: 'Text content within a cell.',
  }),
});

// Define Row Schema
const RowSchema = z.object({
  cells: z.array(CellSchema).optional().openapi({
    description: 'Array of cells within a row.',
  }),
});

// New types for header styles
export const HeaderStyleSchema = z.object({
  size: z.number().min(8).max(72),
  color: z.string().regex(/^[0-9A-Fa-f]{6}$/),
  font: z.string(),
  bold: z.boolean(),
});

export const HeaderStylesSchema = z.object({
  h1: HeaderStyleSchema,
  h2: HeaderStyleSchema,
  h3: HeaderStyleSchema,
});

// New type for image content
export const ImageContentSchema = z.object({
  type: z.literal('image'),
  url: z.string().url(),
  alt: z.string().optional(),
  width: z.number().optional(),
  height: z.number().optional(),
});

// New type for template selection
export const TemplateSchema = z.enum([
  'traffic-builders',
  'shoq',
  'datahive',
  'unbound-group',
  'default'
]);

// Extend existing content type to include images
export const ContentSchema = z.object({
  type: z.enum(['paragraph', 'listing', 'table', 'pageBreak', 'emptyLine', 'image']),
  text: z.string().optional(),
  items: z.array(z.string()).optional(),
  headers: z.array(z.string()).optional(),
  rows: z.array(z.object({
    cells: z.array(z.object({
      text: z.string(),
    })),
  })).optional(),
  url: z.string().url().optional(),
  alt: z.string().optional(),
  width: z.number().optional(),
  height: z.number().optional(),
});

// Define the base schema for a section
const SectionSchema = z.object({
  sectionId: z.string().openapi({
    description: 'A unique identifier for the section.',
  }),
  heading: z.string().optional().openapi({
    description: 'Heading of the section.',
  }),
  headingLevel: z.number().int().min(1).optional().openapi({
    description: 'Level of the heading (e.g., 1 for main heading, 2 for subheading).',
  }),
  parentSectionId: z.string().optional().openapi({
    description:
      'The unique identifier of the parent section, if this section is a child of another. Leave empty if this section has no parent.',
  }),
  content: z.array(ContentSchema).optional().openapi({
    description: 'Content contained within the section, including paragraphs, tables, etc.',
  }),
});

// Extend existing request body schema
export const WordGeneratorRequestBodySchema = z.object({
  title: z.string(),
  header: z.object({
    text: z.string(),
    alignment: z.enum(['left', 'center', 'right']),
  }).optional(),
  footer: z.object({
    text: z.string(),
    alignment: z.enum(['left', 'center', 'right']),
  }).optional(),
  sections: z.array(z.object({
    sectionId: z.string(),
    content: z.array(ContentSchema),
    heading: z.string().optional(),
    headingLevel: z.number().min(1).max(3).optional(),
    parentSectionId: z.string().optional(),
  })),
  wordConfig: z.object({
    fontSize: z.number(),
    lineHeight: z.number(),
    fontFamily: z.string(),
    showPageNumber: z.boolean(),
    showTableOfContent: z.boolean(),
    showNumberingInHeader: z.boolean(),
    numberingReference: z.string(),
    pageOrientation: z.enum(['portrait', 'landscape']),
    margins: z.string(),
    // New optional fields
    headerStyles: HeaderStylesSchema.optional(),
    template: TemplateSchema.optional(),
  }),
});

export type WordGeneratorRequestBody = z.infer<typeof WordGeneratorRequestBodySchema>;
