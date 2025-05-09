// wordTemplateManager.ts
import fs from 'fs';
import path from 'path';

// Definieer het correcte template directory pad
const templatesDir = path.join(process.cwd(), 'word-templates');
console.log('Templates directory path:', templatesDir);

// Controleer of de directory bestaat
try {
  console.log('Templates directory exists:', fs.existsSync(templatesDir));
  if (fs.existsSync(templatesDir)) {
    console.log('Available templates:', fs.readdirSync(templatesDir));
  }
} catch (error) {
  console.error('Error checking templates directory:', error);
}

// Ensure the templates directory exists
if (!fs.existsSync(templatesDir)) {
  try {
    fs.mkdirSync(templatesDir, { recursive: true });
    console.log('Created templates directory');
  } catch (error) {
    console.error('Error creating templates directory:', error);
  }
}

// Define company templates configuration
export interface CompanyTemplate {
  id: string;
  name: string;
  description: string;
  fileName: string;
  defaultHeaderLogo?: string; // Optional backup if template doesn't contain logo
  defaultFooterText?: string; // Optional backup if template doesn't have footer
  colorPalette?: {
    primary: string;
    secondary: string;
    accent: string;
    text: string;
  };
  fontSettings?: {
    heading1: { font: string; size: number; weight?: string; color: string };
    heading2: { font: string; size: number; weight?: string; color: string };
    heading3: { font: string; size: number; weight?: string; color: string };
    normal: { font: string; size: number; weight?: string; color: string };
  };
}

// Unbound Group templates
export const COMPANY_TEMPLATES: CompanyTemplate[] = [
  {
    id: 'default',
    name: 'Default',
    description: 'A basic default template',
    fileName: 'default-template.dotx'
  },
  {
    id: 'traffic-builders',
    name: 'Traffic Builders',
    description: 'Traffic Builders corporate template',
    fileName: 'traffic-builders-template.dotx'
  },
  {
    id: 'shoq',
    name: 'Shoq',
    description: 'Shoq corporate template',
    fileName: 'shoq-template.dotx'
  },
  {
    id: 'unbound-group',
    name: 'Unbound Group',
    description: 'Unbound Group corporate template',
    fileName: 'unbound-group-template.dotx'
  }
];

/**
 * Get list of available templates
 */
export function getAvailableTemplates(): CompanyTemplate[] {
  return COMPANY_TEMPLATES;
}

/**
 * Get template by ID
 */
export function getTemplateById(templateId: string): CompanyTemplate | undefined {
  return COMPANY_TEMPLATES.find((template) => template.id === templateId);
}

/**
 * Get default template
 */
export function getDefaultTemplate(): CompanyTemplate {
  return COMPANY_TEMPLATES[0]; // Unbound Group template is default
}

/**
 * Get template file path
 */
export function getTemplateFilePath(templateId: string): string {
  const template = getTemplateById(templateId);
  if (!template) throw new Error(`Template ID not found: ${templateId}`);
  const filePath = path.join(templatesDir, template.fileName);
  if (!fs.existsSync(filePath)) throw new Error(`Template file not found: ${filePath}`);
  return filePath;
}

/**
 * Check if template file exists
 */
export function templateFileExists(templateId: string): boolean {
  const filePath = getTemplateFilePath(templateId);
  return fs.existsSync(filePath);
}

// NOTE: Template IDs and file names must match the options in the plugin config userSettings.
// If you add/remove a template here, update the plugin config as well.