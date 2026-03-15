/**
 * Shared file type → Fluent UI icon name map.
 * Used by SearchResultsBlock, DocumentLibraryBlock, and FilePreviewBlock.
 */

export const FILE_TYPE_ICONS: Record<string, string> = {
  docx: 'WordDocument',
  doc: 'WordDocument',
  xlsx: 'ExcelDocument',
  xls: 'ExcelDocument',
  pptx: 'PowerPointDocument',
  ppt: 'PowerPointDocument',
  pdf: 'PDF',
  one: 'OneNoteLogo',
  aspx: 'SharepointLogo',
  folder: 'FabricFolder',
  default: 'Document'
};

/** Get icon name for a file extension, with fallback to 'Document'. */
export function getFileTypeIcon(ext?: string): string {
  if (!ext) return FILE_TYPE_ICONS.default;
  return FILE_TYPE_ICONS[ext.toLowerCase()] || FILE_TYPE_ICONS.default;
}
