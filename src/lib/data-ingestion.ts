import type { DataCategory, ExpenseSubCategory, ParsedFileData } from "./data-parser"

export const INGESTION_METADATA_KEY = "data_ingestion_files"
export const INGESTION_PARSED_DATA_KEY_PREFIX = "data_ingestion_file:"

export type IngestionFileMetadata = {
  fileId: string
  categoryType: DataCategory
  customLabel: string
  fileName: string
  fileType: "xlsx" | "csv" | "pdf"
  uploadTimestamp: string
  subCategory: ExpenseSubCategory | null
  parsedData?: ParsedFileData
}
