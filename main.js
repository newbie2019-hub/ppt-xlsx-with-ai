import { parsePPT } from './parser/ppt.js'
import { parsePDF } from './parser/pdf.js'
import { parseDocs } from './parser/docs.js'
import { fileTypeFromFile } from 'file-type'
import { logger } from './utils/logger.js'

export const parseFile = async (filePath, outputDir) => {
  const type = await fileTypeFromFile(filePath)

  if (!type) {
    console.log('[Error] Could not determine file type.')
    logger(`[Error] ParseFile: Could not determine file type for ${filePath}`)
    return
  }

  switch (type.mime) {
    case 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
      logger(`Detected PPT file: ${filePath}`)
      await parsePPT(filePath, outputDir)
      break
    case 'application/pdf':
      logger(`Detected PDF file: ${filePath}`)
      await parsePDF(filePath, outputDir)
      break
    case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
      await parseDocs(filePath, outputDir)
      break
    default:
      console.log(`Unsupported file type: ${type.mime}`)
      break
  }
}
