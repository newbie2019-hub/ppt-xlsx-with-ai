import fs from 'fs'
import AdmZip from 'adm-zip'
import path from 'path'
import { logger } from './logger.js'

export async function extractPPT(file, outputDir = 'extracted_pptx') {
  try {
    // Get the base name of the file (e.g., 'presentation.pptx' -> 'presentation')
    const filename = path.parse(file).name
    const outputPath = path.join(outputDir, filename)

    if (!fs.existsSync(outputPath)) {
      fs.mkdirSync(outputPath, { recursive: true })
    }

    const zip = new AdmZip(file)
    zip.extractAllTo(outputPath, true)

    logger(`[Success] Extracted the PPT successfully to ${outputPath}`)

    return outputPath
  } catch (e) {
    logger(
      `[Error] Something went wrong extracting the PPT. Error Message: ${e.message}`,
    )
  }
}
