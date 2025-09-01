import fs from 'fs'
import path from 'path'
import { readPdfPages } from 'pdf-text-reader'
import { logger } from '../utils/logger.js'
import { processByChunk } from '../services/genai.js'
import { generateExcelFile } from '../utils/index.js'

export const parsePDF = async (filePath, outputDir) => {
  try {
    const pages = await readPdfPages({ url: filePath })
    const pageContents = pages.map((page, index) => {
      let pageText = page.lines.join(' ')
      // Remove "Wondershare PDFelement" and trim any extra whitespace
      pageText = pageText.replace(/Wondershare PDFelement/g, '').trim()

      return {
        id: `page_${index + 1}`,
        texts: [pageText],
      }
    })

    fs.mkdirSync(outputDir, { recursive: true })

    const jsonOutput = JSON.stringify(pageContents, null, 2)
    fs.writeFileSync(`${outputDir}/output-pdf.json`, jsonOutput, 'utf8')

    // Generate Excel
    const finalPDFJson = await processByChunk(pageContents)
    fs.mkdirSync(outputDir, { recursive: true })
    const outputFile = path.join(outputDir, 'final_slides.json')

    // Step 3: Save one final JSON file
    fs.writeFileSync(outputFile, JSON.stringify(finalPDFJson, null, 2))

    // Generate excel sheet
    // Step 4: Load the final JSON file and generate the excel sheet
    const jsonFileContent = fs.readFileSync(outputFile, 'utf-8')
    const loadedData = JSON.parse(jsonFileContent)
    logger(`[Success] Loaded JSON data from ${JSON.stringify(loadedData)}`)
    await generateExcelFile(loadedData, outputDir)

    return pageContents
  } catch (error) {
    logger(`Error extracting text from PDF: ${error?.message}`)
    return null
  }
}
