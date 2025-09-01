import fs from 'fs'
import path from 'path'
import xml2js from 'xml2js'
import { extractPPT, generateExcelFile } from '../utils/index.js'
import { logger } from '../utils/logger.js'
import { fileURLToPath } from 'url'
import { processByChunk } from '../services/genai.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

export const parsePPT = async (filePath, outputDir) => {
  // Extract PPT
  const extractedDirectory = await extractPPT(filePath)

  // Process XML for notes and slides
  const parsedNotesData = await parsePPTNotes(extractedDirectory)
  const parsedSlidesData = await parsePPTSlides(extractedDirectory)

  // Process the json output
  const finalNotes = await processByChunk(parsedNotesData)
  const finalSlides = await processByChunk(parsedSlidesData)

  fs.mkdirSync(outputDir, { recursive: true })
  const outputFile = path.join(outputDir, 'final_slides.json')

  // Step 3: Save one final JSON file
  fs.writeFileSync(
    outputFile,
    JSON.stringify([...finalNotes, ...finalSlides], null, 2),
  )
  logger(`[Success] Final structured JSON saved as ${outputFile}`)

  // Generate excel sheet
  // Step 4: Load the final JSON file and generate the excel sheet
  const jsonFileContent = fs.readFileSync(outputFile, 'utf-8')
  const loadedData = JSON.parse(jsonFileContent)
  await generateExcelFile(loadedData, outputDir)

  logger(
    `[Success] Excel sheet generated as ${outputFile.replace(
      '.json',
      '.xlsx',
    )}`,
  )
}

async function parsePPTNotes(extractedPptxPath) {
  try {
    const notesSlidesDir = path.join(extractedPptxPath, 'ppt', 'notesSlides')
    logger(`Starting to parse notes from directory: ${notesSlidesDir}`)

    const notesFiles = fs.readdirSync(notesSlidesDir)
    const parser = new xml2js.Parser()

    const parsedData = []

    for (const fileName of notesFiles) {
      // We are only interested in the notesSlide*.xml files
      if (fileName.startsWith('notesSlide') && fileName.endsWith('.xml')) {
        const filePath = path.join(notesSlidesDir, fileName)
        const fileContent = fs.readFileSync(filePath, 'utf-8')

        // Convert the XML content to a JavaScript object
        const result = await parser.parseStringPromise(fileContent)

        const slideId = `slide_${fileName.match(/\d+/)[0]}`
        const texts = []

        // Check if the p:notes element exists before trying to access its children
        const notes = result['p:notes']
        if (
          !notes ||
          !notes['p:cSld'] ||
          !notes['p:cSld'][0]['p:spTree'] ||
          !notes['p:cSld'][0]['p:spTree'][0]['p:sp']
        ) {
          logger(`No text shapes found in file: ${fileName}`)
          continue
        }

        // Iterate through all 'p:sp' elements to find the one with a text body
        const shapeElements = notes['p:cSld'][0]['p:spTree'][0]['p:sp']

        for (const shape of shapeElements) {
          if (shape['p:txBody'] && shape['p:txBody'][0]['a:p']) {
            for (const p of shape['p:txBody'][0]['a:p']) {
              if (p['a:r']) {
                // Concatenate all text from run elements within a paragraph
                const fullText = p['a:r'].map((run) => run['a:t'][0]).join('')
                // Only push non-empty strings
                if (fullText.trim() !== '') {
                  texts.push(fullText)
                }
              }
            }
          }
        }

        // Add the array of texts directly to the parsed data
        if (texts.length > 0) {
          parsedData.push({ id: slideId, texts: texts })
        }
      }
    }

    // Define the output directory and file path
    const outputDir = path.join(__dirname, 'output')
    const outputPath = path.join(outputDir, 'ppt_notes.json')

    // Create the output directory if it doesn't exist
    fs.mkdirSync(outputDir, { recursive: true })
    fs.writeFileSync(outputPath, JSON.stringify(parsedData, null, 2))

    logger(
      `[Success] Successfully parsed PowerPoint notes and wrote to ${outputPath}.`,
    )

    return parsedData
  } catch (error) {
    logger(`[Error] Failed to parse PowerPoint notes: ${error.message}`)
  }
}

export async function parsePPTSlides(extractedPptxPath) {
  try {
    const slidesDir = path.join(extractedPptxPath, 'ppt', 'slides')
    logger(`Starting to parse slides from directory: ${slidesDir}`)

    const slideFiles = fs.readdirSync(slidesDir)
    const parser = new xml2js.Parser()

    const parsedData = []

    for (const fileName of slideFiles) {
      // We are only interested in the slide*.xml files
      if (fileName.startsWith('slide') && fileName.endsWith('.xml')) {
        const filePath = path.join(slidesDir, fileName)
        const fileContent = fs.readFileSync(filePath, 'utf-8')

        // Extract the slide ID from the file name (e.g., 'slide1.xml' -> 1)
        const slideId = parseInt(
          fileName.replace('slide', '').replace('.xml', ''),
        )

        const result = await parser.parseStringPromise(fileContent)

        const slideText = []

        // Traverse the XML to find text content. The path is typically
        // `p:sld.p:cSld[0].p:spTree[0].p:sp[i].p:txBody[0].a:p[j].a:r[k].a:t[0]`
        // This is a common but not the only path for text.
        const slideRoot = result['p:sld']
        if (
          slideRoot &&
          slideRoot['p:cSld'] &&
          slideRoot['p:cSld'][0]['p:spTree']
        ) {
          const shapeTree = slideRoot['p:cSld'][0]['p:spTree'][0]

          if (shapeTree['p:sp']) {
            for (const shape of shapeTree['p:sp']) {
              if (shape['p:txBody'] && shape['p:txBody'][0]['a:p']) {
                for (const paragraph of shape['p:txBody'][0]['a:p']) {
                  if (paragraph['a:r']) {
                    for (const run of paragraph['a:r']) {
                      if (run['a:t'] && run['a:t'][0]) {
                        // Push the text to the array
                        slideText.push(run['a:t'][0])
                      }
                    }
                  }
                }
              }
            }
          }
        }

        parsedData.push({
          id: slideId,
          texts: slideText,
        })
      }
    }

    logger('[Success] Finished parsing slide data.')

    // Define the output directory and file path
    const outputDir = path.join(__dirname, 'output')
    const outputPath = path.join(outputDir, 'ppt_content.json')

    // Create the output directory if it doesn't exist
    fs.mkdirSync(outputDir, { recursive: true })
    fs.writeFileSync(outputPath, JSON.stringify(parsedData, null, 2))

    return parsedData
  } catch (e) {
    logger(
      `[Error] Something went wrong parsing the slides. Error Message: ${e.message}`,
    )
    return []
  }
}
