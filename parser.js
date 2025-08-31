#!/usr/bin/env node
import PptxParser from 'node-pptx-parser'
import AdmZip from 'adm-zip'
import fs from 'fs'
import path from 'path'
import * as XLSX from 'xlsx'
import xml2js from 'xml2js'
import { extractPPT } from './utils/index.js'

const outputDir = 'extracted_pptx'
const notesDir = `${outputDir}/ppt/notesSlides`
const outputJsonPath = `output/json/notes.json`

// Helper function to parse flashcards from a string
function parseFlashcards(text) {
  const lines = text
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0)

  const flashcards = []
  let currentCard = null

  lines.forEach((line) => {
    const match = line.match(/^(.+?)\s*-\s*(.+)$/) // matches "Term - Definition"
    if (match) {
      // Push previous card if exists
      if (currentCard) flashcards.push(currentCard)

      // Start a new card
      currentCard = {
        term: match[1].trim(),
        definition: match[2].trim(),
      }
    } else if (currentCard) {
      // Append line to current card's definition
      currentCard.definition += '\n' + line
    }
  })

  // Push the last card
  if (currentCard) flashcards.push(currentCard)

  return flashcards
}

async function extractImages(pptxPath, outputDir = 'output-images') {
  const zip = new AdmZip(pptxPath)
  const entries = zip.getEntries()

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true })
  }

  entries.forEach((entry) => {
    if (entry.entryName.startsWith('ppt/media/')) {
      const fileName = path.basename(entry.entryName)
      const filePath = path.join(outputDir, fileName)
      fs.writeFileSync(filePath, entry.getData())
      console.log(`🖼️ Extracted: ${filePath}`)
    }
  })
}

async function generateXLSX(parsedText) {
  // Create a JSON-friendly structure
  const slides = parsedText
    .filter((slide) => slide.text.length > 1)
    .map((slide) => ({
      id: slide.id,
      texts: parseFlashcards(slide.text[1]),
    }))

  // Save JSON file
  fs.writeFileSync(
    'output/json/slides.json',
    JSON.stringify(slides, null, 2),
    'utf-8',
  )

  // Flatten all flashcards into rows
  const rows = []

  // Loop through slides
  slides.forEach((slide) => {
    slide.texts.forEach((flashcard) => {
      rows.push([flashcard.term || '', flashcard.definition || ''])
    })
  })

  // Optional: add headers
  rows.unshift(['Title', 'Content'])

  // Create worksheet
  const worksheet = XLSX.utils.aoa_to_sheet(rows)
  // Set column widths (approximate character count)
  worksheet['!cols'] = [
    { wch: 40 }, // Title column width
    { wch: 100 }, // Content column width
  ]

  // Apply text wrap to Definition column (column B)
  for (let R = 1; R <= rows.length; ++R) {
    // start from 1 to skip header
    const cellAddress = `B${R + 1}` // B column, 1-indexed
    if (worksheet[cellAddress]) {
      worksheet[cellAddress].s = {
        alignment: { wrapText: true, vertical: 'top' },
      }
    }
  }

  // Create a new workbook and append worksheet
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Slides')

  // Write to Excel file
  XLSX.writeFile(workbook, 'output/slides.xlsx')
}

async function generateXLSXNotes(parsedText) {
  // Flatten all flashcards into rows
  const rows = []

  // Loop through slides
  parsedText.forEach((slide) => {
    slide.texts.forEach((item) => {
      // Check if the definition is an array (nested structure)
      if (Array.isArray(item.definition)) {
        item.definition.forEach((subItem) => {
          rows.push([subItem.term || '', subItem.definition || ''])
        })
      } else {
        // Handle the simple term/definition structure
        rows.push([item.term || '', item.definition || ''])
      }
    })
  })

  // Optional: add headers
  rows.unshift(['Title', 'Content'])

  // Create worksheet
  const worksheet = XLSX.utils.aoa_to_sheet(rows)
  // Set column widths (approximate character count)
  worksheet['!cols'] = [
    { wch: 40 }, // Title column width
    { wch: 100 }, // Content column width
  ]

  // Apply text wrap to Definition column (column B)
  for (let R = 1; R <= rows.length; ++R) {
    // start from 1 to skip header
    const cellAddress = `B${R + 1}` // B column, 1-indexed
    if (worksheet[cellAddress]) {
      worksheet[cellAddress].s = {
        alignment: { wrapText: true, vertical: 'top' },
      }
    }
  }

  // Create a new workbook and append worksheet
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Slides')

  // Write to Excel file
  XLSX.writeFile(workbook, 'output/slides-notes.xlsx')
}

async function extractNotes(file) {
  // Create the extraction directory
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true })
  }

  const zip = new AdmZip(file)
  zip.extractAllTo(outputDir, true)
}
// Helper function to format the sub-definitions as an array of objects
function formatSubDefinitions(lines) {
  const output = []
  let currentTerm = null
  let currentDefinition = ''

  lines.forEach((line) => {
    const trimmedLine = line.trim()
    // Check for new sub-points starting with "o " followed by a hyphen
    const match = trimmedLine.match(/^o\s+(.*?)\s+-\s+(.*)/)
    if (match) {
      // If a previous term/def existed, push it to the array
      if (currentTerm !== null) {
        output.push({
          term: currentTerm,
          definition: currentDefinition.trim(),
        })
      }
      // Start a new sub-term
      currentTerm = match[1].trim()
      currentDefinition = match[2].trim()
    } else {
      // Append line to the current definition
      if (currentTerm !== null) {
        currentDefinition += ' ' + trimmedLine
      }
    }
  })

  // Push the last term/definition
  if (currentTerm !== null) {
    output.push({
      term: currentTerm,
      definition: currentDefinition.trim(),
    })
  }

  return output
}

// Function to parse a single notes XML file
async function parseNotesFile(filePath) {
  const xml = fs.readFileSync(filePath, 'utf-8')
  return new Promise((resolve, reject) => {
    xml2js.parseString(xml, (err, result) => {
      if (err) return reject(err)

      let finalNotes = []
      let mainTerm = ''
      let subDefinitions = []

      try {
        const shapeTree = result['p:notes']['p:cSld'][0]['p:spTree'][0]
        const shapes = shapeTree['p:sp'] || []

        let allText = ''
        shapes.forEach((sp) => {
          if (sp['p:txBody'] && sp['p:txBody'][0]['a:p']) {
            sp['p:txBody'][0]['a:p'].forEach((p) => {
              let paragraphText = ''
              if (p['a:r']) {
                p['a:r'].forEach((r) => {
                  paragraphText += r['a:t'][0] || ''
                })
              }
              if (paragraphText) {
                const pPr = p['a:pPr'] ? p['a:pPr'][0] : {}
                const indentLevel = pPr.$.lvl ? parseInt(pPr.$.lvl, 10) : 0
                const bullet = pPr['a:buChar'] ? pPr['a:buChar'][0].$.char : ''

                let prefix = ''
                if (bullet) {
                  prefix = ' '.repeat(indentLevel * 2) + bullet + ' '
                } else if (indentLevel > 0) {
                  prefix = ' '.repeat(indentLevel * 2)
                }

                allText += prefix + paragraphText.trim() + '\n'
              }
            })
          }
        })

        const lines = allText.split('\n').filter((line) => line.trim() !== '')

        lines.forEach((line) => {
          const trimmedLine = line.trim()

          // Check for a new main term (starts with a bullet '∙')
          if (trimmedLine.startsWith('∙ ')) {
            // If we have a previous main term, process and save it
            if (mainTerm !== '') {
              finalNotes.push({
                term: mainTerm,
                definition: formatSubDefinitions(subDefinitions),
              })
            }
            mainTerm = trimmedLine.substring(2)
            subDefinitions = []
          } else {
            // All other lines are considered part of the current definition
            subDefinitions.push(line)
          }
        })

        // Push the last main term and its definitions
        if (mainTerm !== '') {
          finalNotes.push({
            term: mainTerm,
            definition: formatSubDefinitions(subDefinitions),
          })
        }

        // Final cleanup for terms with hyphens and empty definitions (like slide 178)
        const cleanedNotes = finalNotes.map((note) => {
          if (
            Array.isArray(note.definition) &&
            note.definition.length === 0 &&
            note.term.includes(' - ')
          ) {
            const parts = note.term.split(' - ', 2)
            return {
              term: parts[0].trim(),
              definition: parts[1].trim(),
            }
          }
          return note
        })

        resolve(cleanedNotes)
      } catch (e) {
        console.warn(`Error processing file ${filePath}: ${e.message}`)
        resolve([])
      }
    })
  })
}

// Main function to process all notes files
async function processAllNotes() {
  const finalOutput = []
  try {
    const files = fs.readdirSync(notesDir)
    const notesFiles = files.filter(
      (file) => file.startsWith('notesSlide') && file.endsWith('.xml'),
    )

    for (const file of notesFiles) {
      const filePath = path.join(notesDir, file)
      const slideNumber = parseInt(
        file.replace('notesSlide', '').replace('.xml', ''),
        10,
      )
      const notes = await parseNotesFile(filePath)

      if (notes.length > 0) {
        finalOutput.push({
          id: slideNumber.toString(),
          texts: notes,
        })
      }
    }

    fs.writeFileSync(outputJsonPath, JSON.stringify(finalOutput, null, 2))
    console.log(`All notes saved to ${outputJsonPath}`)
  } catch (err) {
    console.error(`Error processing notes: ${err}`)
  }
}

export function readAndMergeXLSX(file1Path, file2Path, outputFilePath) {
  try {
    // Check if files exist before trying to read them
    if (!fs.existsSync(file1Path)) {
      console.error(`❌ Error: File not found at ${file1Path}`)
      return
    }
    if (!fs.existsSync(file2Path)) {
      console.error(`❌ Error: File not found at ${file2Path}`)
      return
    }

    // Read the first workbook. Using .read() with a file buffer is more robust
    // for this type of environment.
    const workbook1 = XLSX.read(fs.readFileSync(file1Path))
    const sheetName1 = workbook1.SheetNames[0]
    const sheet1Data = XLSX.utils.sheet_to_json(workbook1.Sheets[sheetName1], {
      header: 1,
    })

    // Read the second workbook using the same method
    const workbook2 = XLSX.read(fs.readFileSync(file2Path))
    const sheetName2 = workbook2.SheetNames[0]
    const sheet2Data = XLSX.utils.sheet_to_json(workbook2.Sheets[sheetName2], {
      header: 1,
    })

    // Combine the data, including the header row from the first file
    const combinedData = sheet1Data.concat(sheet2Data.slice(1))

    // Create a new worksheet from the combined data
    const newWorksheet = XLSX.utils.aoa_to_sheet(combinedData)

    // Set column widths (approximate character count)
    newWorksheet['!cols'] = [
      { wch: 40 }, // Title column width
      { wch: 100 }, // Content column width
    ]

    // Create a new workbook and append the new worksheet
    const newWorkbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Merged Data')

    // Write the new workbook to the output file
    XLSX.writeFile(newWorkbook, outputFilePath)

    console.log(
      `✅ Successfully merged ${path.basename(file1Path)} and ${path.basename(
        file2Path,
      )} into ${path.basename(outputFilePath)}`,
    )
  } catch (error) {
    console.error('❌ Error merging files:', error.message)
  }
}

export async function parse(filePath) {
  try {
    const fullPath = path.resolve(filePath)
    const baseName = path.basename(fullPath, '.pptx')

    // Notes Extraction
    extractPPT(fullPath)
    processAllNotes()

    const parser = new PptxParser(fullPath)
    const parsedText = await parser.extractText()

    generateXLSX(parsedText)

    // Excel Sheet for Notes
    const jsonData = fs.readFileSync(outputJsonPath, 'utf-8')
    const parsedNotes = JSON.parse(jsonData)
    await generateXLSXNotes(parsedNotes)

    readAndMergeXLSX(
      path.join('output', 'slides-notes.xlsx'),
      path.join('output', 'slides.xlsx'),
      path.join('output', 'slides-output.xlsx'),
    )

    // // Extract images into folder with PPTX name
    // const imagesDir = `${baseName}-images`
    // await extractImages(fullPath, imagesDir)
    // console.log(`📂 Images saved in: ${imagesDir}`)
  } catch (err) {
    console.error('❌ Error parsing PPTX:', err.message)
  }
}
