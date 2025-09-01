import fs from 'fs'
import AdmZip from 'adm-zip'
import path from 'path'
import { logger } from './logger.js'
import ExcelJS from 'exceljs'

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

export async function generateExcelFile(finalSlides, outputDir) {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Summary')

  // Set column widths
  sheet.columns = [
    { header: 'Term', key: 'term', width: 40 },
    { header: 'Definition', key: 'definition', width: 120 },
  ]

  // Border style
  const borderStyle = {
    top: { style: 'thin', color: { argb: 'D3D3D3' } },
    left: { style: 'thin', color: { argb: 'D3D3D3' } },
    bottom: { style: 'thin', color: { argb: 'D3D3D3' } },
    right: { style: 'thin', color: { argb: 'D3D3D3' } },
  }

  // Style the header row
  const headerRow = sheet.getRow(1)
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Arial', size: 12, bold: false } // bigger + medium weight
    cell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    }
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'E0E0E0' }, // light gray background
    }
    cell.border = borderStyle
  })

  let rowIndex = 2 // Start after headers

  finalSlides.forEach((section) => {
    // Merge row across 2 columns
    sheet.mergeCells(`A${rowIndex}:B${rowIndex}`)
    const headerCell = sheet.getCell(`A${rowIndex}`)

    // Apply styles
    headerCell.value = section.keys
    headerCell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    }
    headerCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'D3D3D3' }, // light gray
    }
    headerCell.border = borderStyle

    rowIndex++

    // Add term/definition rows with wrap text
    section.items.forEach((item) => {
      const termCell = sheet.getCell(`A${rowIndex}`)
      const defCell = sheet.getCell(`B${rowIndex}`)

      termCell.value = item.term
      defCell.value = item.definition

      termCell.alignment = { wrapText: true, vertical: 'top' }
      defCell.alignment = { wrapText: true, vertical: 'top' }

      rowIndex++
    })

    // Add an empty row for spacing
    rowIndex++
  })

  // Save file
  await workbook.xlsx.writeFile(path.join(outputDir, 'output.xlsx'))
  logger(
    `[Excel File Generated] Excel file generated: ${path.join(
      outputDir,
      'output.xlsx',
    )}`,
  )
}
