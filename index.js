import express from 'express'
import path from 'path'
import multer from 'multer'
import fs from 'fs'
import { fileURLToPath } from 'url'
import { logger } from './utils/logger.js'
import { parseFile } from './main.js'

const API_KEY = process.env.APP_API_KEY || 'my-secret-key'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const app = express()
const PORT = 3000

// Set a higher limit for the request body size to handle large file uploads
app.use(express.json({ limit: '100mb' }))
app.use(express.urlencoded({ limit: '100mb', extended: true }))

// Set up Multer for file uploads
const uploadDir = 'uploads'
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir)
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir)
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`)
  },
})

const upload = multer({ storage: storage })

app.post('/upload', upload.single('pptxFile'), async (req, res) => {
  // 🔑 Check API key from form input
  const clientKey = req.body.apiKey
  if (!clientKey || clientKey !== API_KEY) {
    return res.status(403).send('Unauthorized: Invalid API key')
  }

  if (!req.file) {
    return res.status(400).send('No file uploaded.')
  }

  try {
    logger(`[Info] File uploaded: ${req.file.originalname}`)
    const filePath = req.file.path
    // Get base filename (without extension)
    const baseName = path.basename(filePath, path.extname(filePath))
    const outputDir = path.join(process.cwd(), 'output', baseName)

    await parseFile(req.file.path, outputDir)

    const outputFilePath = path.join(outputDir, 'output.xlsx')
    const filename = 'slides-output.xlsx'

    // Download the generated file
    res.download(outputFilePath, filename, (err) => {
      if (err) {
        logger(`Error downloading PPTX file: ${JSON.stringify(err)}`)

        res.status(500).send('An error occurred during file download.')
      }
    })
  } catch (error) {
    console.error('Error processing PPTX file:', error)
    logger(`Error processing PPTX file: ${error?.message}`)
    res.status(500).send('An error occurred during file processing.')
  } finally {
    // Clean up the uploaded file
    fs.unlink(req.file.path, (err) => {
      if (err) console.error('Error deleting uploaded file:', err)
    })
  }
})

// Serve the HTML client for file upload
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'))
})

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running at http://localhost:${PORT}`)
})
