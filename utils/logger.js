import fs from 'fs'

const logFilePath = 'logs/app.log'

export async function logger(message) {
  try {
    const timestamp = new Date().toLocaleString()
    const logEntry = `[${timestamp}] ${message}\n`

    await fs.appendFile(logFilePath, logEntry, 'utf-8', (err) => {
      if (err) {
        console.error(`Failed to write to log file: ${err.message}`)
      }
    })
  } catch (error) {
    console.error(`Failed to write to log file: ${error.message}`)
  }
}
