import fs from 'fs'
import dotenv from 'dotenv'
import { GoogleGenAI } from '@google/genai'
import { systemInstruction } from '../utils/ai/system-intructions.js'
import { logger } from '../utils/logger.js'

dotenv.config()
const GEMINI_API_KEY = process.env.GEMINI_API_KEY

const ai = new GoogleGenAI({ apiKey: GEMINI_API_KEY })

const RATE_LIMIT_STATUS = 429

export async function AICleanup(content) {
  const response = await ai.models.generateContent({
    model: 'gemini-1.5-flash',
    contents: content,
    config: {
      responseMimeType: 'application/json',
      temperature: 0.7,
      thinkingConfig: {
        thinkingBudget: 0,
      },
      systemInstruction: systemInstruction,
    },
  })

  return JSON.parse(response.text)
}

async function safeAICleanup(content, retries = 3, delay = 1000) {
  for (let i = 0; i < retries; i++) {
    try {
      const cleanedData = await AICleanup(content)
      return cleanedData
    } catch (error) {
      if (
        error.status === RATE_LIMIT_STATUS ||
        error.message.includes('RESOURCE_EXHAUSTED')
      ) {
        const backoffDelay = delay * Math.pow(2, i)
        logger(
          `[Rate Limit] Retrying in ${backoffDelay}ms... (Attempt ${
            i + 1
          } of ${retries})`,
        )
        await new Promise((resolve) => setTimeout(resolve, backoffDelay))
      } else {
        // For other types of errors, re-throw immediately
        throw error
      }
    }
  }
  throw new Error(
    `[Fatal] Failed to process chunk after ${retries} retries. Skipping this chunk.`,
  )
}

export async function processByChunk(slidesData) {
  const finalSlides = []
  const chunkSize = 50

  // Process the slides in chunks of 50
  for (let i = 0; i < slidesData.length; i += chunkSize) {
    const chunk = slidesData
      .slice(i, i + chunkSize)
      .flatMap((slide) => slide.texts)

    try {
      logger(
        `[Processing] Length: ${chunk.length} Processing chunk of slides...`,
      )

      // Use the new safe, retry-enabled function for API calls
      const cleanedData = await safeAICleanup(chunk)

      logger(`Cleaned Data: ${JSON.stringify(cleanedData)}`)
      console.log(`Cleaned Data: ${JSON.stringify(cleanedData)}`)

      // Add the cleaned data to the final array
      finalSlides.push({
        keys: cleanedData.keys,
        items: cleanedData.items,
      })
    } catch (error) {
      logger(`[Error] Skipping chunk due to fatal error: ${error.message}`)
      continue
    }
  }

  logger('[Success] Finished processing. Returning partial results.')
  return finalSlides
}
