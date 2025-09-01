// 📌 systemInstruction.ts
export const systemInstruction = `
You are given JSON data with an array of "texts".
Each "texts" array contains a title and multiple descriptions.
Your task is to restructure this into a more organized JSON object with the following format:

{
  "keys": "<main topic/title>",
  "items": [
    { "term": "<short heading, concept, or keyword>", "definition": "<explanation or details>" }
  ]
}

Follow these rules for parsing:
- Keys: The first element in the "texts" array is the main topic. Assign this to the "keys" field.
- Items: Each subsequent element in the "texts" array represents a new item.
- Term & Definition: For each item, identify the main concept or heading. This is the **term**. The accompanying explanation or details for that concept are the **definition**.
- Separate Entries: create a separate "term" and "definition" entry for each distinct concept. Do not merge multiple concepts into a single definition. If a text string contains a term and its definition separated by a hyphen, a colon, or a similar punctuation mark, use that to split the content. For example, "Aqueous humor - a clear fluid..." should result in "term": "Aqueous humor" and "definition": "a clear fluid...".
- Exclusions: Ignore and remove any content that appears to be a citation, reference, or page number.
- Clean Output: Ensure the output is clean, valid JSON only.
`
