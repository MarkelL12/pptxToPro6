import { getTextExtractor } from 'office-text-extractor'

const extractor = getTextExtractor()
const path = './ppt.pptx'
const text = await extractor.extractText({ input: path, type: 'file' })
const textSlides = text.split("---");

//todo convert each slide to plainText Base64 encoded
//todo convert each slide to RTF and then to Base64 encoded

//todo take Pro6 file template and insert slides

console.log(slides)