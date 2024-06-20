import { getTextExtractor } from 'office-text-extractor'
import fs from 'fs';


const extractor = getTextExtractor()
const path = './ppt.pptx'
const text = await extractor.extractText({ input: path, type: 'file' })
const textSlides = text.split("---");

var b64Slides = [];
textSlides.forEach(slide => {
    let b64Slide = Buffer.from(slide).toString('base64');
    b64Slides.push(b64Slide);
})

var b64RTFSlides = [];
textSlides.forEach(slide =>{ //might not be working
    const convertToRTF = (text) => {
        let rtf = '{\\rtf1\\ansi\n';
        rtf += text.split('\n').map(line => line.replace(/\\/g, '\\\\').replace(/{/g, '\\{').replace(/}/g, '\\}')).join('\\par\n');
        rtf += '}';
        return rtf;
    };
    let rtfContent = convertToRTF(slide);
    b64RTFSlides.push(rtfContent);
})
const slideHeader = fs.readFileSync('./presentationSrc/presentationHeader.txt').toString();
const slideFooter = fs.readFileSync('./presentationSrc/presentationFooter.txt').toString();
const slideTemplate = fs.readFileSync('./presentationSrc/presentationSlide.txt').toString();
var pro6Slides = [];
var i = 0;

textSlides.forEach(slide =>{
    let plainTextReplaced = slideTemplate.replace('<NSString rvXMLIvarName="PlainText"></NSString>', '<NSString rvXMLIvarName="PlainText">' + b64Slides[i] + '</NSString>')
    let RTFReplaced = plainTextReplaced.replace('<NSString rvXMLIvarName="RTFData"></NSString>', '<NSString rvXMLIvarName="RTFData">' + b64RTFSlides[i] + '</NSString>')
    pro6Slides.push(RTFReplaced)
    i+= 1;
})

const slides = pro6Slides.join('');

const presentationString = slideHeader + slides + slideFooter;

fs.writeFile('./pro6.pro6', presentationString, err => {
    if (err) {
      console.error(err);
    } else {
      // file written successfully
    }
  });