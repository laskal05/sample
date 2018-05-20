const Tesseract = require('tesseract.js');

const image = './tesseract/public/sample.png';

Tesseract.recognize(image, {
  // lang: 'eng'
  lang: 'jpn'
})
.progress((p) => {
  console.log('p: ', p);
})
.then((res) => {
  console.log('res: ', res.text);
  process.exit(0);
});
