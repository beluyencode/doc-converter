const { PDFNet } = require('@pdftron/pdfnet-node');
async function main() {
  await PDFNet.addResourceSearchPath('./Lib/');
  // check if the module is available
  if (!(await PDFNet.StructuredOutputModule.isModuleAvailable())) {
    console.log(123123);
    
    return;
  }
  await PDFNet.Convert.fileToWord('test.pdf', 'output.docx');
}
PDFNet.runWithCleanup(main, 'demo:1739949060645:617adfb903000000000935bcbf740717e9c6b2c11e6fd7ac9496321dc6')
.catch(err => {
  console.error(err);
})
.then(function (res) {
  PDFNet.shutdown();
});;