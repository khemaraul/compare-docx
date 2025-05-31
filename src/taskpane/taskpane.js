/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    let body = document.body;
    const main = document.createElement("div");
    main.className = 'main-div';
    body.appendChild(main);
    const header = document.createElement("div");
    header.className = 'title';
    header.innerText = 'Compare-DOCX';
    main.appendChild(header);
    const form = document.createElement("form");
    form.className = 'form';
    main.appendChild(form);
    let uploadMainSection = document.createElement("div");
    uploadMainSection.className = 'upload-main-section';
    form.appendChild(uploadMainSection);
    let uploadSectionOne = document.createElement("div");
    uploadSectionOne.className = 'upload-section';
    let uploadTitle1 = document.createElement("div");
    uploadTitle1.innerText = 'File #1 📁';
    let subTitle1 = document.createElement("div");
    subTitle1.innerText = 'Please select file to upload.';
    let allowedText1 = document.createElement("div");
    allowedText1.innerText = '(Only .docx file allowed)';
    let fileInput1 = document.createElement('input');
    fileInput1.type = 'file';
    fileInput1.accept = ".docx";         // Allow only .docx files
    fileInput1.required = true;
    fileInput1.style.display = "none";
    const uploadBtn1 = document.createElement("button");
    uploadBtn1.className = 'upload-btn';
    uploadBtn1.textContent = "Upload File";
    const fileName1 = document.createElement("span");
    fileName1.className = 'filename';
    uploadSectionOne.appendChild(uploadTitle1);
    uploadSectionOne.appendChild(subTitle1);
    uploadSectionOne.appendChild(allowedText1);
    uploadSectionOne.appendChild(fileInput1);
    uploadSectionOne.appendChild(uploadBtn1);
    uploadSectionOne.appendChild(fileName1);
    uploadMainSection.append(uploadSectionOne);
    let uploadSectionTwo = document.createElement("div");
    uploadSectionTwo.className = 'upload-section';
    let uploadTitle2 = document.createElement("div");
    uploadTitle2.innerText = 'File #2 📁';
    let subTitle2 = document.createElement("div");
    subTitle2.innerText = 'Please select file to upload.';
    let allowedText2 = document.createElement("div");
    allowedText2.innerText = '(Only .docx file allowed)';
    let fileInput2 = document.createElement('input');
    fileInput2.type = 'file';
    fileInput2.accept = ".docx";         // Allow only .docx files
    fileInput2.required = true;
    fileInput2.style.display = "none";
    const uploadBtn2 = document.createElement("button");
    uploadBtn2.className = 'upload-btn';
    uploadBtn2.textContent = "Upload File";
    const fileName2 = document.createElement("span");
    fileName2.className = 'filename';
    uploadSectionTwo.appendChild(uploadTitle2);
    uploadSectionTwo.appendChild(subTitle2);
    uploadSectionTwo.appendChild(allowedText2);
    uploadSectionTwo.appendChild(fileInput2);
    uploadSectionTwo.appendChild(uploadBtn2);
    uploadSectionTwo.appendChild(fileName2);
    uploadMainSection.append(uploadSectionTwo);

    let submitBtn = document.createElement("input");
    submitBtn.className = "submit-btn";
    submitBtn.type = "submit";
    submitBtn.value = "Compare";
    form.appendChild(submitBtn);

    let fileOne;
    let fileTwo;

    // Hover effect (via JavaScript event)
    submitBtn.addEventListener("mouseover", () => {
      submitBtn.style.backgroundColor = "#45a049";
    });
    submitBtn.addEventListener("mouseout", () => {
      submitBtn.style.backgroundColor = "#4CAF50";
    });

    uploadBtn1.addEventListener("click", function () {
      fileInput1.click();
    });

    fileInput1.addEventListener("change", function () {
      fileOne = fileInput1.files[0];
      fileName1.innerHTML = fileInput1.files[0].name;
      console.log("++++",fileOne);
    });

    uploadBtn2.addEventListener("click", function () {
      fileInput2.click();
    });

    fileInput2.addEventListener("change", function () {
      fileTwo = fileInput2.files[0];
      fileName2.innerText = fileInput2.files[0].name;
    });

    form.addEventListener("submit", async (event) => {
      event.preventDefault(); // Prevent page refresh or form submission
      if (!fileOne || !fileTwo) {
        console.log("Please upload both versions.");
        return;
      }

      try{
        const text1 = await readDocx(fileOne);
        const text2 = await readDocx(fileTwo);

        // Compare text using diffWords
        const diff = Diff.diffWords(text1, text2);

        // Create HTML with highlights
        let resultHtml = '';
        diff.forEach(part => {
          const color = part.added ? 'green' :
                        part.removed ? 'red' : 'black';
          const weight = (part.added || part.removed) ? 'bold' : 'normal';
          resultHtml += `<span style="color:${color}; font-weight:${weight};">${part.value}</span>`;
        });


        await Word.run(async (context) => {
          const body = context.document.body;
      
          body.clear();
          body.insertParagraph("Document Comparison", Word.InsertLocation.start)
              .font.set({ bold: true, size: 18 });
      
          body.insertHtml(resultHtml, Word.InsertLocation.end);
      
          await context.sync();
        });
      }catch (err) {
        console.error("Error comparing documents:", err);
        console.log("An error occurred while reading or comparing files.");
      }
    });

    async function readDocx(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async (event) => {
          const arrayBuffer = event.target.result;
          mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
            .then(result => resolve(result.value))
            .catch(err => reject(err));
        };
        reader.readAsArrayBuffer(file);
      });
    }
    
  }
});

// export async function run() {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */

//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";

//     await context.sync();
//   });
// }
