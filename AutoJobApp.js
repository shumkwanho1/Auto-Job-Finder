import puppeteer from 'puppeteer-core';
import fs from 'fs';
import xlsx from 'xlsx';


const workbook = xlsx.readFile('Apply History.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

const sequenceJSON = JSON.parse(fs.readFileSync('./sequence.json', 'utf8'));
let sequence = sequenceJSON.sequence;
// if duplicate -> continue
// else save to excel

(async () => {
  //  1.  Open Browser
  const browser = await puppeteer.launch({
    executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    headless: false,
  });
  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1400, height: 978 });


    //  2. loop over pages
    let pageNumber = 1;
    let isLastPage = false;

    while (!isLastPage) {
    await page.goto(`https://hk.jobsdb.com/programmer-jobs?page=${pageNumber}&sortmode=ListedDate`);
    await delay(500)
    const nextBtn = page.locator('a[title="Next"]');
    if (!nextBtn) {
      isLastPage = true;
      break;
    }

    //   2.1 loop over post
    for (let i = 1; i <= 32; i++) {
      // 2.2 get gob info
      const jobPost = page.locator(`#jobcard-${i}`);


      const jobTitle = await getJobTitle(page, i)
      const companyName = await getCompanName(page, i)
      // Check valid Job Post
      if (!jobTitle?.includes('programmer') && !jobTitle?.includes('developer')) { 
        continue
      }
      // Check Duplicate
      await jobPost.click()
      await delay(500)
      const url = page.url()
      if (checkDuplicate(url)) {
        continue
      }

      const newRow = {
        A: sequence,
        B: jobTitle,
        C: companyName,
        D: url,
      };
      console.log(newRow);

      xlsx.utils.sheet_add_aoa(sheet, [Object.values(newRow)], { origin: sequence });
      xlsx.writeFile(workbook, 'Apply History.xlsx');
      sequence += 1

    }
    const data = JSON.stringify({ sequence })
    fs.writeFileSync('./sequence.json', data);
    pageNumber++;
    }
  } catch (error) {
    console.error('An error occurred:', error);
  } finally {
    await browser.close();
  }

})();



async function getJobTitle(page, index) {
  const jobTitle = await page.evaluate((index) => {
    const element = document.querySelector(`#jobcard-${index} a[data-automation=jobTitle]`);
    return element ? element.innerHTML : null;
  }, index);

  return jobTitle?.toLowerCase()
}

async function getCompanName(page, index) {
  const companyName = await page.evaluate((index) => {
    const element = document.querySelector(`#jobcard-${index} a[data-automation=jobCompany]`);
    return element ? element.innerHTML : null;
  }, index);
  return companyName?.toLowerCase()
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function checkDuplicate(url){
  for (let jobPost of jsonData){
    if (jobPost[2] === url){
      return true 
    }
  }
  return false
}