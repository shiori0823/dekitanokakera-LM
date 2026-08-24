const path = require('path');
const { chromium } = require('playwright-core');

(async () => {
  const executablePath = process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium';
  const browser = await chromium.launch({ executablePath });
  const page = await browser.newPage({ viewport: { width: 1080, height: 1350 } });
  const htmlPath = path.join(__dirname, '..', 'src', 'index.html');
  await page.goto('file://' + htmlPath);
  await page.evaluate(async () => { await document.fonts.ready; });
  await page.waitForTimeout(300);

  const outDir = path.join(__dirname, '..', 'dist');

  // full multi-page PDF
  await page.pdf({
    path: path.join(outDir, 'dekitanokakera-shindan.pdf'),
    width: '1080px',
    height: '1350px',
    printBackground: true,
    margin: { top: 0, bottom: 0, left: 0, right: 0 },
  });

  // per-page PNG previews for visual check
  const pages = await page.$$('.page');
  for (let i = 0; i < pages.length; i++) {
    await pages[i].screenshot({ path: path.join(outDir, `preview-page-${i + 1}.png`) });
  }

  await browser.close();
  console.log('done', pages.length, 'pages');
})();
