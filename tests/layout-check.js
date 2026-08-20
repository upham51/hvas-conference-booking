// Layout regression check for the parent booking page.
//
// This exists because of a real bug: the code card is pulled up over the
// bottom of the hero photo with a negative margin, and the hero's bottom
// padding is what keeps the date banner clear of it. Those two numbers were
// set by hand at three different breakpoints, drifted apart, and on any
// desktop window shorter than about 830px the card landed on top of the
// banner and hid the conference times. It looked fine on a maximised window,
// so it shipped.
//
// The CSS now derives the padding from --card-lift and --hero-gap so the
// arithmetic cannot drift again. This check is the backstop: it opens the
// real page at the window sizes people actually use and fails the build if
// anything overlaps or the page scrolls sideways.
//
// Run: node tests/layout-check.js

const path = require('path');
const { chromium } = require('playwright');

const PAGE_URL = 'file://' + path.join(__dirname, '..', 'index.html');

// Maximised displays, plus the shorter windows you get once a browser has a
// tab strip, an address bar and a bookmarks bar eating into the viewport.
const VIEWPORTS = [
  { w: 1920, h: 1080, note: 'large desktop' },
  { w: 1728, h: 1117, note: '16in laptop' },
  { w: 1512, h: 982, note: '14in laptop' },
  { w: 1536, h: 864, note: 'common windows laptop' },
  { w: 1440, h: 900, note: 'macbook, maximised' },
  { w: 1280, h: 800, note: 'macbook air' },
  { w: 1440, h: 760, note: 'macbook with bookmarks bar' },
  { w: 1370, h: 700, note: 'windowed browser' },
  { w: 1280, h: 660, note: 'short windowed browser' },
  { w: 1024, h: 640, note: 'small laptop' },
  { w: 768, h: 1024, note: 'tablet' },
  { w: 390, h: 844, note: 'phone' },
];

// Minimum daylight required between the date banner and the card on top of it.
const MIN_BANNER_GAP = 12;

async function measure(page) {
  return page.evaluate(() => {
    const box = (sel) => {
      const el = document.querySelector(sel);
      return el ? el.getBoundingClientRect() : null;
    };
    const banner = box('.date-banner');
    const card = box('.hero-overlap');
    const brand = box('.brand');
    const h1 = box('.hero h1');
    const doc = document.documentElement;
    return {
      bannerToCard: Math.round(card.top - banner.bottom),
      headlineToBanner: Math.round(banner.top - h1.bottom),
      brandToHeadline: Math.round(h1.top - brand.bottom),
      horizontalOverflow: doc.scrollWidth - doc.clientWidth,
      verticalOverflow: doc.scrollHeight - doc.clientHeight,
    };
  });
}

function problemsFor(m) {
  const problems = [];
  if (m.bannerToCard < MIN_BANNER_GAP) {
    problems.push(`card sits ${MIN_BANNER_GAP - m.bannerToCard}px into the date banner (gap ${m.bannerToCard}px)`);
  }
  if (m.headlineToBanner < 0) problems.push(`date banner overlaps the headline by ${-m.headlineToBanner}px`);
  if (m.brandToHeadline < 0) problems.push(`headline overlaps the school logo by ${-m.brandToHeadline}px`);
  if (m.horizontalOverflow > 0) problems.push(`page scrolls sideways by ${m.horizontalOverflow}px`);
  return problems;
}

(async () => {
  // CHROMIUM_PATH lets this run against a preinstalled browser (some sandboxes
  // ship one); CI installs its own and leaves it unset.
  const browser = await chromium.launch(
    process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {}
  );
  const page = await browser.newPage();
  await page.goto(PAGE_URL, { waitUntil: 'load' });
  await page.waitForTimeout(700); // let the webfonts settle before measuring

  const failures = [];
  for (const vp of VIEWPORTS) {
    await page.setViewportSize({ width: vp.w, height: vp.h });
    await page.waitForTimeout(260);

    const m = await measure(page);
    const problems = problemsFor(m);
    const label = `${vp.w}x${vp.h}`.padEnd(10);
    const fits = m.verticalOverflow === 0 ? 'one page' : `scrolls ${m.verticalOverflow}px`;

    if (problems.length) {
      failures.push({ vp, problems });
      console.log(`FAIL  ${label} ${vp.note.padEnd(26)} ${problems.join('; ')}`);
    } else {
      console.log(`ok    ${label} ${vp.note.padEnd(26)} banner gap ${String(m.bannerToCard).padStart(3)}px, ${fits}`);
    }
  }

  await browser.close();

  if (failures.length) {
    console.error(`\n${failures.length} of ${VIEWPORTS.length} viewports have overlapping layout.`);
    process.exit(1);
  }
  console.log(`\nAll ${VIEWPORTS.length} viewports clear.`);
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
