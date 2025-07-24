import { Browser, chromium, Page, ElementHandle } from "playwright";
import axios from "axios";
// import { google } from "googleapis";
import * as XLSX from "xlsx";
import * as fs from "fs";
import * as path from "path";

const adspowerApiUrl = "http://local.adspower.net:50325";

const profileId = "kuunq3j";
const date = "202507";

// Excel export configuration
const EXCEL_OUTPUT_DIR = "./exports";



async function createExcelFile(data: any[], adStats: any[] = []) {
  try {
    // Create exports directory if it doesn't exist
    if (!fs.existsSync(EXCEL_OUTPUT_DIR)) {
      fs.mkdirSync(EXCEL_OUTPUT_DIR, { recursive: true });
    }

    // Create a new workbook
    const workbook = XLSX.utils.book_new();

    // Prepare data for each worksheet
    const tabNames: string[] = [];

    for (const ad of data) {
      const adId = ad.tds?.[0]?.text || "";
      const campaignName = ad.campaignName || ad.text || "";

      // Skip if ad ID is empty or campaign name is empty/null/Unknown
      // But allow pr-cell elements without hrefs (they might not have campaign names)
      const hasPrCellWithoutHref = ad.tds.some(
        (td) => td.isPrCell && !td.hasHref && td.text.trim()
      );

      if (
        !adId ||
        (!campaignName && !hasPrCellWithoutHref) ||
        (campaignName &&
          (campaignName === "Unknown" || campaignName.trim() === ""))
      ) {
        continue;
      }

      // Create tab for ALL ad links, even if they don't have href (no detailed stats)
      // This ensures we have a complete record of all campaigns

      // Create tab name using AD ID + campaign title
      let tabName = adId || "Unknown";
      const campaignTitle =
        ad.tds?.[1]?.text || ad.text || ad.campaignName || "";

      // Combine AD ID and campaign title, but keep it within Excel's 31 character limit
      if (campaignTitle && campaignTitle.trim() && campaignTitle !== tabName) {
        tabName = `${adId}-${campaignTitle}`
          .replace(/[\\\/\?\*\[\]:]/g, "_") // Replace invalid Excel characters with underscores
          .replace(/\s+/g, "_") // Replace spaces with underscores
          .substring(0, 31);
      } else {
        // If no campaign title or it's the same as tabName, just use the AD ID
        tabName = `${adId}`
          .replace(/[\\\/\?\*\[\]:]/g, "_") // Replace invalid Excel characters with underscores
          .substring(0, 31);
      }

      tabNames.push(tabName);

      // Debug: Show the tab name creation
      console.log(
        `📋 Created tab name: "${tabName}" (AD ID: "${adId}", Campaign: "${campaignTitle}")`
      );

      // Prepare the data array: [id, name, views, amount]
      const rowData = [
        ad.tds?.[0]?.text || ad.campaignName || "", // ID or campaign name
        ad.tds?.[1]?.text || ad.text || "", // Name or text
        ad.tds?.[2]?.text || ad.views || "", // Views
        ad.tds?.[3]?.text || ad.amount || "", // Amount
      ];

      // Debug: Log the data structure for this tab
      console.log(`Tab ${tabName} data:`, {
        adId,
        campaignName: ad.campaignName,
        text: ad.text,
        tds: ad.tds?.map((td) => ({
          text: td.text,
          hasHref: td.hasHref,
          isPrCell: td.isPrCell,
        })),
        rowData,
      });

      // Find corresponding detailed stats from adStats by AD ID
      let detailedData: string[][] = [];
      let matchingStats: {
        campaignName?: string;
        href: string | null;
        originalIndex?: number;
        rows: Array<{
          date: string;
          views: string;
          amount: string;
        }>;
      } | null = null;

      // First try to match by exact AD ID from href
      if (ad.href) {
        const hrefAdId = ad.href.match(/\/account\/ad\/(\d+)/)?.[1];
        if (hrefAdId) {
          matchingStats = adStats.find((stat) => {
            if (stat.href) {
              const statHrefAdId = stat.href.match(/\/account\/ad\/(\d+)/)?.[1];
              return statHrefAdId === hrefAdId;
            }
            return false;
          });
        }
      }

      // If no match found, try to match by campaign name as fallback
      if (!matchingStats && ad.text) {
        matchingStats = adStats.find(
          (stat) =>
            stat.campaignName && stat.campaignName.trim() === ad.text.trim()
        );
      }

      // If still no match, try to match by original index (third fallback)
      // But only if this ad actually has an href (to prevent incorrect matching)
      if (!matchingStats && ad.href) {
        const currentIndex = data.indexOf(ad);
        matchingStats = adStats.find(
          (stat) => stat.originalIndex === currentIndex
        );
      }

      // Log the mapping for debugging
      if (matchingStats) {
        console.log(
          `✓ Mapped ${ad.text} (ID: ${adId}, href: ${ad.href}) to stats with ${
            matchingStats.rows?.length || 0
          } rows`
        );
      } else {
        console.log(
          `✗ No matching stats found for ${ad.text} (ID: ${adId}, href: ${ad.href})`
        );
        if (ad.href) {
          console.log(
            `Available stats:`,
            adStats.map((s) => ({
              name: s.campaignName,
              href: s.href,
              originalIndex: s.originalIndex,
              rows: s.rows?.length || 0,
            }))
          );
        }
      }

      // Only add detailed stats if this ad has an href and we found matching stats
      if (
        ad.href &&
        matchingStats &&
        matchingStats.rows &&
        matchingStats.rows.length > 0
      ) {
        detailedData = matchingStats.rows.map((row) => [
          row.date || "",
          row.views || "",
          row.amount || "",
        ]);
      } else if (!ad.href) {
        // Explicitly log that this ad has no href, so no detailed stats
        console.log(`⚠️ Ad ${adId} has no href, skipping detailed stats`);
      }

      // Combine campaign data and detailed stats
      const allData = [rowData, [], ...detailedData]; // Empty row for spacing

      // Create worksheet
      const worksheet = XLSX.utils.aoa_to_sheet(allData);
      XLSX.utils.book_append_sheet(workbook, worksheet, tabName);

      const isPrCellWithoutHref = hasPrCellWithoutHref
        ? " (pr-cell without href)"
        : "";
      console.log(
        `Added data to worksheet: ${tabName}${isPrCellWithoutHref}, ${rowData} + ${detailedData.length} detailed rows`
      );

      // For now, let's not create additional tabs for pr-cell elements without hrefs
      // This will help us debug the main tab structure first
      if (hasPrCellWithoutHref) {
        console.log(
          `Found pr-cell without href in ${tabName}:`,
          ad.tds
            .filter((td) => td.isPrCell && !td.hasHref && td.text.trim())
            .map((td) => td.text)
        );
      }
    }

    // Handle total row if it exists
    for (let i = 0; i < data.length; i++) {
      const ad = data[i];
      if (ad.tds && ad.tds.length === 3) {
        // This is the total row, create a "Total" worksheet
        const totalData = [
          ad.tds[0]?.text || "",
          ad.tds[1]?.text || "",
          ad.tds[2]?.text || "",
        ];

        const totalWorksheet = XLSX.utils.aoa_to_sheet([totalData]);
        XLSX.utils.book_append_sheet(workbook, totalWorksheet, "Total");
        console.log(`Added Total worksheet: ${totalData}`);
        break;
      }
    }

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().split("T")[0];
    const filename = `Telegram_Ads_Stats_${timestamp}.xlsx`;
    const filepath = path.join(EXCEL_OUTPUT_DIR, filename);

    // Write the workbook to file
    XLSX.writeFile(workbook, filepath);

    console.log(`Excel file created: ${filepath}`);
    return filepath;
  } catch (error) {
    console.error("Error creating Excel file:", error);
    throw error;
  }
}

async function checkTelegramStat(profileId: string) {
  let page: Page | null = null;
  let browser: Browser | null = null;
  try {
    //step 1: ======================open fucking browser===============

    console.log(`Starting browser for profile ${profileId}...`);
    const response = await axios.get(
      `${adspowerApiUrl}/api/v1/browser/start?user_id=${profileId}&headless=1`
      // `${adspowerApiUrl}/api/v1/browser/start?user_id=${profileId}`
    );
    const wsEndpoint = response.data.data.ws.puppeteer;
    browser = await chromium.connectOverCDP(wsEndpoint, {
      timeout: 30000,
    });

    // Get existing context or create new one
    const contexts = browser.contexts();
    const context =
      contexts.length > 0 ? contexts[0] : await browser.newContext();

    // Add script to hide webdriver if context is new
    if (contexts.length === 0) {
      await context.addInitScript(() => {
        Object.defineProperty(navigator, "webdriver", { get: () => false });
      });
    }

    console.log("Creating new tab...");
    page = await context.newPage();
    await page.setViewportSize({ width: 1920, height: 1080 });
    const size = await page.evaluate(() => {
      return {
        width: window.innerWidth,
        height: window.innerHeight,
      };
    });
    console.log("🖥️ Browser inner size:", size);
    await page.setExtraHTTPHeaders({
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    });

    //step 2: ======================open telegram dashboard===============
    console.log("Opening telegram dashboard...");
    await page.goto(
      `https://ads.telegram.org/account/stats?month=${date}#report`,
      {
        timeout: 120000,
      }
    );

    // Scroll down to load more content
    console.log("Scrolling down to load more content...");
    await page.evaluate(() => {
      window.scrollTo(0, document.body.scrollHeight);
    });

    // Wait a bit for content to load after scrolling
    await page.waitForTimeout(2000);

    const data = await page.evaluate(() => {
      const rows = document.querySelectorAll(
        ".table.pr-table.pr-table-sticky.pr-table-vtop tr"
      );

      // Extract all td elements from each row
      const adLinks: Array<{
        href: string | null;
        text: string | undefined;
        tds: Array<{
          text: string;
          hasHref: boolean;
          href?: string;
          isPrCell: boolean;
        }>;
      }> = [];

      for (let i = 1; i < rows.length; i++) {
        // Start from index 1 to skip header row
        const row = rows[i];
        const cells = row.querySelectorAll("td");

        if (cells.length === 0) continue;

        // Check if this row has any anchor tags (hrefs)
        const anchors = row.querySelectorAll("a");
        let href: string | null = null;
        let text = "";

        // Only set href and text if this row actually has anchor tags
        if (anchors.length > 0) {
          // Get the first anchor tag for href (campaign name)
          const firstAnchor = anchors[0];
          href = firstAnchor.getAttribute("href");
          text = firstAnchor.textContent?.trim() || "";
        }

        // Extract all td content
        const tds = Array.from(cells).map((cell) => {
          let cellText = cell.textContent?.trim() || "";
          // Remove diamond icon (💎) from the text
          cellText = cellText.replace(/💎/g, "");
          const cellAnchor = cell.querySelector("a");
          const hasHref = !!cellAnchor;
          const cellHref = cellAnchor?.getAttribute("href") || undefined;

          // Check if this cell has a pr-cell div (with or without href)
          const prCell = cell.querySelector(".pr-cell");
          const isPrCell = !!prCell;

          return {
            text: cellText,
            hasHref,
            href: cellHref,
            isPrCell,
          };
        });

        adLinks.push({
          href,
          text,
          tds,
        });

        // Debug: Log the href assignment for this row
        console.log(
          `Row ${i}: href=${href}, text="${text}", tds=${tds
            .map((td) => td.text)
            .join(", ")}`
        );
      }

      return { adLinks };
    });

    console.log(
      "Found ad links:",
      data.adLinks,
      data.adLinks.map((ad) => ad.tds.map((td) => td.text))
    );

    // Log information about pr-cell elements without hrefs
    const prCellsWithoutHrefList = data.adLinks.flatMap((ad) =>
      ad.tds.filter((td) => td.isPrCell && !td.hasHref && td.text.trim())
    );
    console.log(
      "Found pr-cell elements without hrefs:",
      prCellsWithoutHrefList.map((td) => td.text)
    );

    // Process each ad link to get detailed stats
    const adStats: Array<{
      campaignName?: string;
      href: string | null;
      originalIndex?: number;
      rows: Array<{
        date: string;
        views: string;
        amount: string;
      }>;
    }> = [];

    // Process links in parallel using up to 10 tabs
    const maxTabs = 10;
    const batchSize = Math.min(maxTabs, data.adLinks.length);

    for (let i = 0; i < data.adLinks.length; i += batchSize) {
      const batch = data.adLinks.slice(i, i + batchSize);
      console.log(
        `Processing batch ${Math.floor(i / batchSize) + 1}: ${batch.length} ads`
      );

      // Create tabs for this batch - only process links that have hrefs
      const validLinks = batch.filter(
        (adLink) => adLink.href && adLink.href !== "null"
      );
      console.log(
        `Batch has ${batch.length} total links, ${validLinks.length} have hrefs for detailed stats`
      );

      const tabs = await Promise.all(
        validLinks.map(async (adLink, index) => {
          const newPage = await context.newPage();
          await newPage.setViewportSize({ width: 1920, height: 1080 });
          await newPage.setExtraHTTPHeaders({
            "User-Agent":
              "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
          });

          console.log(`Tab ${index + 1}: Processing ${adLink.text}`);

          try {
            // Navigate to the ad's stats page
            const fullUrl = `https://ads.telegram.org${adLink.href}`;
            console.log(`Tab ${index + 1}: Navigating to ${fullUrl}`);

            await newPage.goto(fullUrl, { timeout: 60000 });
            await newPage.waitForLoadState("networkidle");

            // Wait for the table to load
            await newPage.waitForSelector(
              ".table.pr-table.pr-table-sticky.pr-table-vtop",
              { timeout: 30000 }
            );

            // Extract stats data from all rows
            const statsData = await newPage.evaluate(() => {
              const rows = document.querySelectorAll(
                ".table.pr-table.pr-table-sticky.pr-table-vtop tbody tr"
              );

              if (rows.length === 0) return null;

              const allRowsData: Array<{
                date: string;
                views: string;
                amount: string;
              }> = [];

              for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const cells = row.querySelectorAll("td");

                if (cells.length < 3) continue;

                // Extract date from first cell
                const dateCell = cells[0].querySelector(".pr-cell");
                const date = dateCell ? dateCell.textContent?.trim() || "" : "";

                // Extract views from second cell
                const viewsCell = cells[1].querySelector(".pr-cell");
                const views = viewsCell
                  ? viewsCell.textContent?.trim() || ""
                  : "";

                // Extract amount from third cell
                const amountCell = cells[2].querySelector(".pr-cell");
                let amount = "";
                if (amountCell) {
                  // Get all text content and extract the numeric amount
                  const amountText = amountCell.textContent?.trim() || "";
                  // Look for pattern like "6.90" or similar
                  const amountMatch = amountText.match(/(\d+\.?\d*)/);
                  amount = amountMatch ? amountMatch[1] : amountText;
                }

                allRowsData.push({ date, views, amount });
              }

              return allRowsData;
            });

            await newPage.close();

            if (statsData && statsData.length > 0) {
              return {
                campaignName: adLink.text,
                href: adLink.href,
                originalIndex: i + index, // Add original index for better tracking
                rows: statsData,
              };
            } else {
              console.log(
                `Tab ${index + 1}: No stats data found for ${adLink.text}`
              );
              return null;
            }
          } catch (error) {
            console.error(
              `Tab ${index + 1}: Error processing ad ${adLink.text}:`,
              error
            );
            await newPage.close();
            return null;
          }
        })
      );

      // Add successful results to adStats
      const validResults = tabs.filter((result) => result !== null);
      adStats.push(...validResults);

      console.log(
        `Batch completed: ${validResults.length}/${batch.length} successful`
      );

      // Small delay between batches
      await page.waitForTimeout(2000);
    }

    console.log(
      "Final ad stats:",
      adStats.map((ad) => ad.campaignName),
      adStats.map((ad) => ad.rows)
    );

    // Log summary of what will be created
    const totalAdLinks = data.adLinks.length;
    const linksWithHrefs = data.adLinks.filter(
      (ad) => ad.href && ad.href !== "null"
    ).length;
    const linksWithStats = adStats.length;
    const prCellsWithoutHref = data.adLinks.flatMap((ad) =>
      ad.tds.filter((td) => td.isPrCell && !td.hasHref && td.text.trim())
    ).length;

    console.log(`📊 Summary: ${totalAdLinks} total ad links found`);
    console.log(
      `📊 ${linksWithHrefs} links have hrefs (can get detailed stats)`
    );
    console.log(`📊 ${linksWithStats} links successfully got detailed stats`);
    console.log(
      `📊 ${prCellsWithoutHref} pr-cell elements without hrefs found`
    );
    console.log(`📊 Will create ${totalAdLinks} tabs in Excel (all campaigns)`);

    // Debug: Show the structure of the first few ad links
    console.log(
      "📊 First 3 ad links structure:",
      data.adLinks.slice(0, 3).map((ad) => ({
        href: ad.href,
        text: ad.text,
        tdsCount: ad.tds?.length,
        tdsTexts: ad.tds?.map((td) => td.text),
        tdsWithHref: ad.tds?.filter((td) => td.hasHref).length,
        tdsPrCell: ad.tds?.filter((td) => td.isPrCell).length,
      }))
    );

    // Create Excel file with the collected data
    if (data.adLinks.length > 0) {
      console.log("Creating Excel file...");
      try {
        const filepath = await createExcelFile(data.adLinks, adStats);
        console.log("Excel file created successfully:", filepath);
      } catch (error) {
        console.error("Error creating Excel file:", error);
      }
    } else {
      console.log("No data to create Excel file");
    }

    await page.close();
    await browser.close();
  } catch (error) {
    console.error(`Error checking Telegram stat: ${error}`);
  } finally {
    if (page) {
      await page.close();
    }
    if (browser) {
      await browser.close();
    }
    const stopBrowserResponse = await axios.get(
      `${adspowerApiUrl}/api/v1/browser/stop?user_id=${profileId}`
    );
  }
}

checkTelegramStat(profileId);
