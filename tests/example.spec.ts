import { test, expect, request, APIRequestContext } from "@playwright/test";
import { join } from "path";
import {
  existsSync,
  readdirSync,
  statSync,
  readFileSync,
  writeFileSync,
  mkdirSync,
} from "fs";
const baseLocation = join(__dirname, "..", "CDN");
// Shared API context with auth headers
let apiContext: APIRequestContext;

test.beforeAll(async ({ playwright }) => {
  apiContext = await request.newContext({
    baseURL: "http://localhost:5000",
    extraHTTPHeaders: {
      Accept: "*/*",
    },
  });
});

test.afterAll(async () => {
  await apiContext.dispose();
});

const files = readdirSync(join(baseLocation, "generateExcel")).filter(
  (name: string) =>
    name.endsWith(".html") &&
    !name.includes("validator") && //this example include invalid schema so it's not download
    !name.includes("test"), //it's take time to load
);
const emptySheetName = [
  "ex0-0.xlsx", //no data
  "ex0.xlsx", // no data
  "ex21.xlsx", //
];
const largePayloadError = ["ex1-1.xlsx"];
files.forEach((filename: string) => {
  test(`generateExcel->${filename.substring(0, filename.lastIndexOf("."))} download file example`, async ({
    page,
  }) => {
    await page.waitForTimeout(1_000);
    await page.goto(`http://localhost/generateExcel/${filename}`);

    const downloadPromise = page.waitForEvent("download");
    // Expects page to have a heading with the name of Installation.
    {
      const downloadBtn = await page.getByTestId("generate-excel-btn");
      await expect(downloadBtn).toBeVisible();
      downloadBtn.click();
    }
    const download = await downloadPromise;
    const downloadFileName = download.suggestedFilename();
    expect(downloadFileName.endsWith(".xlsx")).toBeTruthy();

    // Save the file to a local path
    const savePath = join("downloads", downloadFileName);
    await download.saveAs(savePath);

    // ✅ Assert: file exists on disk
    expect(existsSync(savePath)).toBe(true);
    const stats = statSync(savePath);
    const fileSizeInBytes = stats.size;
    expect(fileSizeInBytes).toBeGreaterThan(1); //greater then 1B;
    if (largePayloadError.includes(downloadFileName)) {
      return;
    }

    const fileBuffer = readFileSync(savePath);
    const response = await apiContext.post("/convert", {
      multipart: {
        file: {
          name: downloadFileName,
          mimeType:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          buffer: fileBuffer,
        },
      },
    });
    if (emptySheetName.includes(downloadFileName)) {
      expect(response.status()).toBe(400);
      const body = await response.json();
      expect(body.error).toBe(
        "Sheet 'sheet1' is completely empty — nothing to render.",
      );
      const emptyResponse = await apiContext.post("/convert?empty=1", {
        multipart: {
          file: {
            name: downloadFileName,
            mimeType:
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            buffer: fileBuffer,
          },
        },
      });
      expect(emptyResponse.status()).toBe(200);
    } else {
      expect(response.status()).toBe(200);
      const body = await response.body();
      mkdirSync(join(__dirname, "test-results"), {
        recursive: true,
      });
      writeFileSync(
        join(
          __dirname,
          "test-results",
          downloadFileName.substring(0, downloadFileName.lastIndexOf(".")) +
            ".png",
        ),
        body,
      );
    }
    // expect(body).toHaveProperty("file_url");
    // expect(body.filename).toBe("sample.pdf");
    expect(fileSizeInBytes).toBeLessThan(10 * 1024 * 1024); //less then 10MB;
  });
});
