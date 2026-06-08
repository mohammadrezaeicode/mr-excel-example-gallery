import { test, expect, request, APIRequestContext } from "@playwright/test";
import { join } from "path";
// import { spawn, ChildProcessWithoutNullStreams } from "child_process";
import { existsSync, mkdirSync, writeFileSync } from "fs";

let expressApiContext: APIRequestContext;
let imageGeneratorApiContext: APIRequestContext;
// let serverProcess: ChildProcessWithoutNullStreams;

const reportFolder = join(__dirname, "test-results", "express");

// const waitForServer = async (proc: ChildProcessWithoutNullStreams) => {
//   return new Promise<void>((resolve, reject) => {
//     const timeout = setTimeout(() => {
//       reject(new Error("Express server did not start within 15 seconds."));
//     }, 15000);

//     const handleOutput = (chunk: Buffer) => {
//       const text = chunk.toString();
//       if (text.includes("run on 3000") || text.includes("listening")) {
//         clearTimeout(timeout);
//         resolve();
//       }
//     };

//     proc.stdout.on("data", handleOutput);
//     proc.stderr.on("data", handleOutput);
//     proc.on("exit", (code) => {
//       clearTimeout(timeout);
//       reject(new Error(`Express server exited early with code ${code}`));
//     });
//   });
// };

test.beforeAll(async ({ playwright }) => {
  //   serverProcess = spawn(process.execPath, ["./express/index.js"], {
  //     cwd: join(__dirname, ".."),
  //     env: { ...process.env },
  //     stdio: ["ignore", "pipe", "pipe"],
  //   });

  //   await waitForServer(serverProcess);

  expressApiContext = await request.newContext({
    baseURL: "http://localhost:4000",
    timeout: 120_000,
    extraHTTPHeaders: {
      Accept: "*/*",
    },
  });
  imageGeneratorApiContext = await request.newContext({
    baseURL: "http://localhost:5000",
    extraHTTPHeaders: {
      Accept: "*/*",
    },
  });
});

test.afterAll(async () => {
  await expressApiContext.dispose();
  //   if (serverProcess && !serverProcess.killed) {
  //     serverProcess.kill();
  //   }
});

test("Express endpoints /ex2/:id generate all examples and save Excel reports", async () => {
  mkdirSync(reportFolder, { recursive: true });

  const listResponse = await expressApiContext.get("/list");
  expect(listResponse.ok()).toBeTruthy();

  const examples = (await listResponse.json()) as string[];
  expect(Array.isArray(examples)).toBeTruthy();
  expect(examples.length).toBeGreaterThan(0);

  const ids = examples
    .filter(
      (name) =>
        typeof name === "string" &&
        name.startsWith("ex") &&
        name.toLowerCase().indexOf("dash") < 0,
    )
    .map((name) => name.replace(/^ex/, ""));

  for (const id of ids) {
    const responses = [{ path: `/ex2/${id}`, file: `ex2-${id}.xlsx` }];

    for (const item of responses) {
      console.log(item.path);

      const res = await expressApiContext.get(item.path);
      expect(res.ok()).toBeTruthy();
      const body = await res.body();
      expect(body.length).toBeGreaterThan(16);

      const filePath = join(reportFolder, item.file);
      writeFileSync(filePath, body);
      expect(existsSync(filePath)).toBeTruthy();
      const imageResponse = await imageGeneratorApiContext.post("/convert?empty=1", {
        multipart: {
          file: {
            name: item.file,
            mimeType:
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            buffer: body,
          },
        },
      });
      expect(imageResponse.status()).toBe(200);
      const imageBody = await imageResponse.body();
      mkdirSync(join(__dirname, "test-results","express"), {
        recursive: true,
      });
      writeFileSync(
        join(
          __dirname,
          "test-results","express",
          item.file.substring(0, item.file.lastIndexOf(".")) +
            ".png",
        ),
        imageBody,
      );

    }
  }
});
